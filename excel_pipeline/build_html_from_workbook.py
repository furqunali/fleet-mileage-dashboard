"""
build_html_from_workbook.py
----------------------------
Single-source-of-truth generator. The Excel workbook is authoritative; the
standalone HTML calculator is a GENERATED artifact. This script reads the
workbook's Directory + Matrix tabs and rebuilds the `const DATA = {...}` block
embedded in the HTML, so the two can never silently drift.

Usage:
  python build_html_from_workbook.py            # generate the HTML
  python build_html_from_workbook.py --verify   # check parity, write nothing

Data mapping (workbook -> HTML DATA):
  matrix[from][to]        <- Matrix!B2:AD30 (rows A2:A30, cols B1:AD1)
  sites[code].name        <- Directory col C
  sites[code].address     <- Directory col D  (U+00A0 normalised to space)
  sites[code].type        <- Directory col E  ("Head Office" -> "C-Store" + is_office)
  sites[code].office_miles<- DERIVED from matrix[hq][code]  (not a stored copy)
  office_code             <- Directory!J1 (hq_code)

Input template : Fleet_Distance_Calculator_CORRECTED (1).html
Workbook source: Fleet_Distance_Dashboard_v2.xlsx
Output         : Fleet_Distance_Calculator_Generate.html
"""
import os
import re
import sys
import json

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WORKBOOK = os.path.join(HERE, "Fleet_Distance_Dashboard_v7.xlsx")
TEMPLATE = os.path.join(HERE, "Fleet_Distance_Calculator_template.html")
OUTPUT = os.path.join(HERE, "Fleet_Distance_Calculator_Generate.html")

DATA_RE = re.compile(r"const DATA = \{.*?\}\s*;\s*</script>", re.S)


def clean(text):
    """Normalise non-breaking spaces and trim."""
    if text is None:
        return ""
    return str(text).replace(" ", " ").strip()


def read_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    # --- Matrix (row = FROM, col = TO) ---
    # Dynamic extent: scan header row / first column until blank so the generator
    # auto-adapts when locations are added (e.g. v6 added the two home rows/cols).
    mx = wb["Matrix"]
    col_codes = []
    c = 2
    while clean(mx.cell(1, c).value):
        col_codes.append(clean(mx.cell(1, c).value))
        c += 1
    matrix = {}
    r = 2
    while clean(mx.cell(r, 1).value):                                # until spacer/footnote
        row_code = clean(mx.cell(r, 1).value)
        matrix[row_code] = {}
        for j, cc in enumerate(col_codes):
            matrix[row_code][cc] = mx.cell(r, 2 + j).value
        r += 1

    # --- Directory ---
    dr = wb["Directory"]
    hq_code = clean(dr["J1"].value)
    sites = {}
    r = 5
    while True:                                                      # until first blank code
        code = clean(dr.cell(r, 2).value)   # B
        if not code:
            break
        name = clean(dr.cell(r, 3).value)   # C
        address = clean(dr.cell(r, 4).value)  # D
        dtype = clean(dr.cell(r, 5).value)  # E
        lat = dr.cell(r, 11).value          # K
        lon = dr.cell(r, 12).value          # L
        is_office = (code == hq_code) or (dtype == "Head Office")
        html_type = "C-Store" if is_office else dtype
        office_miles = matrix.get(hq_code, {}).get(code)
        site = {
            "name": name,
            "address": address,
            "office_miles": office_miles,
            "type": html_type,
            "lat": float(lat) if isinstance(lat, (int, float)) else None,
            "lon": float(lon) if isinstance(lon, (int, float)) else None,
        }
        if is_office:
            site["is_office"] = True
        sites[code] = site
        r += 1

    wb.close()
    return {"sites": sites, "matrix": matrix, "office_code": hq_code}


def extract_html_data(html):
    m = DATA_RE.search(html)
    if not m:
        raise SystemExit("ERROR: could not locate the `const DATA = {...};` block in the template.")
    block = m.group(0)
    json_text = block[len("const DATA = "):block.rindex(";")].strip()
    return json.loads(json_text)


def build():
    data = read_workbook(WORKBOOK)
    with open(TEMPLATE, encoding="utf-8") as f:
        html = f.read()

    payload = json.dumps(data, ensure_ascii=False)
    replacement = "const DATA = " + payload + "\n;\n</script>"
    new_html, n = DATA_RE.subn(lambda _m: replacement, html, count=1)
    if n != 1:
        raise SystemExit("ERROR: could not substitute the DATA block (matched %d times)." % n)

    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(new_html)
    missing = [c for c, s in data["sites"].items() if s.get("lat") is None or s.get("lon") is None]
    print("Wrote", os.path.basename(OUTPUT))
    print("  sites:", len(data["sites"]), "| matrix rows:", len(data["matrix"]),
          "| office:", data["office_code"])
    print("  sites with coordinates:", len(data["sites"]) - len(missing), "/", len(data["sites"]))
    if missing:
        print("  WARNING: sites missing coordinates:", missing)
    return data


def verify():
    """Compare the generated HTML's DATA against the workbook. Returns True if identical."""
    wb_data = read_workbook(WORKBOOK)
    with open(OUTPUT, encoding="utf-8") as f:
        html_data = extract_html_data(f.read())

    problems = []

    # office code
    if wb_data["office_code"] != html_data.get("office_code"):
        problems.append("office_code: wb=%r html=%r" % (wb_data["office_code"], html_data.get("office_code")))

    # sites
    wb_sites, html_sites = wb_data["sites"], html_data["sites"]
    if set(wb_sites) != set(html_sites):
        problems.append("site codes differ: only_wb=%s only_html=%s"
                        % (set(wb_sites) - set(html_sites), set(html_sites) - set(wb_sites)))
    for code in set(wb_sites) & set(html_sites):
        for field in ("name", "address", "type", "office_miles", "lat", "lon"):
            a, b = wb_sites[code].get(field), html_sites[code].get(field)
            if a != b:
                problems.append("site %s.%s: wb=%r html=%r" % (code, field, a, b))
        if wb_sites[code].get("is_office") != html_sites[code].get("is_office"):
            problems.append("site %s.is_office: wb=%r html=%r"
                            % (code, wb_sites[code].get("is_office"), html_sites[code].get("is_office")))

    # matrix
    wb_mx, html_mx = wb_data["matrix"], html_data["matrix"]
    mismatches = 0
    for r in wb_mx:
        for c in wb_mx[r]:
            if wb_mx[r][c] != html_mx.get(r, {}).get(c):
                mismatches += 1
                if mismatches <= 20:
                    problems.append("matrix[%s][%s]: wb=%r html=%r"
                                    % (r, c, wb_mx[r][c], html_mx.get(r, {}).get(c)))

    # explicit spot checks the user cares about
    e27 = wb_mx.get("0077", {}).get("0017")
    print("Spot check  matrix[0077][0017] (Matrix!E27): workbook=%r  html=%r"
          % (e27, html_mx.get("0077", {}).get("0017")))
    print("Spot check  matrix[0017][0077] (symmetric)  : workbook=%r  html=%r"
          % (wb_mx.get("0017", {}).get("0077"), html_mx.get("0017", {}).get("0077")))

    # non-breaking space check across all generated site text
    nbsp_hits = [code for code, s in html_sites.items()
                 if " " in s.get("address", "") or " " in s.get("name", "")]
    print("NBSP check  addresses/names containing U+00A0 in HTML:", nbsp_hits or "none")

    # coordinate coverage
    no_coords = [c for c, s in html_sites.items() if s.get("lat") is None or s.get("lon") is None]
    print("Coord check sites with numeric lat/lon: %d/%d %s"
          % (len(html_sites) - len(no_coords), len(html_sites),
             ("(missing: %s)" % no_coords) if no_coords else ""))
    office = html_sites.get(html_data.get("office_code"), {})
    print("Coord check office has coordinates:", office.get("lat") is not None and office.get("lon") is not None)

    if problems:
        print("\nVERIFY: FAIL —", len(problems), "problem(s):")
        for p in problems:
            print("  -", p)
        return False
    print("\nVERIFY: PASS — generated HTML matches the workbook exactly (%d sites, %d matrix cells)."
          % (len(wb_sites), sum(len(v) for v in wb_mx.values())))
    return True


if __name__ == "__main__":
    if "--verify" in sys.argv:
        ok = verify()
        sys.exit(0 if ok else 1)
    else:
        build()
