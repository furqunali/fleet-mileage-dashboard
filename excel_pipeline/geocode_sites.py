"""
geocode_sites.py
----------------
ONE-TIME build step. Geocodes the 29 Fleet site addresses once (via Nominatim),
validates the results against the trusted Matrix distances, and writes the
coordinates into NEW columns K (LAT) and L (LON) of the Directory tab.

The Directory tab is the single master for coordinates. The Map of Site tab
(static picture + slicer + Power Pivot data model) is left completely intact.
The HTML never geocodes at runtime — it consumes these stored coordinates.

Method:
  * reads addresses from Directory (v2 workbook)
  * geocodes each UNIQUE address (cached in geocode_cache.json, 1.1s/request,
    proper User-Agent per Nominatim policy); co-located sites share coordinates
    automatically because they share an address string
  * optional manual overrides via coord_overrides.json  {code: [lat, lon]}
  * validates: Texas bounding box + haversine(office,site) vs recorded HQ miles
  * writes coordinates into a NEW versioned workbook via surgical XML editing so
    the slicer / data model / pivot / image / formulas are preserved

Input : Fleet_Distance_Dashboard_v2.xlsx
Output: Fleet_Distance_Dashboard_v3.xlsx
"""
import json
import math
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import zipfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Fleet_Distance_Dashboard_v2.xlsx")
OUT = os.path.join(HERE, "Fleet_Distance_Dashboard_v3.xlsx")
CACHE = os.path.join(HERE, "geocode_cache.json")
OVERRIDES = os.path.join(HERE, "coord_overrides.json")

DIRECTORY_SHEET = "xl/worksheets/sheet3.xml"
USER_AGENT = os.getenv("GEOCODER_USER_AGENT",
                       "fleet-distance-dashboard/1.0 (contact: geocoder@example.com)")

# Texas-ish bounding box for a sanity check
TX = dict(lat_min=25.5, lat_max=36.8, lon_min=-107.0, lon_max=-93.0)


# ----------------------------- read the workbook -----------------------------
def read_directory_and_matrix():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    dr = wb["Directory"]
    hq = str(dr["J1"].value).strip()
    rows = []
    for r in range(5, 34):
        code = dr.cell(r, 2).value
        if code is None:
            continue
        rows.append(dict(row=r, code=str(code).strip(),
                         name=str(dr.cell(r, 3).value).strip(),
                         address=str(dr.cell(r, 4).value).strip(),
                         type=str(dr.cell(r, 5).value).strip()))
    mx = wb["Matrix"]
    col_codes = [str(mx.cell(1, c).value).strip() for c in range(2, 31)]
    office_miles = {}
    for rr in range(2, 31):
        if str(mx.cell(rr, 1).value).strip() == hq:
            for c in range(2, 31):
                office_miles[col_codes[c - 2]] = mx.cell(rr, c).value
    wb.close()
    return hq, rows, office_miles


# ----------------------------- geocoding -------------------------------------
def load_json(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def nominatim(query):
    url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(
        {"format": "json", "limit": 1, "countrycodes": "us", "q": query})
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as resp:
        data = json.load(resp)
    if data:
        return [float(data[0]["lat"]), float(data[0]["lon"])]
    return None


def simplify(address):
    """Fallback query: drop suite/unit tokens that confuse the geocoder."""
    a = re.sub(r"\bSuit(e)?\b\.?\s*\d+", "", address, flags=re.I)
    a = re.sub(r"\s{2,}", " ", a).strip(" ,")
    return a


def zip_of(address):
    m = re.findall(r"\b(\d{5})\b", address)
    return m[-1] if m else None


def geocode_address(address, cache):
    """Returns (coords, method). Tries full -> simplified -> ZIP centroid."""
    if address in cache:
        return cache[address], "cache"
    # full
    coords = nominatim(address)
    method = "full"
    time.sleep(1.1)
    if coords is None:
        s = simplify(address)
        if s != address:
            coords = nominatim(s)
            method = "simplified"
            time.sleep(1.1)
    if coords is None:
        z = zip_of(address)
        if z:
            coords = nominatim(z + ", USA")
            method = "zip"
            time.sleep(1.1)
    if coords is not None:
        cache[address] = coords
        with open(CACHE, "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2)
    return coords, method


def haversine_mi(a, b):
    R = 3958.7613
    lat1, lon1, lat2, lon2 = map(math.radians, [a[0], a[1], b[0], b[1]])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 2 * R * math.asin(math.sqrt(h))


# ----------------------------- XML writing -----------------------------------
def add_cells_to_row(xml, r, extra):
    pat = re.compile(r'(<row r="%d"[^>]*?)(>)(.*?)(</row>)' % r, re.S)
    def repl(m):
        attrs = re.sub(r'spans="1:10"', 'spans="1:12"', m.group(1))
        return attrs + m.group(2) + m.group(3) + extra + m.group(4)
    new, n = pat.subn(repl, xml, count=1)
    if n != 1:
        raise SystemExit("ERROR: Directory row %d not found for cell insert" % r)
    return new


def write_v3(hq, rows, coords_by_code):
    zin = zipfile.ZipFile(SRC, "r")
    changed = {}

    # --- Directory sheet: add K/L header + data cells ---
    xml = zin.read(DIRECTORY_SHEET).decode("utf-8")
    header = ('<c r="K4" s="28" t="inlineStr"><is><t>LAT</t></is></c>'
              '<c r="L4" s="28" t="inlineStr"><is><t>LON</t></is></c>')
    xml = add_cells_to_row(xml, 4, header)
    for item in rows:
        r = item["row"]
        c = coords_by_code.get(item["code"])
        if c is None:
            continue
        cells = ('<c r="K%d" s="2"><v>%.7f</v></c>'
                 '<c r="L%d" s="2"><v>%.7f</v></c>' % (r, c[0], r, c[1]))
        xml = add_cells_to_row(xml, r, cells)
    # dimension + cols
    xml = xml.replace('<dimension ref="A1:J40"', '<dimension ref="A1:L40"', 1)
    xml = xml.replace('</cols>',
                      '<col min="11" max="12" width="12" style="1" customWidth="1"/></cols>', 1)
    changed[DIRECTORY_SHEET] = xml.encode("utf-8")

    # --- workbook.xml: add dir_lat / dir_lon named ranges ---
    wbx = zin.read("xl/workbook.xml").decode("utf-8")
    inject = ('<definedName name="dir_lat">Directory!$K$5:$K$33</definedName>'
              '<definedName name="dir_lon">Directory!$L$5:$L$33</definedName>')
    wbx2 = wbx.replace('<definedName name="hq_code">', inject + '<definedName name="hq_code">', 1)
    if wbx2 == wbx:
        raise SystemExit("ERROR: could not inject named ranges into workbook.xml")
    changed["xl/workbook.xml"] = wbx2.encode("utf-8")

    # --- How To Use text fix (best-effort) ---
    ss = zin.read("xl/sharedStrings.xml").decode("utf-8")
    old = "Edit the green LAT/LON values to refine site positions."
    new = "Edit the LAT/LON columns (K/L) on the Directory tab to refine site positions."
    if old in ss:
        changed["xl/sharedStrings.xml"] = ss.replace(old, new, 1).encode("utf-8")

    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, changed.get(info.filename, zin.read(info.filename)))
    zin.close()


# ----------------------------- main ------------------------------------------
def main():
    hq, rows, office_miles = read_directory_and_matrix()
    cache = load_json(CACHE)
    overrides = load_json(OVERRIDES)

    print("Geocoding %d sites (%d unique addresses)…" %
          (len(rows), len({r["address"] for r in rows})))
    coords_by_code = {}
    methods = {}
    for item in rows:
        code, addr = item["code"], item["address"]
        if code in overrides:
            coords_by_code[code] = [float(overrides[code][0]), float(overrides[code][1])]
            methods[code] = "override"
            continue
        coords, method = geocode_address(addr, cache)
        if coords is None:
            methods[code] = "FAILED"
            print("  ! %s  FAILED to geocode: %s" % (code, addr))
            continue
        coords_by_code[code] = coords
        methods[code] = method

    # -------- validation report --------
    print("\n=== VALIDATION ===")
    office = coords_by_code.get(hq)
    flags = []
    for item in rows:
        code = item["code"]
        c = coords_by_code.get(code)
        if c is None:
            flags.append("%s: no coordinates" % code)
            continue
        if not (TX["lat_min"] <= c[0] <= TX["lat_max"] and TX["lon_min"] <= c[1] <= TX["lon_max"]):
            flags.append("%s: outside Texas bbox %s" % (code, c))
        if office and code != hq:
            crow = haversine_mi(office, c)
            rec = office_miles.get(code)
            if rec:
                if crow > rec * 1.05:
                    flags.append("%s: crow-flies %.1f > recorded road %.1f mi (impossible — check pin)"
                                 % (code, crow, rec))
                elif rec > 3 and crow < rec * 0.35:
                    flags.append("%s: crow-flies %.1f << recorded road %.1f mi (possible wrong pin)"
                                 % (code, crow, rec))

    # co-location consistency (same address -> same coords)
    by_addr = {}
    for item in rows:
        by_addr.setdefault(item["address"], []).append(item["code"])
    for addr, codes in by_addr.items():
        cs = {tuple(coords_by_code[c]) for c in codes if c in coords_by_code}
        if len(cs) > 1:
            flags.append("shared address geocoded inconsistently: %s -> %s" % (codes, cs))

    method_counts = {}
    for m in methods.values():
        method_counts[m] = method_counts.get(m, 0) + 1
    print("methods:", method_counts)
    print("coverage: %d/%d sites have coordinates" % (len(coords_by_code), len(rows)))
    if flags:
        print("FLAGS (%d) — review these (add corrections to coord_overrides.json and re-run):" % len(flags))
        for f in flags:
            print("  -", f)
    else:
        print("No validation flags — all sites within TX and consistent with Matrix distances.")

    if len(coords_by_code) != len(rows):
        print("\nNOT writing workbook: coverage incomplete. Fix failures first.")
        return

    write_v3(hq, rows, coords_by_code)
    print("\nWrote", os.path.basename(OUT), "with LAT/LON in Directory!K5:L33")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
