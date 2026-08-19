#!/usr/bin/env python3
"""
add_homes_v6.py  —  Add Employee A's two home addresses to the Fleet Distance
Dashboard as new locations, expanding BOTH the Directory and the Matrix.

Reads : Fleet_Distance_Dashboard_v5.xlsx   (never modified)
Writes: Fleet_Distance_Dashboard_v6.xlsx

All edits are surgical XML edits inside the .xlsx zip so the slicer / pivotCache /
data-model / Table1 / embedded image / cached formulas are preserved (a full
openpyxl re-save would drop them).

New locations (both flagged as "Employee Home"):
  EH-1  Employee A Home (Apt)   123 Demo St, Democity, TX 70001
  EH-2  Employee A Home (Second) 456 Sample Ave, Demotown, TX 70002

Home<->site distances are OSRM road-routing estimates (NOT hand-verified like the
rest of the matrix). The moved footnote still marks the grid as the source of truth.
"""
import zipfile, json, re, os, shutil, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC  = os.path.join(HERE, "Fleet_Distance_Dashboard_v5.xlsx")
DST  = os.path.join(HERE, "Fleet_Distance_Dashboard_v6.xlsx")
DATA = os.path.join(HERE, "data", "gathered.json")   # OSRM leg data (not included in this demo)

g = json.load(open(DATA, encoding="utf-8"))
sites, homes, osrm = g["sites"], g["homes"], g["osrm"]
assert len(sites) == 29 and len(homes) == 2
col_cc, col_cv = osrm["col_cc"], osrm["col_cv"]   # site i -> home (len 31, use 0..28)
row_cc, row_cv = osrm["row_cc"], osrm["row_cv"]   # home -> everything (len 31, B..AF)

# home coords (finalized)
H = {h["code"]: h for h in homes}
CC, CV = H["EH-1"], H["EH-2"]

def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

# ---- shared string indices (append after existing 203) ----
S_CC, S_CV, S_N1, S_N2, S_A1, S_A2, S_TYPE = 203, 204, 205, 206, 207, 208, 209
NEW_STRINGS = [
    "EH-1", "EH-2",
    "Employee A Home (Apt)", "Employee A Home (Second)",
    "123 Demo St, Democity, TX 70001",
    "456 Sample Ave, Demotown, TX 70002",
    "Employee Home",
]

def must_replace(text, old, new, n=1, label=""):
    cnt = text.count(old)
    if cnt < n:
        raise SystemExit(f"FATAL: expected >= {n} of [{label or old[:40]}], found {cnt}")
    return text.replace(old, new, n)

# ============================================================ sharedStrings.xml
def edit_sharedstrings(x):
    x = must_replace(x, 'count="424" uniqueCount="203"',
                        'count="436" uniqueCount="210"', label="sst header")
    add = "".join(f"<si><t xml:space=\"preserve\">{s}</t></si>" for s in NEW_STRINGS)
    x = must_replace(x, "</sst>", add + "</sst>", label="sst tail")
    return x

# ============================================================ Matrix (sheet2)
def edit_matrix(x):
    # widen every row's span
    x = x.replace('spans="1:31"', 'spans="1:32"')
    # dimension
    x = must_replace(x, '<dimension ref="A1:AE34"/>', '<dimension ref="A1:AF34"/>',
                     label="mtx dimension")
    # header row 1: fill AE1/AF1 with the two new codes
    x = must_replace(x, '<c r="AE1" s="2"/></row>',
                     f'<c r="AE1" s="15" t="s"><v>{S_CC}</v></c>'
                     f'<c r="AF1" s="15" t="s"><v>{S_CV}</v></c></row>',
                     label="mtx AE1")
    # rows 2..30: append AE (site->EH-1) and AF (site->EH-2)
    for r in range(2, 31):
        i = r - 2
        x = must_replace(
            x, f'<c r="AE{r}" s="2"/></row>',
            f'<c r="AE{r}" s="21"><v>{col_cc[i]}</v></c>'
            f'<c r="AF{r}" s="21"><v>{col_cv[i]}</v></c></row>',
            label=f"mtx AE{r}")
    # helper to build a full home data row (B..AF from a 31-long value list)
    def home_row(r, code_sid, vals):
        cells = [f'<c r="A{r}" s="20" t="s"><v>{code_sid}</v></c>']
        for j in range(31):                       # j=0 -> col B(2) ... j=30 -> AF(32)
            cells.append(f'<c r="{col_letter(j+2)}{r}" s="21"><v>{vals[j]}</v></c>')
        return (f'<row r="{r}" spans="1:32" ht="15" customHeight="1" '
                f'x14ac:dyDescent="0.25">' + "".join(cells) + "</row>")
    # replace whole placeholder row 31 (empty) -> EH-1
    x = re.subn(r'<row r="31"[ >].*?</row>', home_row(31, S_CC, row_cc), x, flags=re.S)
    x = _one(x, "mtx row31")
    # replace whole placeholder row 32 (old footnote) -> EH-2
    x = re.subn(r'<row r="32"[ >].*?</row>', home_row(32, S_CV, row_cv), x, flags=re.S)
    x = _one(x, "mtx row32")
    # rebuild row 34 as the relocated footnote (merged A34:AF34, string #95)
    foot_cells = ['<c r="A34" s="91" t="s"><v>95</v></c>']
    for j in range(31):
        foot_cells.append(f'<c r="{col_letter(j+2)}34" s="91"/>')
    foot_row = ('<row r="34" spans="1:32" ht="15" customHeight="1" '
                'x14ac:dyDescent="0.25">' + "".join(foot_cells) + "</row>")
    x = re.subn(r'<row r="34"[ >].*?</row>', foot_row, x, flags=re.S)
    x = _one(x, "mtx row34")
    # move the merge
    x = must_replace(x, 'A32:AD32', 'A34:AF34', label="mtx merge")
    return x

_pending = None
def _one(subn_result, label):
    global _pending
    text, n = subn_result
    if n != 1:
        raise SystemExit(f"FATAL: {label} replaced {n} times (expected 1)")
    return text

# ============================================================ Directory (sheet3)
def edit_directory(x):
    def dir_row(r, code_sid, name_sid, addr_sid, f_cached, label_cached, lat, lon):
        return (
            f'<row r="{r}" spans="1:12" ht="21.75" customHeight="1" x14ac:dyDescent="0.25">'
            f'<c r="A{r}" s="2"/>'
            f'<c r="B{r}" s="35" t="s"><v>{code_sid}</v></c>'
            f'<c r="C{r}" s="36" t="s"><v>{name_sid}</v></c>'
            f'<c r="D{r}" s="37" t="s"><v>{addr_sid}</v></c>'
            f'<c r="E{r}" s="38" t="s"><v>{S_TYPE}</v></c>'
            f'<c r="F{r}" s="29"><f>IFERROR(INDEX(mtx_vals,MATCH(hq_code,mtx_rows,0),MATCH(B{r},mtx_cols,0)),"")</f><v>{f_cached}</v></c>'
            f'<c r="G{r}" s="2"/>'
            f'<c r="H{r}" s="30" t="str"><f>B{r}&amp;" — "&amp;C{r}</f><v>{label_cached}</v></c>'
            f'<c r="I{r}" s="2"/><c r="J{r}" s="2"/>'
            f'<c r="K{r}" s="53"><v>{lat}</v></c>'
            f'<c r="L{r}" s="53"><v>{lon}</v></c>'
            f'</row>')
    r34 = dir_row(34, S_CC, S_N1, S_A1, col_cc[0],
                  "EH-1 — Employee A Home (Apt)", CC["lat"], CC["lon"])
    r35 = dir_row(35, S_CV, S_N2, S_A2, col_cv[0],
                  "EH-2 — Employee A Home (Second)", CV["lat"], CV["lon"])
    x = _one(re.subn(r'<row r="34"[ >].*?</row>', r34, x, flags=re.S), "dir row34")
    x = _one(re.subn(r'<row r="35"[ >].*?</row>', r35, x, flags=re.S), "dir row35")
    return x

# ============================================================ workbook.xml
def edit_workbook(x):
    for c in ["B", "C", "D", "E", "F", "H", "K", "L"]:
        x = must_replace(x, f'Directory!${c}$5:${c}$33', f'Directory!${c}$5:${c}$35',
                         label=f"dir range {c}")
    x = must_replace(x, 'Matrix!$B$1:$AD$1', 'Matrix!$B$1:$AF$1', label="mtx_cols")
    x = must_replace(x, 'Matrix!$A$2:$A$30', 'Matrix!$A$2:$A$32', label="mtx_rows")
    x = must_replace(x, 'Matrix!$B$2:$AD$30', 'Matrix!$B$2:$AF$32', label="mtx_vals")
    x = must_replace(x, 'Directory!$B$1:$F$33', 'Directory!$B$1:$F$35', label="dir print")
    x = must_replace(x, 'Matrix!$A$1:$AD$30', 'Matrix!$A$1:$AF$32', label="mtx print")
    x = must_replace(x, '<calcPr calcId="191029"/>',
                     '<calcPr calcId="191029" fullCalcOnLoad="1"/>', label="calcPr")
    return x

EDITORS = {
    "xl/sharedStrings.xml": edit_sharedstrings,
    "xl/worksheets/sheet2.xml": edit_matrix,
    "xl/worksheets/sheet3.xml": edit_directory,
    "xl/workbook.xml": edit_workbook,
}

def main():
    if not os.path.exists(SRC):
        raise SystemExit("missing source v5")
    zin = zipfile.ZipFile(SRC, "r")
    tmp = DST + ".tmp"
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    touched = set()
    for item in zin.infolist():
        raw = zin.read(item.filename)
        if item.filename in EDITORS:
            text = raw.decode("utf-8")
            text = EDITORS[item.filename](text)
            raw = text.encode("utf-8")
            touched.add(item.filename)
            print(f"  edited {item.filename}  ({len(raw)} bytes)")
        zout.writestr(item, raw)
    zin.close(); zout.close()
    missing = set(EDITORS) - touched
    if missing:
        os.remove(tmp); raise SystemExit(f"FATAL: never edited {missing}")
    shutil.move(tmp, DST)
    print("WROTE", DST)

if __name__ == "__main__":
    main()
