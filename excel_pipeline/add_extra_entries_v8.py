#!/usr/bin/env python3
r"""
add_extra_entries_v8.py  ->  Fleet_Distance_Dashboard_v8.xlsx

Ports the HR-HTML "Additional / Temporary Entries" feature into the Excel
Monthly Plan tab (sheet6.xml). For a site NOT yet in the system, HR types the
Site / Address + Miles by hand; those miles roll into a GRAND "TOTAL PLAN MILES"
and the ESTIMATED REIMBURSEMENT, exactly like the HTML.

Surgical XML zip edit from the CLEAN protected v7 master (openpyxl re-save would
drop slicer/pivotCache/data-model/Table1/image). Only styles.xml, workbook.xml and
worksheets/sheet6.xml change; every other part is copied byte-for-byte.

Layout change on the Monthly Plan tab (cols B..G, dark margins A/I/J, hidden H):
  row 56  "PLAN TRIP MILES"          = SUM(F16:F55)      (was "TOTAL PLAN MILES")
  row 58  "03  ADDITIONAL / TEMPORARY ENTRIES"           (section header)
  row 61  table header  # | Date | Site / Address | Miles | Notes
  row 62..71  10 hand-entry rows (Date/Site/Miles/Notes all UNLOCKED)
  row 72  "EXTRA / TEMPORARY MILES"  = SUM(F62:F71)
  row 73  "TOTAL PLAN MILES"         = F56+F72           (grand total)
  row 74  "ESTIMATED REIMBURSEMENT"  = IF(F10="","",F73*F10)
  row 75  footnote
Print area B1:G58 -> B1:G75 ; dimension A1:J58 -> A1:J75.
Two new unlocked numeric cellXfs (105 non-last / 106 last-row) for the Miles input.
"""
import zipfile, re, os, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
SRC  = os.path.join(HERE, "Latest Version (v7)", "Fleet_Distance_Dashboard_v7.xlsx")
OUT  = os.path.join(HERE, "Fleet_Distance_Dashboard_v8.xlsx")

COLIDX = {c: i + 1 for i, c in enumerate("ABCDEFGHIJ")}

def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
def istr(ref, s, text):
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{esc(text)}</t></is></c>'
def num(ref, s, v):
    return f'<c r="{ref}" s="{s}"><v>{v}</v></c>'
def blank(ref, s):
    return f'<c r="{ref}" s="{s}"/>'
def formula(ref, s, f):
    return f'<c r="{ref}" s="{s}"><f>{esc(f)}</f></c>'

# ---- new rows 57..75 --------------------------------------------------------
def build_new_rows():
    cells = {}
    def put(r, ref, xml):
        col = COLIDX[re.match(r'[A-Z]+', ref).group(0)]
        cells.setdefault(r, {})[col] = xml
    def margins(r):
        put(r, f"A{r}", blank(f"A{r}", 2))
        put(r, f"H{r}", blank(f"H{r}", 2))   # r>24 -> no month helper
        put(r, f"I{r}", blank(f"I{r}", 2))
        put(r, f"J{r}", blank(f"J{r}", 2))
    def band(r, first_s, rest_s, text=None):
        put(r, f"B{r}", istr(f"B{r}", first_s, text) if text is not None else blank(f"B{r}", first_s))
        for c in "CDEFG":
            put(r, f"{c}{r}", blank(f"{c}{r}", rest_s))

    for r in range(57, 76):
        margins(r)

    band(57, 2, 2)                                   # spacer
    band(58, 72, 72, "03   ADDITIONAL / TEMPORARY ENTRIES")
    band(59, 74, 74,
         "For a site NOT in the system yet, type the Site / Address and the Miles by hand. "
         "These add into the Total Plan Miles and the Estimated Reimbursement below — "
         "like adding extra rows at the bottom of a sheet.")
    band(60, 2, 2)                                   # spacer
    # table header (row 61)  B # | C Date | D:E Site/Address | F Miles | G Notes
    put(61, "B61", istr("B61", 3, "#"))
    put(61, "C61", istr("C61", 4, "Date"))
    put(61, "D61", istr("D61", 4, "Site / Address"))
    put(61, "E61", blank("E61", 4))
    put(61, "F61", istr("F61", 4, "Miles"))
    put(61, "G61", istr("G61", 5, "Notes"))
    # 10 hand-entry rows 62..71
    E_FIRST, E_LAST = 62, 71
    for i in range(10):
        r = E_FIRST + i
        last = (r == E_LAST)
        s_num  = 10 if last else 6      # locked #
        s_date = 101 if last else 99    # unlocked date
        s_site = 65 if last else 63     # unlocked text (D:E merged)
        s_mile = 106 if last else 105   # unlocked NUMBER (new)
        s_note = 102 if last else 100   # unlocked notes
        put(r, f"B{r}", num(f"B{r}", s_num, i + 1))
        put(r, f"C{r}", blank(f"C{r}", s_date))
        put(r, f"D{r}", blank(f"D{r}", s_site))
        put(r, f"E{r}", blank(f"E{r}", s_site))
        put(r, f"F{r}", blank(f"F{r}", s_mile))
        put(r, f"G{r}", blank(f"G{r}", s_note))
    # EXTRA subtotal (row 72)
    put(72, "B72", istr("B72", 67, "EXTRA / TEMPORARY MILES"))
    for c in "CDE": put(72, f"{c}72", blank(f"{c}72", 67))
    put(72, "F72", formula("F72", 62, f"SUM(F{E_FIRST}:F{E_LAST})"))
    put(72, "G72", blank("G72", 62))
    # GRAND total (row 73)
    put(73, "B73", istr("B73", 67, "TOTAL PLAN MILES"))
    for c in "CDE": put(73, f"{c}73", blank(f"{c}73", 67))
    put(73, "F73", formula("F73", 62, "F56+F72"))
    put(73, "G73", blank("G73", 62))
    # reimbursement (row 74)
    put(74, "B74", istr("B74", 67, "ESTIMATED REIMBURSEMENT"))
    for c in "CDE": put(74, f"{c}74", blank(f"{c}74", 67))
    put(74, "F74", formula("F74", 103, 'IF(F10="","",F73*F10)'))
    put(74, "G74", blank("G74", 103))
    # footnote (row 75)
    band(75, 74, 74,
         "Distances for system sites are Matrix estimates; extra / temporary entries are "
         "hand-entered. Set Rate $/mi above to estimate reimbursement.")

    ht = {57: 10, 58: 20, 59: 30, 60: 6, 61: 20, 72: 22, 73: 22, 74: 22, 75: 32}
    out = []
    for r in sorted(cells):
        rowcells = "".join(cells[r][c] for c in sorted(cells[r]))
        h = ht.get(r, 19)
        out.append(f'<row r="{r}" ht="{h}" customHeight="1">{rowcells}</row>')
    return "".join(out)

def new_merges():
    merges = ['B1:G1','B2:G2','B3:G3','B4:G4','B6:G6','B7:G7',
              'B9:C9','D9:E9','F9:G9','B10:C10','D10:E10','F10:G10',
              'B12:G12','B13:G13','B56:E56','F56:G56',
              'B58:G58','B59:G59','D61:E61']
    for r in range(62, 72):
        merges.append(f'D{r}:E{r}')
    merges += ['B72:E72','F72:G72','B73:E73','F73:G73','B74:E74','F74:G74','B75:G75']
    return (f'<mergeCells count="{len(merges)}">'
            + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + '</mergeCells>')

# ---- styles: append 2 unlocked numeric cellXfs (105,106) ---------------------
NEW_XFS = [
    '<xf numFmtId="0" fontId="18" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="center"/><protection locked="0"/></xf>',
    '<xf numFmtId="0" fontId="18" fillId="4" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="center"/><protection locked="0"/></xf>',
]
def patch_styles(st):
    m = re.search(r'<cellXfs count="(\d+)">', st)
    cnt = int(m.group(1))
    if cnt == 107:
        return st          # already patched
    if cnt != 105:
        raise SystemExit(f"cellXfs count {cnt} != 105 or 107; refusing to guess")
    st = st.replace('<cellXfs count="105">', '<cellXfs count="107">', 1)
    st = st.replace('</cellXfs>', "".join(NEW_XFS) + '</cellXfs>', 1)
    return st

def patch_workbook(wb):
    old = "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"1\">'Monthly Plan'!$B$1:$G$58</definedName>"
    new = "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"1\">'Monthly Plan'!$B$1:$G$75</definedName>"
    if old not in wb:
        if "$G$75</definedName>" in wb:
            return wb
        raise SystemExit("Monthly Plan Print_Area (B1:G58) not found in workbook.xml")
    return wb.replace(old, new, 1)

def patch_sheet6(x):
    # dimension
    x = x.replace('<dimension ref="A1:J58"/>', '<dimension ref="A1:J75"/>', 1)
    # version bump v7 -> v8 (subtitle)
    x = x.replace("Version v7  ·", "Version v8  ·", 1)
    # relabel row 56 subtotal
    x = x.replace(
        '<c r="B56" s="67" t="inlineStr"><is><t xml:space="preserve">TOTAL PLAN MILES</t></is></c>',
        '<c r="B56" s="67" t="inlineStr"><is><t xml:space="preserve">PLAN TRIP MILES</t></is></c>', 1)
    # drop old rows 57 (reimbursement) & 58 (footnote)
    x = re.sub(r'<row r="57"[ >].*?</row>', '', x, count=1, flags=re.S)
    x = re.sub(r'<row r="58"[ >].*?</row>', '', x, count=1, flags=re.S)
    # append new rows just before </sheetData>
    x = x.replace('</sheetData>', build_new_rows() + '</sheetData>', 1)
    # replace mergeCells block
    x = re.sub(r'<mergeCells count="\d+">.*?</mergeCells>', new_merges(), x, count=1, flags=re.S)
    return x

def main():
    zin = zipfile.ZipFile(SRC, "r")
    names = zin.namelist()
    parts = {n: zin.read(n) for n in names}
    infos = {i.filename: i for i in zin.infolist()}
    zin.close()

    parts["xl/styles.xml"]   = patch_styles(parts["xl/styles.xml"].decode("utf-8")).encode("utf-8")
    parts["xl/workbook.xml"] = patch_workbook(parts["xl/workbook.xml"].decode("utf-8")).encode("utf-8")
    parts["xl/worksheets/sheet6.xml"] = patch_sheet6(parts["xl/worksheets/sheet6.xml"].decode("utf-8")).encode("utf-8")

    tmp = OUT + ".tmp"
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for n in names:
        zout.writestr(infos[n], parts[n]) if n in infos else zout.writestr(n, parts[n])
    zout.close()
    if os.path.exists(OUT):
        os.remove(OUT)
    os.rename(tmp, OUT)
    print("WROTE:", OUT)

if __name__ == "__main__":
    main()
