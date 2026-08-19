#!/usr/bin/env python3
r"""
restyle_monthly_plan_v6.py — re-skin the "Monthly Plan" worksheet (sheet6.xml)
in Fleet_Distance_Dashboard_v6.xlsx to match the dark, branded look of the
Dashboard tab.

Why: the Monthly Plan tab was functional but visually plain — dark but missing
the gold brand band, hero title, numbered section headers, styled input pills and
a proper rounded data-table container; the "Employee" label was clipped and there
was a white gap on the right. This rebuild reuses the Dashboard's OWN style
indices (verified in sheet1.xml) so the two tabs read as one system, and adds the
Rate $/mi -> estimated-reimbursement figure that the HTML plan has.

Approach: surgical XML zip edit (openpyxl re-save would drop the slicer/pivotCache/
data-model/Table1/image). We ONLY:
  1. append 1 numFmt (currency) + 5 cellXfs to styles.xml for table-edge/currency
     combos the Dashboard didn't need,
  2. replace xl/worksheets/sheet6.xml with the new branded layout,
  3. fix the Print_Area/Print_Titles localSheetIds that went stale when the
     Monthly Plan sheet was inserted at index 1 (Matrix 1->2, Directory 2->3) and
     add a Print_Area for the Monthly Plan itself.
Every other worksheet's bytes are left untouched. Existing formulas / named ranges
/ dropdowns are preserved (Miles auto-calc from the Matrix, From/To = dir_label,
Month = plan_months).

Reads a pristine backup (…v6.pre_restyle.xlsx, created on first run) so the script
is safe to re-run. Writes a temp; caller COM-verifies before replacing.
"""
import zipfile, re, os, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, "Fleet_Distance_Dashboard_v6.xlsx")
BACKUP = os.path.join(HERE, "Fleet_Distance_Dashboard_v6.pre_restyle.xlsx")

# ---- layout constants -------------------------------------------------------
# Content lives in columns B..G; A is a dark left margin, H is the hidden month
# helper (plan_months), I/J are dark right margin.
FIRST = 16                     # first trip-log data row
ROWS = 40
LAST = FIRST + ROWS - 1        # 55
TOTAL_ROW = LAST + 1           # 56
EST_ROW = TOTAL_ROW + 1        # 57
FOOT_ROW = EST_ROW + 1         # 58

MONTHS = [f"{m} {y}" for y in (2026, 2027) for m in
          ["January","February","March","April","May","June","July","August",
           "September","October","November","December"]]

COLIDX = {c: i + 1 for i, c in enumerate("ABCDEFGHIJ")}

def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def istr(ref, style, text):
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{esc(text)}</t></is></c>'

def num(ref, style, v):
    return f'<c r="{ref}" s="{style}"><v>{v}</v></c>'

def blank(ref, style):
    return f'<c r="{ref}" s="{style}"/>'

def formula(ref, style, f):
    return f'<c r="{ref}" s="{style}"><f>{esc(f)}</f></c>'

# ---- build sheet6.xml -------------------------------------------------------
def build_sheet_data():
    cells = {}                 # row -> {colidx: xml}
    def put(r, ref, xml):
        col = COLIDX[re.match(r'[A-Z]+', ref).group(0)]
        cells.setdefault(r, {})[col] = xml

    # dark margins (A, I, J) + hidden month helper (H) on every used row
    for r in range(1, FOOT_ROW + 1):
        put(r, f"A{r}", blank(f"A{r}", 2))
        put(r, f"I{r}", blank(f"I{r}", 2))
        put(r, f"J{r}", blank(f"J{r}", 2))
        if r <= 24:
            put(r, f"H{r}", istr(f"H{r}", 2, MONTHS[r - 1]))
        else:
            put(r, f"H{r}", blank(f"H{r}", 2))

    def band(r, first_style, rest_style, text=None):
        put(r, f"B{r}", istr(f"B{r}", first_style, text) if text is not None
            else blank(f"B{r}", first_style))
        for c in "CDEFG":
            put(r, f"{c}{r}", blank(f"{c}{r}", rest_style))

    # --- hero ---------------------------------------------------------------
    band(1, 54, 55, "DEMO FLEET CO     ·     MONTHLY TRAVEL PLAN")
    band(2, 86, 86, "●  Fleet OPERATIONS  ·  MILEAGE PLANNING")
    band(3, 87, 87, "MONTHLY TRAVEL PLAN")
    band(4, 88, 88,
         "Plan one employee's month of trips — pick the month, then fill each day's "
         "Date and From → To. Miles auto-fill from the Matrix (road-distance estimates; "
         "home↔site legs are OSRM estimates).   Version v6  ·  Source of truth = "
         "Matrix + Directory tabs.")
    band(5, 2, 2)  # spacer

    # --- 01 plan setup ------------------------------------------------------
    band(6, 72, 72, "01   PLAN SETUP")
    band(7, 74, 74,
         "Enter the employee's name, pick the plan Month, and (optionally) the mileage "
         "Rate in $/mi to estimate the reimbursement.")
    band(8, 2, 2)
    # labels row 9
    put(9, "B9", istr("B9", 104, "EMPLOYEE")); put(9, "C9", blank("C9", 104))
    put(9, "D9", istr("D9", 104, "MONTH"));    put(9, "E9", blank("E9", 104))
    put(9, "F9", istr("F9", 104, "RATE  $/MI"));put(9, "G9", blank("G9", 104))
    # inputs row 10 (pills: first cell s77 unlocked, continuation s78)
    put(10, "B10", istr("B10", 77, "")); put(10, "C10", blank("C10", 78))
    put(10, "D10", istr("D10", 77, "")); put(10, "E10", blank("E10", 78))
    put(10, "F10", istr("F10", 77, "")); put(10, "G10", blank("G10", 78))
    band(11, 2, 2)

    # --- 02 trip log --------------------------------------------------------
    band(12, 72, 72, "02   TRIP LOG")
    band(13, 74, 74,
         "Fill each day's Date and choose From / To sites. Miles fill automatically "
         "from the Matrix. Totals and the estimate update as you type.")
    band(14, 2, 2)
    # header row 15
    put(15, "B15", istr("B15", 3, "#"))
    put(15, "C15", istr("C15", 4, "Date"))
    put(15, "D15", istr("D15", 4, "From"))
    put(15, "E15", istr("E15", 4, "To"))
    put(15, "F15", istr("F15", 4, "Miles"))
    put(15, "G15", istr("G15", 5, "Notes"))

    MILES_F = ("IFERROR(INDEX(mtx_vals,"
               "MATCH(INDEX(dir_code,MATCH(D{r},dir_label,0)),mtx_rows,0),"
               "MATCH(INDEX(dir_code,MATCH(E{r},dir_label,0)),mtx_cols,0)),\"\")")
    for i in range(ROWS):
        r = FIRST + i
        last = (r == LAST)
        s_num  = 10 if last else 6
        s_date = 101 if last else 99
        s_site = 65 if last else 63
        s_mile = 12 if last else 8
        s_note = 102 if last else 100
        put(r, f"B{r}", num(f"B{r}", s_num, i + 1))
        put(r, f"C{r}", blank(f"C{r}", s_date))
        put(r, f"D{r}", blank(f"D{r}", s_site))
        put(r, f"E{r}", blank(f"E{r}", s_site))
        put(r, f"F{r}", formula(f"F{r}", s_mile, MILES_F.format(r=r)))
        put(r, f"G{r}", blank(f"G{r}", s_note))

    # total row (gold bar)
    put(TOTAL_ROW, f"B{TOTAL_ROW}", istr(f"B{TOTAL_ROW}", 67, "TOTAL PLAN MILES"))
    for c in "CDE":
        put(TOTAL_ROW, f"{c}{TOTAL_ROW}", blank(f"{c}{TOTAL_ROW}", 67))
    put(TOTAL_ROW, f"F{TOTAL_ROW}", formula(f"F{TOTAL_ROW}", 62, f"SUM(F{FIRST}:F{LAST})"))
    put(TOTAL_ROW, f"G{TOTAL_ROW}", blank(f"G{TOTAL_ROW}", 62))

    # estimated reimbursement row (gold bar, currency)
    put(EST_ROW, f"B{EST_ROW}", istr(f"B{EST_ROW}", 67, "ESTIMATED REIMBURSEMENT"))
    for c in "CDE":
        put(EST_ROW, f"{c}{EST_ROW}", blank(f"{c}{EST_ROW}", 67))
    put(EST_ROW, f"F{EST_ROW}", formula(f"F{EST_ROW}", 103,
        f'IF(F10="","",F{TOTAL_ROW}*F10)'))
    put(EST_ROW, f"G{EST_ROW}", blank(f"G{EST_ROW}", 103))

    # footnote
    band(FOOT_ROW, 74, 74,
         "Distances are road-mileage estimates from the Matrix tab; home↔site legs are "
         "OSRM estimates and are directional (A→B may differ slightly from B→A). "
         "Set Rate $/mi above to estimate reimbursement.")

    # emit
    ht = {1: 26, 2: 16, 3: 34, 4: 42, 5: 8, 6: 20, 7: 30, 8: 6, 9: 15, 10: 22,
          11: 8, 12: 20, 13: 22, 14: 6, 15: 20, TOTAL_ROW: 22, EST_ROW: 22,
          FOOT_ROW: 32}
    out = []
    for r in sorted(cells):
        rowcells = "".join(cells[r][c] for c in sorted(cells[r]))
        h = ht.get(r, 19)
        out.append(f'<row r="{r}" ht="{h}" customHeight="1">{rowcells}</row>')
    return "".join(out)

def sheet6_xml():
    data = build_sheet_data()
    merges = ['B1:G1','B2:G2','B3:G3','B4:G4','B6:G6','B7:G7',
              'B9:C9','D9:E9','F9:G9','B10:C10','D10:E10','F10:G10',
              'B12:G12','B13:G13',
              f'B{TOTAL_ROW}:E{TOTAL_ROW}', f'F{TOTAL_ROW}:G{TOTAL_ROW}',
              f'B{EST_ROW}:E{EST_ROW}', f'F{EST_ROW}:G{EST_ROW}',
              f'B{FOOT_ROW}:G{FOOT_ROW}']
    mc = (f'<mergeCells count="{len(merges)}">'
          + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + '</mergeCells>')
    dv = ('<dataValidations count="3">'
          f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="D{FIRST}:D{LAST}"><formula1>dir_label</formula1></dataValidation>'
          f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E{FIRST}:E{LAST}"><formula1>dir_label</formula1></dataValidation>'
          '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="D10"><formula1>plan_months</formula1></dataValidation>'
          '</dataValidations>')
    cols = ('<cols>'
            '<col min="1" max="1" width="2.5" style="1" customWidth="1"/>'
            '<col min="2" max="2" width="5.5" style="1" customWidth="1"/>'
            '<col min="3" max="3" width="13" style="1" customWidth="1"/>'
            '<col min="4" max="5" width="30" style="1" customWidth="1"/>'
            '<col min="6" max="6" width="9" style="1" customWidth="1"/>'
            '<col min="7" max="7" width="30" style="1" customWidth="1"/>'
            '<col min="8" max="8" width="14" style="1" hidden="1" customWidth="1"/>'
            '<col min="9" max="10" width="3" style="1" customWidth="1"/>'
            '</cols>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" '
        'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
        'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" '
        'xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" '
        'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" '
        'xr:uid="{00000000-0001-0000-0600-000000000000}">'
        '<sheetPr><tabColor rgb="FF52D1A0"/><pageSetUpPr fitToPage="1"/></sheetPr>'
        f'<dimension ref="A1:J{FOOT_ROW}"/>'
        '<sheetViews><sheetView showGridLines="0" workbookViewId="0"><selection activeCell="D16" sqref="D16"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
        + cols
        + '<sheetData>' + data + '</sheetData>'
        + '<sheetProtection sheet="1" objects="1" scenarios="1"/>'
        + mc + dv
        + '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
        + '<pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>'
        + '</worksheet>')

# ---- styles.xml additions ---------------------------------------------------
NEW_XFS = [
    # 100 Notes body: fill4 input, right border, left, unlocked
    '<xf numFmtId="0" fontId="18" fillId="4" borderId="13" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="left" vertical="center"/><protection locked="0"/></xf>',
    # 101 Date last row: date fmt, fill4, bottom border, center, unlocked
    '<xf numFmtId="14" fontId="18" fillId="4" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="center"/><protection locked="0"/></xf>',
    # 102 Notes last row: fill4, right+bottom border, left, unlocked
    '<xf numFmtId="0" fontId="18" fillId="4" borderId="15" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="left" vertical="center"/><protection locked="0"/></xf>',
    # 103 Est $ value: currency (165), gold fill6, gold border, center
    '<xf numFmtId="165" fontId="22" fillId="6" borderId="16" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>',
    # 104 setup label: muted bold, canvas fill2, no border, left
    '<xf numFmtId="0" fontId="4" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>',
]
CURRENCY_NUMFMT = '<numFmt numFmtId="165" formatCode="&quot;$&quot;#,##0.00"/>'

def patch_styles(st):
    # numFmt (idempotent)
    if 'numFmtId="165"' not in st:
        m = re.search(r'<numFmts count="(\d+)">', st)
        cnt = int(m.group(1))
        st = st.replace(f'<numFmts count="{cnt}">', f'<numFmts count="{cnt+1}">', 1)
        st = st.replace('</numFmts>', CURRENCY_NUMFMT + '</numFmts>', 1)
    # cellXfs (idempotent): append only if not already at 105
    m = re.search(r'<cellXfs count="(\d+)">', st)
    cnt = int(m.group(1))
    if cnt == 100:
        st = st.replace('<cellXfs count="100">', '<cellXfs count="105">', 1)
        st = st.replace('</cellXfs>', "".join(NEW_XFS) + '</cellXfs>', 1)
    elif cnt != 105:
        raise SystemExit(f"cellXfs count {cnt} != 100 or 105; refusing to guess")
    return st

def patch_workbook(wb):
    # Fix stale localSheetIds (Monthly Plan was inserted at doc index 1):
    #   Matrix    Print_Area/Titles: 1 -> 2 ; Directory: 2 -> 3
    fixes = [
        ('<definedName name="_xlnm.Print_Area" localSheetId="2">Directory!$B$1:$F$35</definedName>',
         '<definedName name="_xlnm.Print_Area" localSheetId="3">Directory!$B$1:$F$35</definedName>'),
        ('<definedName name="_xlnm.Print_Area" localSheetId="1">Matrix!$A$1:$AF$32</definedName>',
         '<definedName name="_xlnm.Print_Area" localSheetId="2">Matrix!$A$1:$AF$32</definedName>'),
        ('<definedName name="_xlnm.Print_Titles" localSheetId="2">Directory!$4:$4</definedName>',
         '<definedName name="_xlnm.Print_Titles" localSheetId="3">Directory!$4:$4</definedName>'),
        ('<definedName name="_xlnm.Print_Titles" localSheetId="1">Matrix!$A:$A,Matrix!$1:$1</definedName>',
         '<definedName name="_xlnm.Print_Titles" localSheetId="2">Matrix!$A:$A,Matrix!$1:$1</definedName>'),
    ]
    for old, new in fixes:
        assert old in wb, f"expected definedName not found: {old[:70]}"
        wb = wb.replace(old, new, 1)
    # Add Monthly Plan Print_Area (localSheetId 1) once
    if "Monthly Plan'!$B$1" not in wb:
        anchor = '<definedName name="_xlnm.Print_Area" localSheetId="0">Dashboard!$B$1:$I$35</definedName>'
        assert anchor in wb, "Dashboard Print_Area anchor not found"
        add = anchor + f"<definedName name=\"_xlnm.Print_Area\" localSheetId=\"1\">'Monthly Plan'!$B$1:$G${FOOT_ROW}</definedName>"
        wb = wb.replace(anchor, add, 1)
    return wb

def main():
    if not os.path.exists(BACKUP):
        shutil.copy2(WB, BACKUP)
        print("BACKUP created:", BACKUP)
    src = BACKUP

    zin = zipfile.ZipFile(src, "r")
    names = zin.namelist()
    parts = {n: zin.read(n) for n in names}
    infos = {i.filename: i for i in zin.infolist()}
    zin.close()

    parts["xl/styles.xml"] = patch_styles(parts["xl/styles.xml"].decode("utf-8")).encode("utf-8")
    parts["xl/workbook.xml"] = patch_workbook(parts["xl/workbook.xml"].decode("utf-8")).encode("utf-8")
    parts["xl/worksheets/sheet6.xml"] = sheet6_xml().encode("utf-8")

    tmp = WB + ".tmp"
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for n in names:
        if n in infos:
            zout.writestr(infos[n], parts[n])
        else:
            zout.writestr(n, parts[n])
    zout.close()
    print("WROTE TEMP:", tmp)

if __name__ == "__main__":
    main()
