#!/usr/bin/env python3
"""
add_monthly_plan_sheet.py — add a "Monthly Plan" worksheet to the v6 workbook.

A per-employee 30-day trip planner: Month selector (Jan 2026 - Dec 2027) + 40
dated rows (# | Date | From | To | Miles | Notes). Miles auto-calc from the
Matrix via the workbook's named ranges (mtx_vals/mtx_rows/mtx_cols/dir_code/
dir_label). From/To and Month are dropdowns.

Implemented as a NEW worksheet added by surgical XML zip editing so the existing
Dashboard layout, slicer, pivotCache, data-model, Table1 and image are all left
untouched. Text uses inline strings, so sharedStrings.xml is not modified.

Reads/writes: Fleet_Distance_Dashboard_v6.xlsx (in place, via temp; verified
before replacing). v1-v5 originals are never touched.
"""
import zipfile, re, os, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, "Fleet_Distance_Dashboard_v6.xlsx")

ROWS = 40
FIRST = 6                      # first data row
LAST = FIRST + ROWS - 1        # 45
TOTAL_ROW = LAST + 1           # 46
DATE_STYLE = 99                # new cellXfs index we append

MONTHS = [f"{m} {y}" for y in (2026, 2027) for m in
          ["January","February","March","April","May","June","July","August",
           "September","October","November","December"]]

def esc(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def istr(ref, style, text):
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{esc(text)}</t></is></c>'

def num(ref, style, v):
    return f'<c r="{ref}" s="{style}"><v>{v}</v></c>'

def blank(ref, style):
    return f'<c r="{ref}" s="{style}"/>'

def formula(ref, style, f):
    return f'<c r="{ref}" s="{style}"><f>{f}</f></c>'

def build_sheet_xml():
    rows = []
    # row 1 title
    rows.append(f'<row r="1" ht="26" customHeight="1">{istr("A1",72,"MONTHLY TRAVEL PLAN")}'
                + "".join(blank(c+"1",72) for c in "BCDEF") + "</row>")
    # row 2 instructions
    rows.append(f'<row r="2" ht="30" customHeight="1">'
                + istr("A2",73,"Plan one employee's month of trips. Pick the month, then fill each day's Date and From → To. "
                              "Miles auto-fill from the Matrix (road-distance estimates; home↔site legs are OSRM estimates).")
                + "".join(blank(c+"2",73) for c in "BCDEF") + "</row>")
    # row 3 employee + month
    rows.append(f'<row r="3" ht="20" customHeight="1">'
                + istr("A3",4,"Employee") + istr("B3",63,"") + blank("C3",63)
                + istr("D3",4,"Month") + istr("E3",63,"") + blank("F3",63) + "</row>")
    # row 4 spacer
    rows.append('<row r="4" ht="6" customHeight="1"/>')
    # row 5 headers
    rows.append(f'<row r="5" ht="20" customHeight="1">'
                + istr("A5",3,"#") + istr("B5",4,"Date") + istr("C5",4,"From")
                + istr("D5",4,"To") + istr("E5",5,"Miles") + istr("F5",4,"Notes") + "</row>")
    # data rows
    MILES_F = ("IFERROR(INDEX(mtx_vals,"
               "MATCH(INDEX(dir_code,MATCH(C{r},dir_label,0)),mtx_rows,0),"
               "MATCH(INDEX(dir_code,MATCH(D{r},dir_label,0)),mtx_cols,0)),\"\")")
    for i in range(ROWS):
        r = FIRST + i
        rows.append(
            f'<row r="{r}" ht="19" customHeight="1">'
            + num(f"A{r}", 6, i + 1)
            + blank(f"B{r}", DATE_STYLE)
            + blank(f"C{r}", 63)
            + blank(f"D{r}", 63)
            + formula(f"E{r}", 8, MILES_F.format(r=r))
            + blank(f"F{r}", 63)
            + "</row>")
    # total row
    rows.append(f'<row r="{TOTAL_ROW}" ht="20" customHeight="1">'
                + istr(f"A{TOTAL_ROW}",67,"TOTAL PLAN MILES")
                + "".join(blank(c+str(TOTAL_ROW),67) for c in "BCD")
                + formula(f"E{TOTAL_ROW}",62,f"SUM(E{FIRST}:E{LAST})")
                + blank(f"F{TOTAL_ROW}",63) + "</row>")
    # month helper list in hidden column H
    for i, mlabel in enumerate(MONTHS):
        rows.append(f'<row r="{i+1}" hidden="0" outlineLevel="0"><c r="H{i+1}" s="6" t="inlineStr"><is><t>{esc(mlabel)}</t></is></c></row>'
                    if False else "")  # placeholder; merged below
    rows = [x for x in rows if x]

    # We must merge H-column helper cells into the existing rows 1..24 (rows 1..5 exist,
    # 6..24 exist as data rows). Simpler: emit helper cells as separate cells appended
    # into their row elements. Re-build with helper folded in:
    return rows

# Rebuild rows with the month helper folded into existing rows 1..24 (H column)
def build_sheet_data():
    cells_by_row = {}
    def put(r, xml):
        cells_by_row.setdefault(r, []).append((xml,))
    # helper months in H1:H24
    for i, mlabel in enumerate(MONTHS):
        r = i + 1
        put(r, f'<c r="H{r}" s="6" t="inlineStr"><is><t xml:space="preserve">{esc(mlabel)}</t></is></c>')

    # main content cells
    put(1, istr("A1",72,"MONTHLY TRAVEL PLAN")); [put(1, blank(c+"1",72)) for c in "BCDEF"]
    put(2, istr("A2",73,"Plan one employee's month of trips. Pick the month, then fill each day's Date and From → To. "
                        "Miles auto-fill from the Matrix (road estimates; home↔site legs are OSRM estimates).")); [put(2, blank(c+"2",73)) for c in "BCDEF"]
    put(3, istr("A3",4,"Employee")); put(3, istr("B3",63,"")); put(3, blank("C3",63))
    put(3, istr("D3",4,"Month")); put(3, istr("E3",63,"")); put(3, blank("F3",63))
    put(5, istr("A5",3,"#")); put(5, istr("B5",4,"Date")); put(5, istr("C5",4,"From"))
    put(5, istr("D5",4,"To")); put(5, istr("E5",5,"Miles")); put(5, istr("F5",4,"Notes"))
    MILES_F = ("IFERROR(INDEX(mtx_vals,"
               "MATCH(INDEX(dir_code,MATCH(C{r},dir_label,0)),mtx_rows,0),"
               "MATCH(INDEX(dir_code,MATCH(D{r},dir_label,0)),mtx_cols,0)),\"\")")
    for i in range(ROWS):
        r = FIRST + i
        put(r, num(f"A{r}",6,i+1)); put(r, blank(f"B{r}",DATE_STYLE))
        put(r, blank(f"C{r}",63)); put(r, blank(f"D{r}",63))
        put(r, formula(f"E{r}",8,MILES_F.format(r=r))); put(r, blank(f"F{r}",63))
    put(TOTAL_ROW, istr(f"A{TOTAL_ROW}",67,"TOTAL PLAN MILES"))
    [put(TOTAL_ROW, blank(c+str(TOTAL_ROW),67)) for c in "BCD"]
    put(TOTAL_ROW, formula(f"E{TOTAL_ROW}",62,f"SUM(E{FIRST}:E{LAST})"))
    put(TOTAL_ROW, blank(f"F{TOTAL_ROW}",63))

    # heights for special rows
    ht = {1:'ht="26" customHeight="1"', 2:'ht="34" customHeight="1"', 3:'ht="20" customHeight="1"',
          4:'ht="6" customHeight="1"', 5:'ht="20" customHeight="1"', TOTAL_ROW:'ht="20" customHeight="1"'}
    out = []
    col_order = {c:i for i,c in enumerate("ABCDEFGHIJKLMNOP")}
    for r in sorted(cells_by_row):
        cells = [c[0] for c in cells_by_row[r]]
        # sort cells by column letter so refs are ascending (required by Excel)
        def colkey(xml):
            m = re.search(r'r="([A-Z]+)\d+"', xml); return col_order[m.group(1)]
        cells.sort(key=colkey)
        attr = ht.get(r, '')
        out.append(f'<row r="{r}" {attr}>{"".join(cells)}</row>')
    # row 4 spacer if not already present
    if 4 not in cells_by_row:
        out.append('<row r="4" ht="6" customHeight="1"/>')
        out.sort(key=lambda x: int(re.search(r'r="(\d+)"', x).group(1)))
    return "".join(out)

SHEET_XML = None
def sheet6_xml():
    data = build_sheet_data()
    merges = ['A1:F1','A2:F2','B3:C3','E3:F3',f'A{TOTAL_ROW}:D{TOTAL_ROW}']
    mc = f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + '</mergeCells>'
    dv = ('<dataValidations count="3">'
          f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C{FIRST}:C{LAST}"><formula1>dir_label</formula1></dataValidation>'
          f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="D{FIRST}:D{LAST}"><formula1>dir_label</formula1></dataValidation>'
          '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E3"><formula1>plan_months</formula1></dataValidation>'
          '</dataValidations>')
    cols = ('<cols>'
            '<col min="1" max="1" width="5" customWidth="1"/>'
            '<col min="2" max="2" width="13" customWidth="1"/>'
            '<col min="3" max="4" width="34" customWidth="1"/>'
            '<col min="5" max="5" width="9" customWidth="1"/>'
            '<col min="6" max="6" width="34" customWidth="1"/>'
            '<col min="8" max="8" width="14" hidden="1" customWidth="1"/>'
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
        '<sheetPr><tabColor rgb="FF52D1A0"/></sheetPr>'
        f'<dimension ref="A1:H{TOTAL_ROW}"/>'
        '<sheetViews><sheetView showGridLines="0" workbookViewId="0"><selection activeCell="C6" sqref="C6"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
        + cols
        + '<sheetData>' + data + '</sheetData>'
        + mc + dv
        + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        '</worksheet>')

def main():
    zin = zipfile.ZipFile(WB, "r")
    names = zin.namelist()
    parts = {n: zin.read(n) for n in names}
    infos = {i.filename: i for i in zin.infolist()}
    zin.close()

    # 1) styles.xml: append a date cellXf (numFmtId 14) -> index 99
    st = parts["xl/styles.xml"].decode("utf-8")
    m = re.search(r'<cellXfs count="(\d+)">', st)
    cur = int(m.group(1))
    assert cur == DATE_STYLE, f"cellXfs count {cur} != expected {DATE_STYLE}"
    date_xf = ('<xf numFmtId="14" fontId="18" fillId="4" borderId="0" xfId="0" '
               'applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" '
               'applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="center"/>'
               '<protection locked="0"/></xf>')
    st = st.replace(f'<cellXfs count="{cur}">', f'<cellXfs count="{cur+1}">', 1)
    st = st.replace('</cellXfs>', date_xf + '</cellXfs>', 1)
    parts["xl/styles.xml"] = st.encode("utf-8")

    # 2) workbook.xml.rels: add a relationship for sheet6
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    used = set(re.findall(r'Id="rId(\d+)"', rels))
    nid = 1
    while str(nid) in used: nid += 1
    rid = f"rId{nid}"
    newrel = (f'<Relationship Id="{rid}" '
              'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
              'Target="worksheets/sheet6.xml"/>')
    rels = rels.replace('</Relationships>', newrel + '</Relationships>', 1)
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    # 3) workbook.xml: add <sheet> after Dashboard + add plan_months defined name
    wb = parts["xl/workbook.xml"].decode("utf-8")
    sheet_el = f'<sheet name="Monthly Plan" sheetId="6" r:id="{rid}"/>'
    wb2 = re.sub(r'(<sheet name="Dashboard"[^/]*/>)', r'\1' + sheet_el, wb, count=1)
    assert wb2 != wb, "failed to insert <sheet>"
    wb = wb2
    dn = "<definedName name=\"plan_months\">'Monthly Plan'!$H$1:$H$24</definedName>"
    if "<definedNames>" in wb:
        wb = wb.replace("<definedNames>", "<definedNames>" + dn, 1)
    else:
        wb = wb.replace("</sheets>", "</sheets><definedNames>" + dn + "</definedNames>", 1)
    parts["xl/workbook.xml"] = wb.encode("utf-8")

    # 4) [Content_Types].xml: add Override for sheet6
    ct = parts["[Content_Types].xml"].decode("utf-8")
    ov = ('<Override PartName="/xl/worksheets/sheet6.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct = ct.replace('</Types>', ov + '</Types>', 1)
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    # 5) the new sheet
    parts["xl/worksheets/sheet6.xml"] = sheet6_xml().encode("utf-8")

    # write out (temp, then verify externally before replacing)
    tmp = WB + ".tmp"
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    order = names + ["xl/worksheets/sheet6.xml"]
    for n in order:
        if n in infos:
            zout.writestr(infos[n], parts[n])
        else:
            zout.writestr(n, parts[n])
    zout.close()
    print("WROTE TEMP", tmp)
    print(f"  new sheet rId={rid}, sheetId=6, date style index={DATE_STYLE}")

if __name__ == "__main__":
    main()
