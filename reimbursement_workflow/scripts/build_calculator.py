#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_calculator.py  --  build the self-contained HR gas-reimbursement CALCULATOR.

Like the existing Fleet distance-calculator workbook, this produces ONE Excel file
that computes everything with native Excel formulas (INDEX/MATCH, VLOOKUP,
SUMIFS, COUNTIFS, IF) against an EMBEDDED copy of the authoritative Matrix.
There is NO runtime Python and NO macros: HR pastes data, picks locations from
dropdowns, selects a month, and Excel recalculates.

Python is used ONLY here, once, to build (or refresh) the workbook — e.g. to
re-embed the Matrix if the master changes.

Tabs produced:
  Dashboard      month selector + KPIs (SUMIFS/COUNTIFS)   [live]
  Monthly Input  HR pastes trips; per-row formulas resolve codes, miles,
                 status and reimbursement                   [HR + formulas]
  Employees      roster + rates                             [HR]
  Locations      pick-list + alias table (text -> site/HOME/MULTISTOP/OTHER) [HR]
  Monthly Report per-employee totals for the selected month [live]
  Settings       rate, variance threshold, month list       [HR, small]
  MatrixData     embedded authoritative site-to-site miles   [locked backend]

Matrix miles are authoritative. HOME / MULTI-STOP / ambiguous(OTHER) /
unresolved / off-matrix rows are flagged and never auto-paid ($0).

Usage:  python build_calculator.py [--force]
Output: Reimbursement_Workflow/Gas_Reimbursement_Calculator.xlsx
"""

import argparse
import os
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
import process_month as pm

WORKFLOW_DIR = os.path.dirname(SCRIPT_DIR)
PROJECT_DIR = os.path.dirname(WORKFLOW_DIR)
MASTER = os.path.join(PROJECT_DIR, "Fleet_Distance_Dashboard_v5.xlsx")
OUT_PATH = os.path.join(WORKFLOW_DIR, "Gas_Reimbursement_Calculator.xlsx")
SEED_MONTH = "2026-05"
SEED_RAW = os.path.join(WORKFLOW_DIR, "inbox", SEED_MONTH,
                        "2026-05 Gas Reimbursement - Employee A.xlsx")
MI_ROWS = 1000            # Monthly Input data capacity (rows 2..1001)
MI_LAST = MI_ROWS + 1
LOC_LAST = 500
EMP_LAST = 200
EE_PATH = os.path.join(WORKFLOW_DIR, "Gas_Reimbursement_Employee_Edition.xlsx")
EE_DEFAULT_PW = os.getenv("EMPLOYEE_EDITION_PASSWORD", "change-me")   # override via env/.env

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# ---- styling (DARK THEME) ----
DARK_BG = "161B26"       # page background
LIGHT = "E8EAED"          # primary text
MUTE = "A7B0C0"           # secondary text
ACCENT = "8AB4F8"         # headings / links (light blue)
GOLD = "FFC000"
BRAND = "24304A"          # header fill (dark blue)
HFILL = PatternFill("solid", fgColor=BRAND)
HFONT = Font(color="FFFFFF", bold=True)
TITLE = Font(size=15, bold=True, color=ACCENT)
SUB = Font(italic=True, color=MUTE)
LBL = Font(bold=True, color=LIGHT)
LIGHTF = Font(color=LIGHT)
BAND = PatternFill("solid", fgColor=GOLD)         # gold band (dark text)
INP = PatternFill("solid", fgColor="3A3320")      # dark amber = HR types
CALC = PatternFill("solid", fgColor="262B38")     # dark grey = formula/locked
KPI = PatternFill("solid", fgColor="1E3050")
GREEN = PatternFill("solid", fgColor="1E4D2B")    # dark green
AMBER = PatternFill("solid", fgColor="5A4A12")    # dark amber
RED = PatternFill("solid", fgColor="5A1F27")      # dark red
THIN = Side(style="thin", color="3A4152")
BORDER = Border(THIN, THIN, THIN, THIN)

# accent fills that keep DARK text (everything else uses light text)
_LIGHT_FILLS = {GOLD}


def paint_dark(ws, rmax=80, cmax=26):
    """Dark-theme a sheet: hide gridlines, fill empty cells with the page
    background, and switch default (black) text to light — while preserving any
    colour I set explicitly (headers, titles, gold) and dark text on gold."""
    ws.sheet_view.showGridLines = False
    bg = PatternFill("solid", fgColor=DARK_BG)
    for r in range(1, rmax + 1):
        for c in range(1, cmax + 1):
            cell = ws.cell(row=r, column=c)
            pt = cell.fill.patternType if cell.fill else None
            solid = pt == "solid"
            fg = ""
            if solid and cell.fill.fgColor is not None and cell.fill.fgColor.rgb:
                fg = str(cell.fill.fgColor.rgb)[-6:].upper()
            light_bg = fg in _LIGHT_FILLS
            if not solid:
                cell.fill = bg
            if not light_bg and ((cell.value is not None) or solid):
                f = cell.font
                cur = f.color.rgb if (f and f.color is not None and f.color.rgb) else None
                explicit = isinstance(cur, str) and cur[-6:].upper() != "000000"
                if not explicit:
                    cell.font = Font(name=f.name, size=f.size, bold=f.bold,
                                     italic=f.italic, color=LIGHT)


def header(ws, headers, row=1, fill=HFILL):
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = fill; c.font = HFONT
        c.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 28
    ws.freeze_panes = f"A{row+1}"


def widths(ws, m):
    for k, v in m.items():
        ws.column_dimensions[k].width = v


# ------------------------------------------------------------------ master read
def load_master():
    wb = openpyxl.load_workbook(MASTER, data_only=True, read_only=True)
    d = wb["Directory"]; m = wb["Matrix"]
    sites = []   # (code, name, address, lat, lon)
    for row in d.iter_rows(min_row=5, min_col=1, max_col=12, values_only=True):
        code = row[1]
        if code is None or str(code).strip() == "":
            continue
        sites.append((str(code).strip(),
                      "" if row[2] is None else str(row[2]).strip(),
                      "" if row[3] is None else str(row[3]).strip(),
                      row[10], row[11]))
    mrows = list(m.iter_rows(min_col=1, values_only=True))
    wb.close()
    return sites, mrows


# ------------------------------------------------------------------ builders
def build_matrixdata(ws, mrows):
    """Embed the authoritative matrix verbatim (row1 = codes, colA = codes)."""
    for r, row in enumerate(mrows, 1):
        for c, v in enumerate(row, 1):
            if v is None:
                continue
            cell = ws.cell(row=r, column=c, value=(str(v).strip() if r == 1 or c == 1 else v))
            cell.fill = CALC
            cell.font = LIGHTF
    ws.protection.sheet = True   # lock the authoritative data
    ws.sheet_state = "visible"
    # dims
    ncols = max(len(r) for r in mrows)
    nrows = len(mrows)
    return nrows, ncols


def build_locations(ws, sites, seed_aliases):
    header(ws, ["Location Text (type or pick)", "Key (auto)", "Type",
                "Matrix Code", "Verified Address", "Latitude", "Longitude", "Notes"])
    ws["A1"].comment = Comment(
        "Every place that can appear in Monthly Input. Type = SITE (on the Matrix), "
        "HOME, MULTISTOP, or OTHER (needs review). Add new verified addresses here; "
        "they are used automatically next recalculation.", "Calculator")
    r = 2
    code_to_label = {}
    # canonical site rows
    for code, name, addr, lat, lon in sites:
        label = f"{code}  {name}".strip()
        code_to_label[code] = label
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=f"=TRIM(A{r})")
        ws.cell(row=r, column=3, value="SITE")
        ws.cell(row=r, column=4, value=code)
        ws.cell(row=r, column=5, value=addr)
        if isinstance(lat, (int, float)):
            ws.cell(row=r, column=6, value=lat)
        if isinstance(lon, (int, float)):
            ws.cell(row=r, column=7, value=lon)
        r += 1
    # special tokens
    for text, typ, note in [("HOME", "HOME", "Home/apartment leg -> $0 until HR approval."),
                            ("MULTI-STOP", "MULTISTOP", "Several stores in one day -> itemise, $0."),
                            ("OTHER (review)", "OTHER", "Ambiguous/unknown -> review, $0.")]:
        ws.cell(row=r, column=1, value=text)
        ws.cell(row=r, column=2, value=f"=TRIM(A{r})")
        ws.cell(row=r, column=3, value=typ)
        ws.cell(row=r, column=8, value=note)
        r += 1
    # seeded aliases from the sample month (messy text -> type/code)
    for disp, typ, code in seed_aliases:
        ws.cell(row=r, column=1, value=disp)
        ws.cell(row=r, column=2, value=f"=TRIM(A{r})")
        ws.cell(row=r, column=3, value=typ)
        if code:
            ws.cell(row=r, column=4, value=code)
        r += 1
    # fill remaining rows as editable
    for rr in range(2, LOC_LAST + 1):
        for cc in range(1, 9):
            cell = ws.cell(row=rr, column=cc)
            cell.border = BORDER
            cell.font = LIGHTF
            cell.fill = CALC if cc == 2 else INP    # key col is a helper formula
            if cell.value is None and cc == 2:
                cell.value = f"=IF(A{rr}=\"\",\"\",TRIM(A{rr}))"
    DataValidation  # noqa (type dropdown)
    dv = DataValidation(type="list", formula1='"SITE,HOME,MULTISTOP,OTHER"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"C2:C{LOC_LAST}")
    widths(ws, {"A": 34, "B": 26, "C": 12, "D": 12, "E": 40, "F": 11, "G": 11, "H": 44})
    return code_to_label


def build_employees(ws, roster=None):
    header(ws, ["Employee ID", "Employee Name", "Default Rate", "Active", "Notes"])
    if roster:
        for row in roster:
            r = [("" if x is None else x) for x in row]
            if r:                       # force Employee ID to TEXT for reliable lookups
                r[0] = "" if r[0] == "" else str(r[0]).replace(".0", "")
            ws.append(r)
    else:
        ws.append(["1001", "JANE DOE", 0.27, "YES",
                   "Home unconfirmed; pick HOME on Monthly Input. Home legs stay $0 until approval."])
    for rr in range(2, EMP_LAST + 1):
        for cc in range(1, 6):
            ws.cell(row=rr, column=cc).border = BORDER
            ws.cell(row=rr, column=cc).fill = INP
            ws.cell(row=rr, column=cc).font = LIGHTF
        ws.cell(row=rr, column=1).number_format = "@"   # Employee ID as text
    dv = DataValidation(type="list", formula1='"YES,NO"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"D2:D{EMP_LAST}")
    widths(ws, {"A": 14, "B": 26, "C": 13, "D": 9, "E": 58})


def build_settings(ws):
    header(ws, ["Setting", "Value", "Notes"])
    rows = [("Default rate per mile", 0.27, "Used if an employee has no rate."),
            ("Variance threshold %", 20, "Odometer vs Matrix gap above this % -> variance exception.")]
    for k, v, n in rows:
        ws.append([k, v, n])
    # Months list (optional record of months seen)
    ws.cell(row=1, column=5, value="Months").fill = HFILL
    ws.cell(row=1, column=5).font = HFONT
    ws.cell(row=2, column=5, value=SEED_MONTH)
    for rr in range(2, 60):
        ws.cell(row=rr, column=5).border = BORDER
        ws.cell(row=rr, column=5).fill = INP
        ws.cell(row=rr, column=5).font = LIGHTF
    for rr in range(2, 4):
        for cc in range(1, 4):
            ws.cell(row=rr, column=cc).border = BORDER
            ws.cell(row=rr, column=cc).font = LIGHTF
    ws["B2"].fill = INP; ws["B3"].fill = INP
    widths(ws, {"A": 24, "B": 12, "C": 52, "D": 3, "E": 14})
    ws["E1"].comment = Comment("Add a new row here (e.g. 2026-06) when you start a new month.",
                               "Calculator")


def status_formula(row):
    r = row
    return ("=IF($D{r}=\"\",\"\","
            "IF(OR($K{r}=\"HOME\",$L{r}=\"HOME\"),\"HOME_ADDRESS_REVIEW\","
            "IF(OR($K{r}=\"MULTISTOP\",$L{r}=\"MULTISTOP\"),\"MULTI-STOP\","
            "IF(OR($K{r}=\"OTHER\",$L{r}=\"OTHER\",$K{r}=\"UNKNOWN\",$L{r}=\"UNKNOWN\"),\"REVIEW - OTHER\","
            "IF(OR($I{r}=\"\",$J{r}=\"\"),\"EXCEPTION - UNRESOLVED\","
            "IF($M{r}=\"\",\"EXCEPTION - NO MATRIX DISTANCE\","
            "IF(AND($H{r}<>\"\",$M{r}>0,ABS($H{r}-$M{r})/$M{r}*100>Threshold),\"EXCEPTION - VARIANCE\","
            "\"APPROVED\")))))))").format(r=r)


def build_monthly_input(ws, nrows_m, ncols_m, seed_rows):
    heads = ["Month", "Employee ID", "Date", "Origin", "Destination",
             "Odometer Start", "Odometer End", "Claimed Miles",
             "Origin Code", "Dest Code", "Origin Type", "Dest Type",
             "Matrix Miles", "Rate", "Status", "Reimbursement $", "Claimed $"]
    header(ws, heads)
    # comments on the HR-entered columns
    ws["A1"].comment = Comment("Month as YYYY-MM (e.g. 2026-05). Every month lives here.", "Calculator")
    ws["D1"].comment = Comment("Type or pick the origin. Must match a row on Locations.", "Calculator")
    ws["E1"].comment = Comment("Type or pick the destination. Must match a row on Locations.", "Calculator")
    ws["I1"].comment = Comment("Auto: grey columns are formulas — do not edit.", "Calculator")

    lastcol = get_column_letter(ncols_m)      # matrix last col letter
    lastrow = nrows_m
    loc_key = f"Locations!$B$2:$D${LOC_LAST}"  # Key | Type | Code
    for r in range(2, MI_LAST + 1):
        ws.cell(row=r, column=8, value=(
            f"=IF($D{r}=\"\",\"\",IF(AND($F{r}<>\"\",$G{r}<>\"\"),$G{r}-$F{r},\"\"))"))
        ws.cell(row=r, column=9, value=(
            f"=IF($D{r}=\"\",\"\",IFERROR(VLOOKUP(TRIM($D{r}),{loc_key},3,FALSE),\"\"))"))
        ws.cell(row=r, column=10, value=(
            f"=IF($E{r}=\"\",\"\",IFERROR(VLOOKUP(TRIM($E{r}),{loc_key},3,FALSE),\"\"))"))
        ws.cell(row=r, column=11, value=(
            f"=IF($D{r}=\"\",\"\",IFERROR(VLOOKUP(TRIM($D{r}),{loc_key},2,FALSE),\"UNKNOWN\"))"))
        ws.cell(row=r, column=12, value=(
            f"=IF($E{r}=\"\",\"\",IFERROR(VLOOKUP(TRIM($E{r}),{loc_key},2,FALSE),\"UNKNOWN\"))"))
        ws.cell(row=r, column=13, value=(
            f"=IF(OR($D{r}=\"\",$E{r}=\"\",$I{r}=\"\",$J{r}=\"\"),\"\","
            f"IFERROR(INDEX(MatrixData!$B$2:${lastcol}${lastrow},"
            f"MATCH($I{r},MatrixData!$A$2:$A${lastrow},0),"
            f"MATCH($J{r},MatrixData!$B$1:${lastcol}$1,0)),\"\"))"))
        ws.cell(row=r, column=14, value=(
            f"=IF($B{r}=\"\",DefaultRate,IFERROR(VLOOKUP($B{r},Employees!$A$2:$C${EMP_LAST},3,FALSE),DefaultRate))"))
        ws.cell(row=r, column=15, value=status_formula(r))
        ws.cell(row=r, column=16, value=(
            f"=IF($O{r}=\"APPROVED\",$M{r}*$N{r},0)"))
        ws.cell(row=r, column=17, value=(
            f"=IF($D{r}=\"\",\"\",IF($H{r}=\"\",0,$H{r}*$N{r}))"))
        for cc in range(1, len(heads) + 1):
            cell = ws.cell(row=r, column=cc)
            cell.border = BORDER
            cell.fill = INP if cc <= 7 else CALC     # A..G HR types; H.. formulas
            cell.font = LIGHTF
        ws.cell(row=r, column=2).number_format = "@"   # Employee ID as text

    # seed the sample month
    for i, sr in enumerate(seed_rows):
        r = 2 + i
        ws.cell(row=r, column=1, value=sr["month"])
        ws.cell(row=r, column=2, value=sr["emp"])
        ws.cell(row=r, column=3, value=sr["date"])
        ws.cell(row=r, column=4, value=sr["origin"])
        ws.cell(row=r, column=5, value=sr["dest"])
        if sr["odos"] is not None:
            ws.cell(row=r, column=6, value=sr["odos"])
        if sr["odoe"] is not None:
            ws.cell(row=r, column=7, value=sr["odoe"])

    # dropdowns
    dv_loc = DataValidation(type="list", formula1=f"=Locations!$A$2:$A${LOC_LAST}", allow_blank=True)
    ws.add_data_validation(dv_loc); dv_loc.add(f"D2:E{MI_LAST}")
    dv_emp = DataValidation(type="list", formula1=f"=Employees!$A$2:$A${EMP_LAST}", allow_blank=True)
    ws.add_data_validation(dv_emp); dv_emp.add(f"B2:B{MI_LAST}")
    dv_mon = DataValidation(type="list", formula1="=Settings!$E$2:$E$59", allow_blank=True)
    ws.add_data_validation(dv_mon); dv_mon.add(f"A2:A{MI_LAST}")

    # colour the Status column by outcome
    scol = f"O2:O{MI_LAST}"
    ws.conditional_formatting.add(scol, CellIsRule(operator="equal", formula=['"APPROVED"'], fill=GREEN))
    ws.conditional_formatting.add(scol, CellIsRule(operator="equal", formula=['"HOME_ADDRESS_REVIEW"'], fill=AMBER))
    ws.conditional_formatting.add(scol, CellIsRule(operator="equal", formula=['"MULTI-STOP"'], fill=AMBER))
    ws.conditional_formatting.add(scol, CellIsRule(operator="equal", formula=['"REVIEW - OTHER"'], fill=AMBER))
    ws.conditional_formatting.add(scol, FormulaRule(formula=['ISNUMBER(SEARCH("EXCEPTION",O2))'], fill=RED))
    widths(ws, {"A": 9, "B": 11, "C": 11, "D": 26, "E": 26, "F": 13, "G": 13, "H": 12,
                "I": 11, "J": 10, "K": 12, "L": 12, "M": 12, "N": 8, "O": 24, "P": 15, "Q": 12})


def build_dashboard(ws, nsites, lastcol, lastrow):
    """v4-style: brand band, KPI tiles, a quick calculator, and totals-by-employee
    that update automatically as rows are added (no month cell to manage)."""
    ws.sheet_view.showGridLines = False
    site_range = f"Locations!$A$2:$A${1 + nsites}"          # clean site-only picklist
    md = f"MatrixData!$B$2:${lastcol}${lastrow}"
    md_rows = f"MatrixData!$A$2:$A${lastrow}"
    md_cols = f"MatrixData!$B$1:${lastcol}$1"
    loc_key = "Locations!$B$2:$D$500"

    # brand band
    ws.merge_cells("B2:I2")
    b = ws["B2"]; b.value = "●  Fleet OPERATIONS   ·   GAS REIMBURSEMENT"
    b.fill = BAND; b.font = Font(bold=True, size=13, color="7A5C00")
    b.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[2].height = 26
    ws["B3"] = "MONTHLY GAS REIMBURSEMENT CALCULATOR"; ws["B3"].font = TITLE
    ws["B4"] = ("Reusable each month for all employees. Matrix miles are authoritative; "
                "Home / Multi-stop / off-Matrix trips are flagged and never auto-paid.")
    ws["B4"].font = SUB

    # KPI tiles (all entered data)
    tiles = [
        ("B", "EMPLOYEES", '=COUNTA(Employees!$A$2:$A$200)'),
        ("D", "TRIPS LOGGED", '=COUNTA(miEmp)'),
        ("F", "APPROVED $", '=ROUND(SUM(miReimb),2)'),
        ("H", "PENDING REVIEW $", '=ROUND(SUMIFS(miClaimed,miStatus,"<>APPROVED"),2)'),
    ]
    for col, label, f in tiles:
        vc = ws[f"{col}6"]; vc.value = f
        vc.font = Font(bold=True, size=18, color=GOLD); vc.fill = KPI
        vc.alignment = Alignment(horizontal="center"); vc.border = BORDER
        lc = ws[f"{col}7"]; lc.value = label
        lc.font = Font(bold=True, size=9, color=MUTE); lc.fill = KPI
        lc.alignment = Alignment(horizontal="center"); lc.border = BORDER

    # 01 quick calculator
    ws["B10"] = "01   QUICK REIMBURSEMENT CALCULATOR"; ws["B10"].font = Font(bold=True, color=ACCENT)
    ws["B11"] = "Pick an employee, a From site and a To site — see the reimbursement instantly."
    ws["B11"].font = SUB
    def field(r, label, val=None, formula=None, fill=INP):
        ws.cell(row=r, column=2, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=3, value=(formula if formula else val))
        c.border = BORDER; c.fill = fill
        return c
    field(12, "Employee ID:", val="1001")
    ws["C12"].number_format = "@"    # keep the typed ID as text so lookups match
    ws.cell(row=12, column=4, value='=IF($C$12="","",IFERROR(VLOOKUP($C$12,Employees!$A$2:$B$200,2,FALSE),"?"))').font = SUB
    field(13, "Rate / mile:",
          formula='=IF($C$12="",DefaultRate,IFERROR(VLOOKUP($C$12,Employees!$A$2:$C$200,3,FALSE),DefaultRate))',
          fill=CALC)
    field(14, "From site:", val="")
    field(15, "To site:", val="")
    field(16, "Miles (Matrix):",
          formula=(f'=IFERROR(INDEX({md},'
                   f'MATCH(IFERROR(VLOOKUP(TRIM($C$14),{loc_key},3,FALSE),""),{md_rows},0),'
                   f'MATCH(IFERROR(VLOOKUP(TRIM($C$15),{loc_key},3,FALSE),""),{md_cols},0)),"")'),
          fill=CALC)
    rc = field(17, "Reimbursement $:", formula='=IF(OR($C$16="",$C$13=""),"",ROUND($C$16*$C$13,2))', fill=GREEN)
    rc.font = Font(bold=True)
    dv_emp = DataValidation(type="list", formula1=f"=Employees!$A$2:$A${EMP_LAST}", allow_blank=True)
    ws.add_data_validation(dv_emp); dv_emp.add("C12")
    dv_site = DataValidation(type="list", formula1=f"={site_range}", allow_blank=True)
    ws.add_data_validation(dv_site); dv_site.add("C14:C15")

    # 02 totals by employee (auto — grows with the roster / data)
    ws["B20"] = "02   TOTALS BY EMPLOYEE   (all entered trips)"; ws["B20"].font = Font(bold=True, color=ACCENT)
    tot_hdr = ["Employee", "Emp ID", "Legs", "Claimed Miles", "Approved $", "Pending $"]
    header(ws, tot_hdr, row=21)
    ws.freeze_panes = "A22"
    first = 22
    n = 20
    for i in range(n):
        r = first + i
        er = 2 + i
        ws.cell(row=r, column=2, value=f'=IF(Employees!$A{er}="","",Employees!$B{er})')
        ws.cell(row=r, column=3, value=f'=IF(Employees!$A{er}="","",Employees!$A{er})')
        base = 'miEmp,$C{r}'.format(r=r)
        ws.cell(row=r, column=4, value=f'=IF($C{r}="","",COUNTIF({base}))')
        ws.cell(row=r, column=5, value=f'=IF($C{r}="","",ROUND(SUMIFS(miMiles,{base}),1))')
        ws.cell(row=r, column=6, value=f'=IF($C{r}="","",ROUND(SUMIFS(miReimb,{base}),2))')
        ws.cell(row=r, column=7, value=f'=IF($C{r}="","",ROUND(SUMIFS(miClaimed,{base},miStatus,"<>APPROVED"),2))')
        for cc in range(2, 8):
            ws.cell(row=r, column=cc).border = BORDER
    tr = first + n
    ws.cell(row=tr, column=2, value="COMPANY TOTAL").font = Font(bold=True)
    ws.cell(row=tr, column=4, value='=COUNTA(miEmp)').font = Font(bold=True)
    ws.cell(row=tr, column=5, value='=ROUND(SUM(miMiles),1)').font = Font(bold=True)
    ws.cell(row=tr, column=6, value='=ROUND(SUM(miReimb),2)').font = Font(bold=True)
    ws.cell(row=tr, column=7, value='=ROUND(SUMIFS(miClaimed,miStatus,"<>APPROVED"),2)').font = Font(bold=True)
    for cc in range(2, 8):
        ws.cell(row=tr, column=cc).border = BORDER

    ws.cell(row=tr + 2, column=2, value="Add trips on Monthly Input — totals here update automatically "
            "(press F9 if needed). See the How to Use tab.").font = SUB
    widths(ws, {"A": 3, "B": 22, "C": 20, "D": 14, "E": 14, "F": 13, "G": 13, "H": 16, "I": 4})


def build_monthly_report(ws):
    ws["A1"] = "Monthly Report — all entered trips, by employee"; ws["A1"].font = TITLE
    header(ws, ["Employee", "Emp ID", "Legs", "Approved", "Needs Review", "Exceptions",
                "Matrix Miles (appr)", "Reimbursement $", "Claimed $", "Pending $"], row=2)
    for i in range(20):   # up to 20 employees
        r = 3 + i
        er = 2 + i        # Employees row
        ws.cell(row=r, column=1, value=f'=IF(Employees!$A{er}="","",Employees!$B{er})')
        ws.cell(row=r, column=2, value=f'=IF(Employees!$A{er}="","",Employees!$A{er})')
        base = 'miEmp,$B{r}'.format(r=r)
        ws.cell(row=r, column=3, value=f'=IF($B{r}="","",COUNTIF({base}))')
        ws.cell(row=r, column=4, value=f'=IF($B{r}="","",COUNTIFS({base},miStatus,"APPROVED"))')
        ws.cell(row=r, column=5, value=(
            f'=IF($B{r}="","",COUNTIFS({base},miStatus,"HOME_ADDRESS_REVIEW")'
            f'+COUNTIFS({base},miStatus,"MULTI-STOP")+COUNTIFS({base},miStatus,"REVIEW - OTHER"))'))
        ws.cell(row=r, column=6, value=f'=IF($B{r}="","",COUNTIFS({base},miStatus,"EXCEPTION*"))')
        ws.cell(row=r, column=7, value=f'=IF($B{r}="","",ROUND(SUMIFS(miMatrix,{base},miStatus,"APPROVED"),1))')
        ws.cell(row=r, column=8, value=f'=IF($B{r}="","",ROUND(SUMIFS(miReimb,{base}),2))')
        ws.cell(row=r, column=9, value=f'=IF($B{r}="","",ROUND(SUMIFS(miClaimed,{base}),2))')
        ws.cell(row=r, column=10, value=f'=IF($B{r}="","",ROUND(SUMIFS(miClaimed,{base},miStatus,"<>APPROVED"),2))')
        for cc in range(1, 11):
            ws.cell(row=r, column=cc).border = BORDER
    widths(ws, {"A": 24, "B": 10, "C": 8, "D": 10, "E": 13, "F": 12,
                "G": 17, "H": 16, "I": 12, "J": 12})


def build_howto(ws):
    ws.sheet_view.showGridLines = False
    ws["B2"] = "HOW TO USE — Fleet Gas Reimbursement Calculator"; ws["B2"].font = TITLE
    blocks = [
        ("What this is", [
            "One self-contained Excel file. It calculates mileage and reimbursement with the",
            "company Matrix built in. No add-ins, no macros, nothing to install.",
            "Use it as the TEMPLATE each month, for all employees. Two easy ways to work:",
            "  • Keep one file and keep adding rows (all months stay in it), OR",
            "  • Save a copy per month (e.g. 'Gas Reimbursement 2026-06.xlsx') and enter that month.",
            "Either way the Dashboard totals reflect whatever trips are in the file.",
        ]),
        ("Each month — 3 steps", [
            "1) MONTHLY INPUT tab: add one row per trip.",
            "     Month, Employee ID, Date, Origin, Destination, Odometer Start, Odometer End.",
            "     Pick Origin/Destination from the dropdown (site name, or HOME / MULTI-STOP / OTHER).",
            "2) Press F9 to recalculate (or leave Excel on automatic calculation).",
            "3) Read the DASHBOARD and MONTHLY REPORT — totals update automatically.",
        ]),
        ("Add a new employee", [
            "EMPLOYEES tab: add a row — Employee ID, Name, Default Rate (e.g. 0.27), Active = YES.",
            "That Employee ID then appears in the Monthly Input and calculator dropdowns.",
        ]),
        ("Add a new location / verified address", [
            "LOCATIONS tab: add a row.",
            "  • Location Text  = how it will be typed/picked (e.g. a store name or 'west store').",
            "  • Type = SITE (a real company site on the Matrix), HOME, MULTISTOP, or OTHER.",
            "  • Matrix Code = the site code (e.g. 0083) for a SITE — this is what makes miles compute.",
            "  • Verified Address / Latitude / Longitude = reference only (optional).",
            "New rows are usable immediately in the Origin/Destination dropdowns.",
            "Tip: to confirm an employee's HOME, add a row Type=HOME with their address.",
        ]),
        ("What the Status column means", [
            "APPROVED           - both ends are real Matrix sites, one stop, mileage sensible -> pays Matrix miles x rate.",
            "HOME_ADDRESS_REVIEW- a home/apartment end -> $0 until HR approves.",
            "MULTI-STOP         - several stores in one day -> itemise manually, $0.",
            "REVIEW - OTHER     - ambiguous or not yet known (e.g. two possible cities) -> pick the exact site, $0.",
            "EXCEPTION - ...    - unresolved text / no Matrix distance / odometer far from Matrix -> $0, check the row.",
        ]),
        ("Rules (do not change)", [
            "Matrix road miles are the official distance. Odometer is only a cross-check.",
            "A location that is not on the Matrix is never auto-paid.",
            "The MatrixData tab is the embedded official matrix and is locked — do not edit it.",
            "To refresh the built-in matrix after the master changes, IT runs build_calculator.py.",
        ]),
    ]
    r = 4
    for title, lines in blocks:
        ws.cell(row=r, column=2, value=title).font = Font(bold=True, size=12, color=ACCENT)
        r += 1
        for ln in lines:
            ws.cell(row=r, column=2, value=ln).font = SUB
            r += 1
        r += 1
    widths(ws, {"A": 3, "B": 110})


# ------------------------------------------------------------------ seed helpers
def classify(norm, verified):
    if pm.has_personal(norm):
        return "HOME", ""
    if pm.has_multistop_phrase(norm):
        return "MULTISTOP", ""
    if norm in verified:
        return "SITE", verified[norm]
    return "OTHER", ""


def _fmt_date(v):
    try:
        return v.strftime("%Y-%m-%d")
    except AttributeError:
        s = str(v).split(" ")[0] if v not in (None, "") else ""
        return s


def monthly_from_sample():
    """Read the May sample file into monthly rows."""
    rows = []
    if not os.path.exists(SEED_RAW):
        print(f"[warn] sample file not found; Monthly Input left empty: {SEED_RAW}")
        return rows
    ef = pm.parse_input(SEED_RAW, {"master_workbook": "x"})
    for t in ef.trips:
        rows.append({"month": SEED_MONTH, "emp": ef.emp_id or "1001", "date": _fmt_date(t.date),
                     "origin": t.origin_raw, "dest": t.dest_raw,
                     "odos": t.odo_start, "odoe": t.odo_end})
    return rows


def derive_aliases(monthly):
    """Build the Locations alias rows from whatever text appears in Monthly Input."""
    verified = {"123 demo blvd": "0075", "west store": "0083"}
    seen = {}
    for row in monthly:
        for raw in (row["origin"], row["dest"]):
            n = pm.normalize(raw)
            if not n or n in seen:
                continue
            typ, code = classify(n, verified)
            seen[n] = ((raw or "").strip(), typ, code)
    return list(seen.values())


def read_existing():
    """Preserve HR's data: read Employees + Monthly Input (raw A-G) + Settings
    from the current calculator. Returns dict or None if the file is absent."""
    if not os.path.exists(OUT_PATH):
        return None
    wb = openpyxl.load_workbook(OUT_PATH)
    emp = []
    if "Employees" in wb.sheetnames:
        for r in wb["Employees"].iter_rows(min_row=2, min_col=1, max_col=5, values_only=True):
            if r[0] not in (None, ""):
                emp.append([("" if x is None else x) for x in r])
    monthly = []
    if "Monthly Input" in wb.sheetnames:
        for r in wb["Monthly Input"].iter_rows(min_row=2, min_col=1, max_col=7, values_only=True):
            if all(x in (None, "") for x in r):
                continue    # skip blank rows (compacts any gaps)
            monthly.append({"month": ("" if r[0] is None else str(r[0]).strip()),
                            "emp": ("" if r[1] is None else str(r[1]).strip().replace(".0", "")),
                            "date": _fmt_date(r[2]),
                            "origin": ("" if r[3] is None else str(r[3])),
                            "dest": ("" if r[4] is None else str(r[4])),
                            "odos": r[5], "odoe": r[6]})
    settings = {}
    if "Settings" in wb.sheetnames:
        s = wb["Settings"]
        settings["rate"] = s["B2"].value
        settings["threshold"] = s["B3"].value
    wb.close()
    return {"employees": emp, "monthly": monthly, "settings": settings}


def add_names(wb):
    def nm(name, ref):
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    mi = "'Monthly Input'"
    nm("miMonth", f"{mi}!$A$2:$A${MI_LAST}")
    nm("miEmp", f"{mi}!$B$2:$B${MI_LAST}")
    nm("miMatrix", f"{mi}!$M$2:$M${MI_LAST}")
    nm("miStatus", f"{mi}!$O$2:$O${MI_LAST}")
    nm("miReimb", f"{mi}!$P$2:$P${MI_LAST}")
    nm("miClaimed", f"{mi}!$Q$2:$Q${MI_LAST}")
    nm("miMiles", f"{mi}!$H$2:$H${MI_LAST}")
    nm("DefaultRate", "Settings!$B$2")
    nm("Threshold", "Settings!$B$3")


def protect_employee_edition(wb, password):
    """Lock everything except the travel-entry columns (Monthly Input A-G), hide
    the backend tabs, and password-protect every sheet. Dropdowns still work in
    the unlocked cells; formulas and other tabs become read-only."""
    mi = wb["Monthly Input"]
    unlocked = Protection(locked=False)
    for r in range(2, MI_LAST + 1):
        for c in range(1, 8):          # A..G = Month, Emp, Date, Origin, Dest, Odo start/end
            mi.cell(row=r, column=c).protection = unlocked
    for ws in wb.worksheets:
        ws.protection.password = password
        ws.protection.sheet = True
        ws.protection.selectLockedCells = True     # can view/click, not edit
        ws.protection.selectUnlockedCells = True
        ws.protection.formatCells = False
    for name in ("MatrixData", "Locations", "Employees", "Settings"):
        if name in wb.sheetnames:
            wb[name].sheet_state = "hidden"
    wb.active = wb.sheetnames.index("Dashboard")


def build(existing=None, empty_monthly=False, protect=False, out_path=None, password=None):
    sites, mrows = load_master()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_dash = wb.create_sheet("Dashboard")
    ws_howto = wb.create_sheet("How to Use")
    ws_in = wb.create_sheet("Monthly Input")
    ws_emp = wb.create_sheet("Employees")
    ws_loc = wb.create_sheet("Locations")
    ws_rep = wb.create_sheet("Monthly Report")
    ws_set = wb.create_sheet("Settings")
    ws_mat = wb.create_sheet("MatrixData")

    nrows_m, ncols_m = build_matrixdata(ws_mat, mrows)
    lastcol = get_column_letter(ncols_m)

    # Data source: preserve HR's data if rebuilding, else the May sample.
    if empty_monthly:
        monthly = []                                   # blank sheet for employees to fill
        roster = (existing["employees"] if existing else None) or None
    elif existing and existing["monthly"]:
        monthly = existing["monthly"]
        roster = existing["employees"] or None
    else:
        monthly = monthly_from_sample()
        roster = None
    # aliases still come from the full sample so the dropdowns are populated
    alias_rows = derive_aliases(monthly if monthly else monthly_from_sample())

    build_locations(ws_loc, sites, alias_rows)
    build_employees(ws_emp, roster)
    build_settings(ws_set)
    if existing and existing.get("settings"):
        s = existing["settings"]
        if s.get("rate") is not None:
            ws_set["B2"] = s["rate"]
        if s.get("threshold") is not None:
            ws_set["B3"] = s["threshold"]
    add_names(wb)     # names must exist before formulas reference them
    build_monthly_input(ws_in, nrows_m, ncols_m, monthly)
    build_dashboard(ws_dash, len(sites), lastcol, nrows_m)
    build_monthly_report(ws_rep)
    build_howto(ws_howto)

    # dark-theme the text-on-background tabs (data tabs are already dark-filled)
    paint_dark(ws_dash, rmax=46, cmax=10)
    paint_dark(ws_howto, rmax=62, cmax=4)
    paint_dark(ws_rep, rmax=26, cmax=11)
    for s in (ws_in, ws_emp, ws_loc, ws_set, ws_mat):
        s.sheet_view.showGridLines = False
        s.sheet_properties.tabColor = BRAND
    for s in (ws_dash, ws_howto, ws_rep):
        s.sheet_properties.tabColor = BRAND

    if protect:
        protect_employee_edition(wb, password or EE_DEFAULT_PW)

    wb.active = wb.sheetnames.index("Dashboard")
    wb.calculation.fullCalcOnLoad = True
    target = out_path or OUT_PATH
    out = target
    try:
        wb.save(target)
    except PermissionError:
        base, ext = os.path.splitext(target)
        out = base + " (new)" + ext
        wb.save(out)
        print("[warn] Target file is open in Excel (locked), so I wrote a NEW file instead:")
    print(f"[ok] Wrote {out}")
    print(f"[ok] Embedded matrix {nrows_m-1}x{ncols_m-1}, {len(sites)} sites, "
          f"{len(alias_rows)} aliases, {len(monthly)} trip rows, "
          f"{len(roster) if roster else 1} employees"
          f"{', PROTECTED employee edition' if protect else ''}.")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true", help="Overwrite existing calculator (fresh sample).")
    ap.add_argument("--rebuild", action="store_true",
                    help="Rebuild in place, PRESERVING current Employees + Monthly Input data.")
    ap.add_argument("--employee-edition", action="store_true",
                    help="Build a protected, blank copy for employees (only travel columns editable).")
    ap.add_argument("--password", default=EE_DEFAULT_PW, help="Password for the employee edition.")
    args = ap.parse_args()

    if args.employee_edition:
        existing = read_existing()   # carry the roster so ID dropdowns are populated
        out = build(existing=existing, empty_monthly=True, protect=True,
                    out_path=EE_PATH, password=args.password)
        print(f"[ok] Employee edition password: {args.password}")
        return

    if args.rebuild:
        existing = read_existing()
        if existing is None:
            sys.exit("No existing calculator to rebuild; run without --rebuild first.")
        # back up before overwriting
        import shutil, datetime
        bdir = os.path.join(WORKFLOW_DIR, "backups")
        os.makedirs(bdir, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = os.path.join(bdir, f"Calculator_backup_{ts}.xlsx")
        shutil.copy2(OUT_PATH, bak)
        print(f"[ok] Backed up existing calculator -> {os.path.relpath(bak, WORKFLOW_DIR)}")
        print(f"[ok] Preserving {len(existing['employees'])} employees, "
              f"{len(existing['monthly'])} trip rows.")
        build(existing=existing)
        return

    if os.path.exists(OUT_PATH) and not args.force:
        sys.exit(f"[skip] {OUT_PATH} exists. Use --rebuild (keep data) or --force (fresh sample).")
    build()


if __name__ == "__main__":
    main()
