# -*- coding: utf-8 -*-
"""Shared layout constants for the single HR console workbook.

Both build_console.py (creates the workbook) and process_console.py (fills the
output tabs) import this so the tab names, column headers, and Dashboard anchor
cells never drift out of sync.
"""

CONSOLE_FILENAME = "Gas_Reimbursement_Console.xlsx"

# ---- sheet names ----
SH_DASHBOARD = "Dashboard"
SH_INPUT = "Monthly Input"
SH_EMPLOYEES = "Employees"
SH_LOCATIONS = "Locations"
SH_SETTINGS = "Settings"
SH_REPORT = "Monthly Report"
SH_TRIPS = "Trip Detail"
SH_EXCEPTIONS = "Exceptions"
SH_AUDIT = "Audit Trail"

# Order in which the tabs appear left-to-right.
SHEET_ORDER = [SH_DASHBOARD, SH_INPUT, SH_EMPLOYEES, SH_LOCATIONS, SH_SETTINGS,
               SH_REPORT, SH_TRIPS, SH_EXCEPTIONS, SH_AUDIT]

# Tabs the engine OWNS and fully rebuilds on every run.
ENGINE_OUTPUT_SHEETS = [SH_REPORT, SH_TRIPS, SH_EXCEPTIONS, SH_AUDIT]
# Tabs HR maintains (never rebuilt/cleared by the engine).
HR_INPUT_SHEETS = [SH_INPUT, SH_EMPLOYEES, SH_LOCATIONS, SH_SETTINGS]

# ---- input headers ----
INPUT_HEADERS = ["Month (YYYY-MM)", "Employee ID", "Date", "Origin", "Destination",
                 "Notes", "Odometer Start", "Odometer End", "Claimed Miles"]

EMPLOYEE_HEADERS = ["Employee ID", "Employee Name", "Default Rate", "Active", "Notes"]

LOCATION_HEADERS = ["Location ID", "Employee ID", "Type",
                    "Aliases (as in log; separate with ;)", "Verified Address",
                    "Latitude", "Longitude", "Matrix Site Code", "HR Approved", "Notes"]

# ---- settings: (policy_key, label, default, description) ----
SETTINGS_ROWS = [
    ("default_rate_per_mile", "Default rate per mile", 0.27,
     "Used only if an employee has no rate on the Employees tab."),
    ("variance_threshold_pct", "Variance threshold %", 20,
     "Claimed vs Matrix gap above this % -> flagged as a variance exception."),
    ("fuzzy_match_cutoff", "Fuzzy match cutoff", 0.72,
     "Minimum name similarity (0-1) to accept an automatic match."),
    ("fuzzy_match_margin", "Fuzzy match margin", 0.06,
     "Top candidate must beat the 2nd by this much, else AMBIGUOUS."),
    ("pay_basis", "Pay basis", "matrix",
     "matrix = reimburse on authoritative Matrix road miles."),
    ("master_workbook", "Master workbook", "Fleet_Distance_Dashboard_v5.xlsx",
     "Authoritative Directory/Matrix file (backend, read-only)."),
    ("master_matrix_sheet", "Master matrix sheet", "Matrix", "Sheet name in the master."),
    ("master_directory_sheet", "Master directory sheet", "Directory", "Sheet name in the master."),
    ("home_leg_policy", "Home leg policy", "flag", "Home legs -> HOME_ADDRESS_REVIEW, $0."),
    ("multistop_policy", "Multi-stop policy", "flag", "Multi-store -> manual itemisation, $0."),
]

# ---- output headers ----
REPORT_HEADERS = ["Employee", "Emp ID", "Legs", "Approved", "Needs Review", "Exceptions",
                  "Matrix Miles (approved)", "Reimbursement $ (auto)", "Claimed $ (all)",
                  "Pending $ (review/exc)"]

TRIP_HEADERS = ["Month", "Employee", "Emp ID", "Date", "Origin (raw)", "Destination (raw)",
                "Notes", "Origin Code", "Dest Code", "Route Supported?", "Matrix Miles",
                "Ref Miles (straight-line)", "Claimed Miles", "Variance %", "Rate",
                "Reimbursable $ (matrix)", "Claimed $", "Status",
                "Origin Addr (verified)", "Dest Addr (verified)", "Reason"]

AUDIT_HEADERS = ["Item", "Value"]

# ---- Dashboard anchor cells (engine reads/writes these exact locations) ----
DASH_SELECTED_MONTH_CELL = "C3"     # HR picks the month here (dropdown)
DASH_LAST_RUN_CELL = "C4"           # engine stamps the run time
DASH_KPI_LABEL_COL = "B"
DASH_KPI_VALUE_COL = "C"
DASH_KPI_START_ROW = 7
DASH_KPI_KEYS = ["Total legs", "Employees", "Auto-approved", "Needs review",
                 "Exceptions", "Auto-approved $", "Pending $"]
DASH_HISTORY_HEADER_ROW = 16        # "All months" table header
DASH_HISTORY_DATA_ROW = 17          # first data row of the history table
DASH_HISTORY_HEADERS = ["Month", "Legs", "Approved", "Review", "Exceptions",
                        "Auto $", "Pending $"]
DASH_HISTORY_MAX_ROWS = 120         # engine clears this many rows before rewriting
