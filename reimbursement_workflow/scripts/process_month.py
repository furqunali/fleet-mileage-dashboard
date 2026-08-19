#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
process_month.py  --  Fleet monthly gas-reimbursement processor.

Reusable, config-driven engine that reads every employee's monthly
"Gas Reimbursement" workbook from a month folder in inbox/, matches the messy
free-text locations against the authoritative Fleet site directory + distance
matrix, cross-checks claimed odometer mileage against the matrix road miles,
and writes ONE combined review workbook to output/<month>/.

Design rules (locked with the client -- do not relax without approval):
  1. Matrix road miles are the authoritative reimbursable distance.
     Claimed odometer miles are a CROSS-CHECK only; a large gap -> VARIANCE.
  2. Home / apartment legs are ALWAYS flagged for HR. Never auto-approved.
  3. Multi-store rows are ALWAYS flagged for manual itemization.
     Route order is NEVER guessed.
  4. Ambiguous locations (a city name that maps to >1 site, e.g. Sampleville,
     Democenter) are flagged AMBIGUOUS -- never auto-picked.

Originals (the master workbook and every inbox file) are opened READ-ONLY and
are never modified. Only files under output/<month>/ are written.

Usage:
    python process_month.py --month 2026-05
    python process_month.py --month 2026-05 --root "..\\"     # custom root
    python process_month.py                                    # newest month folder in inbox/

No third-party deps beyond openpyxl. Config lives in ../config/.
"""

import argparse
import csv
import datetime
import difflib
import os
import re
import sys

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required:  pip install openpyxl")

# Make stdout tolerant of the em-dash / arrow characters in the source data.
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


# --------------------------------------------------------------------------
# Paths / config loading
# --------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKFLOW_DIR = os.path.dirname(SCRIPT_DIR)          # Reimbursement_Workflow/
PROJECT_DIR = os.path.dirname(WORKFLOW_DIR)          # Houston_Distance_Dashboard/
CONFIG_DIR = os.path.join(WORKFLOW_DIR, "config")
INBOX_DIR = os.path.join(WORKFLOW_DIR, "inbox")
OUTPUT_DIR = os.path.join(WORKFLOW_DIR, "output")


def load_policy():
    import json
    with open(os.path.join(CONFIG_DIR, "policy.json"), encoding="utf-8") as fh:
        return json.load(fh)


def load_employees():
    """employee_id (str) -> dict of employee fields."""
    path = os.path.join(CONFIG_DIR, "employees.csv")
    out = {}
    if not os.path.exists(path):
        return out
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            eid = (row.get("employee_id") or "").strip()
            if eid:
                out[eid] = row
    return out


def load_aliases():
    """normalized raw_pattern -> (site_code, site_name)."""
    path = os.path.join(CONFIG_DIR, "location_aliases.csv")
    out = {}
    if not os.path.exists(path):
        return out
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            pat = normalize(row.get("raw_pattern", ""))
            code = (row.get("site_code") or "").strip()
            if pat and code:
                out[pat] = (code, (row.get("site_name") or "").strip())
    return out


HR_SETTINGS_FILE = "HR_Settings.xlsx"


class HRLocation:
    """One verified location HR maintains in HR_Settings.xlsx (Locations tab)."""
    __slots__ = ("loc_id", "emp_id", "type", "aliases", "address",
                 "lat", "lon", "matrix_code", "approved", "notes")

    def __init__(self):
        self.loc_id = ""
        self.emp_id = ""          # employee id, or "ALL"
        self.type = "OTHER"       # HOME | WORK | OTHER
        self.aliases = []         # list of normalized alias strings
        self.address = ""
        self.lat = None
        self.lon = None
        self.matrix_code = ""     # optional link to a Directory site code
        self.approved = False
        self.notes = ""

    @property
    def coord(self):
        if self.lat is not None and self.lon is not None:
            return (self.lat, self.lon)
        return None


def _to_float(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def load_hr_settings():
    """Read HR_Settings.xlsx if present.

    Returns (employees_dict, hr_locations_list) or (None, None) if absent.
    employees_dict: employee_id -> {employee_name, default_rate, active, notes}.
    """
    path = os.path.join(CONFIG_DIR, HR_SETTINGS_FILE)
    if not os.path.exists(path):
        return None, None
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

    employees = {}
    if "Employees" in wb.sheetnames:
        ws = wb["Employees"]
        rows = list(ws.iter_rows(min_row=2, min_col=1, max_col=5, values_only=True))
        for r in rows:
            eid = "" if r[0] is None else str(r[0]).strip()
            if eid.endswith(".0"):
                eid = eid[:-2]
            if not eid or eid.lower().startswith("(example"):
                continue
            employees[eid] = {
                "employee_id": eid,
                "employee_name": "" if r[1] is None else str(r[1]).strip(),
                "default_rate": _to_float(r[2]),
                "active": (str(r[3]).strip().upper() != "NO") if r[3] is not None else True,
                "notes": "" if r[4] is None else str(r[4]).strip(),
            }

    locations = []
    if "Locations" in wb.sheetnames:
        ws = wb["Locations"]
        rows = list(ws.iter_rows(min_row=2, min_col=1, max_col=10, values_only=True))
        for r in rows:
            loc_id = "" if r[0] is None else str(r[0]).strip()
            if not loc_id or loc_id.lower().startswith("(example"):
                continue  # blank or example rows are ignored
            loc = HRLocation()
            loc.loc_id = loc_id
            loc.emp_id = ("" if r[1] is None else str(r[1]).strip()) or "ALL"
            if loc.emp_id.endswith(".0"):
                loc.emp_id = loc.emp_id[:-2]
            loc.type = ("OTHER" if r[2] is None else str(r[2]).strip().upper()) or "OTHER"
            raw_aliases = "" if r[3] is None else str(r[3])
            loc.aliases = [normalize(a) for a in raw_aliases.split(";") if normalize(a)]
            loc.address = "" if r[4] is None else str(r[4]).strip()
            loc.lat = _to_float(r[5])
            loc.lon = _to_float(r[6])
            loc.matrix_code = "" if r[7] is None else str(r[7]).strip()
            loc.approved = (str(r[8]).strip().upper() == "YES") if r[8] is not None else False
            loc.notes = "" if r[9] is None else str(r[9]).strip()
            locations.append(loc)

    wb.close()
    return employees, locations


def haversine_miles(a, b):
    """Great-circle distance in miles between (lat,lon) tuples a and b.
    This is a straight-line REFERENCE only; Matrix road miles stay authoritative."""
    import math
    if not a or not b:
        return None
    lat1, lon1 = map(math.radians, a)
    lat2, lon2 = map(math.radians, b)
    dlat, dlon = lat2 - lat1, lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return round(2 * 3958.7613 * math.asin(math.sqrt(h)), 1)


# --------------------------------------------------------------------------
# Text normalization helpers
# --------------------------------------------------------------------------

def normalize(s):
    """Lowercase, strip, collapse internal whitespace. None-safe."""
    if s is None:
        return ""
    s = str(s).replace(" ", " ")          # NBSP -> space
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s


# Personal / commute keywords -> ALWAYS flagged (rule 2).
PERSONAL_KEYWORDS = ("home", "demo apt", "apart", "aparmen", "aparment", "residence")

# Multi-stop signals -> ALWAYS flagged (rule 3).
MULTISTOP_PHRASES = (" and ", "all stores", "all store", "& ")


def has_personal(norm):
    return any(k in norm for k in PERSONAL_KEYWORDS)


def has_multistop_phrase(norm):
    if any(p in norm for p in MULTISTOP_PHRASES):
        return True
    # plural "stores" (with or without trailing space) implies more than one store
    if re.search(r"\bstores\b", norm):
        return True
    return False


# --------------------------------------------------------------------------
# Master directory + matrix
# --------------------------------------------------------------------------

class MasterData:
    """Authoritative site directory + distance matrix, loaded read-only."""

    def __init__(self, policy):
        self.policy = policy
        path = os.path.join(PROJECT_DIR, policy["master_workbook"])
        if not os.path.exists(path):
            sys.exit(f"Master workbook not found: {path}")
        self.path = path
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

        self.code_to_name = {}       # "0075" -> "Demo Store North"
        self.name_to_code = {}       # normalized name -> code
        self.label_to_code = {}      # normalized dropdown label -> code
        self.city_to_codes = {}      # "sampleville" -> ["0025", "0078"]
        self.code_to_coord = {}      # "0075" -> (lat, lon)
        self.code_to_address = {}    # "0075" -> "501 E-Business 190, ..."
        self.codes = []

        self._load_directory(wb[policy["master_directory_sheet"]])
        self.matrix = {}             # (from_code, to_code) -> miles
        self._load_matrix(wb[policy["master_matrix_sheet"]])
        wb.close()

    def _load_directory(self, ws):
        # With min_col=1 (column A = index 0) the Directory layout is:
        #   B(1)=CODE  C(2)=NAME  D(3)=ADDRESS  E(4)=TYPE  F(5)=HQ MILES
        #   H(7)=LABEL  K(10)=LAT  L(11)=LON
        # min_col=1 pins these indices; read-only mode otherwise trims
        # leading empty columns and shifts everything.
        for row in ws.iter_rows(min_row=5, min_col=1, max_col=12, values_only=True):
            code = row[1]
            name = row[2]
            address = row[3]
            label = row[7] if len(row) > 7 else None
            if code is None or str(code).strip() == "":
                continue
            code = str(code).strip()
            name = "" if name is None else str(name).strip()
            self.codes.append(code)
            self.code_to_name[code] = name
            if address:
                self.code_to_address[code] = str(address).strip()
            lat = row[10] if len(row) > 10 else None
            lon = row[11] if len(row) > 11 else None
            if isinstance(lat, (int, float)) and isinstance(lon, (int, float)):
                self.code_to_coord[code] = (float(lat), float(lon))
            if name:
                self.name_to_code[normalize(name)] = code
            if label:
                self.label_to_code[normalize(label)] = code
            city = self._city_from_address(address)
            if city:
                self.city_to_codes.setdefault(city, []).append(code)

    @staticmethod
    def _city_from_address(address):
        """Best-effort city extraction: the token(s) just before the state.

        Handles both 'City, State' and 'City State' forms, e.g.
          '123 Demo Blvd, Democity, Texas, 76522' -> 'democity'
          '789 Example Hwy, Sampleville, TX 70003'              -> 'sampleville'
        """
        if not address:
            return ""
        a = str(address).replace(" ", " ").strip()
        # Split into segments on commas first (most reliable when present).
        m = re.search(r",\s*([A-Za-z .]+?)\s*,\s*(TX|Texas)\b", a, re.IGNORECASE)
        if m:
            return normalize(m.group(1))
        # Fallback: word(s) immediately before 'TX' / 'Texas'.
        m = re.search(r"([A-Za-z]+(?:\s+[A-Za-z]+)?)\s+(?:TX|Texas)\b", a, re.IGNORECASE)
        if m:
            phrase = normalize(m.group(1)).split()
            # Drop a leading street/direction word so 'frwy houston' -> 'houston'
            # while genuine two-word cities ('metro city', 'democity') survive.
            street_words = {"blvd", "frwy", "freeway", "belt", "north", "south",
                            "east", "west", "ave", "avenue", "rd", "road", "hwy",
                            "st", "dr", "ln", "pkwy"}
            if len(phrase) == 2 and phrase[0] in street_words:
                phrase = phrase[1:]
            return " ".join(phrase)
        return ""

    def _load_matrix(self, ws):
        # min_col=1 keeps the FROM/TO label in column A and codes from B on,
        # regardless of read-only leading-column trimming.
        rows = list(ws.iter_rows(min_col=1, values_only=True))
        if not rows:
            return
        header = rows[0]
        col_codes = {}
        for ci, val in enumerate(header):
            if ci == 0 or val is None:
                continue
            col_codes[ci] = str(val).strip()
        for r in rows[1:]:
            if not r or r[0] is None:
                continue
            from_code = str(r[0]).strip()
            if not from_code:
                continue
            for ci, to_code in col_codes.items():
                if ci < len(r) and isinstance(r[ci], (int, float)):
                    self.matrix[(from_code, to_code)] = float(r[ci])

    def distance(self, from_code, to_code):
        if from_code == to_code:
            return 0.0
        if (from_code, to_code) in self.matrix:
            return self.matrix[(from_code, to_code)]
        if (to_code, from_code) in self.matrix:   # matrix is meant to be symmetric
            return self.matrix[(to_code, from_code)]
        return None

    def city_keywords_in(self, norm):
        """Return sorted list of directory cities whose name appears in `norm`."""
        found = []
        for city in self.city_to_codes:
            # multi-word cities: match on the full city phrase or its last word
            if city and (city in norm):
                found.append(city)
            else:
                last = city.split()[-1] if city else ""
                if last and len(last) > 3 and re.search(r"\b" + re.escape(last) + r"\b", norm):
                    found.append(city)
        return sorted(set(found))


# --------------------------------------------------------------------------
# Location resolution
# --------------------------------------------------------------------------

class Resolution:
    """Result of resolving one free-text location string."""
    __slots__ = ("raw", "norm", "status", "code", "name", "detail", "candidates",
                 "coord", "address", "hr_source")

    def __init__(self, raw):
        self.raw = raw
        self.norm = normalize(raw)
        # MATCHED | PERSONAL | MULTISTOP | OFFMATRIX | AMBIGUOUS | UNMATCHED
        self.status = "UNMATCHED"
        self.code = None
        self.name = None
        self.detail = ""
        self.candidates = []
        self.coord = None          # (lat, lon) reference coordinate if known
        self.address = ""          # verified address if known
        self.hr_source = None      # the HRLocation that matched, if any


def _match_hr_location(n, emp_id, hr_locations, policy):
    """Return the HRLocation whose alias list contains `n` (EXACT normalized
    match), own-scope rows preferred over ALL-scope. Matching is deliberately
    exact-only: fuzzy matching here is unsafe because it can cross-match two
    different stores (e.g. 'region-a store' -> 'west store') and auto-pay the
    wrong site. HR controls the alias list, so they add spellings explicitly."""
    if not hr_locations:
        return None
    scoped = [(0, loc) if loc.emp_id == emp_id else (1, loc)
              for loc in hr_locations if loc.emp_id in (emp_id, "ALL")]
    for _, loc in sorted(scoped, key=lambda t: t[0]):
        if n in loc.aliases:
            return loc
    return None


def _attach_site(r, master, code):
    """Attach a Directory site's coordinate + address to a resolution (for the
    straight-line reference distance and display)."""
    if not r.coord:
        r.coord = master.code_to_coord.get(code)
    if not r.address:
        r.address = master.code_to_address.get(code, "")


def resolve_location(raw, master, aliases, policy, emp_id=None, hr_locations=None):
    r = Resolution(raw)
    n = r.norm

    if not n:
        r.status = "UNMATCHED"
        r.detail = "empty location"
        return r

    # 0) HR-maintained verified location (Locations tab). HR has explicitly
    #    configured these, so they take priority over generic matching.
    hr = _match_hr_location(n, emp_id, hr_locations, policy)
    if hr is not None:
        r.hr_source = hr
        r.address = hr.address
        r.coord = hr.coord
        if hr.type == "HOME":
            # Home stays a personal leg: HOME_ADDRESS_REVIEW + $0 downstream.
            r.status = "PERSONAL"
            r.detail = f"HR home location {hr.loc_id}"
            return r
        # WORK / OTHER linked to a real Directory site -> Matrix authoritative.
        if hr.matrix_code and hr.matrix_code in master.code_to_name:
            r.status = "MATCHED"
            r.code = hr.matrix_code
            r.name = master.code_to_name.get(hr.matrix_code, hr.loc_id)
            r.coord = r.coord or master.code_to_coord.get(hr.matrix_code)
            r.detail = f"HR location {hr.loc_id} -> site {hr.matrix_code}"
            return r
        # WORK / OTHER with coordinates but NOT on the authoritative Matrix.
        r.status = "OFFMATRIX"
        r.name = hr.loc_id
        r.detail = f"HR location {hr.loc_id} (verified coords, off-Matrix)"
        return r

    # Rule 2 -- personal/home/apartment always flagged.
    if has_personal(n):
        r.status = "PERSONAL"
        r.detail = "home/apartment/personal keyword"
        return r

    # Rule 3 -- explicit multi-stop phrasing always flagged.
    if has_multistop_phrase(n):
        r.status = "MULTISTOP"
        r.detail = "multi-store phrasing"
        return r

    # 1) verified alias (exact normalized match)
    if n in aliases:
        code, name = aliases[n]
        r.status = "MATCHED"
        r.code = code
        r.name = name or master.code_to_name.get(code, "")
        r.detail = "alias"
        _attach_site(r, master, code)
        return r

    # 2) exact site code (e.g. "0075", "0001-o")
    for code in master.codes:
        if n == code.lower():
            r.status = "MATCHED"
            r.code = code
            r.name = master.code_to_name.get(code, "")
            r.detail = "exact code"
            _attach_site(r, master, code)
            return r

    # 3) exact site name or dropdown label
    if n in master.name_to_code:
        code = master.name_to_code[n]
        r.status, r.code, r.name, r.detail = "MATCHED", code, master.code_to_name.get(code, ""), "exact name"
        _attach_site(r, master, code)
        return r
    if n in master.label_to_code:
        code = master.label_to_code[n]
        r.status, r.code, r.name, r.detail = "MATCHED", code, master.code_to_name.get(code, ""), "exact label"
        _attach_site(r, master, code)
        return r

    # 4) city keyword(s). More than one distinct city -> multistop/ambiguous.
    cities = master.city_keywords_in(n)
    if len(cities) > 1:
        r.status = "MULTISTOP"
        r.detail = "multiple cities: " + ", ".join(cities)
        r.candidates = [c for city in cities for c in master.city_to_codes[city]]
        return r
    if len(cities) == 1:
        codes = master.city_to_codes[cities[0]]
        if len(codes) == 1:
            code = codes[0]
            r.status, r.code, r.name, r.detail = "MATCHED", code, master.code_to_name.get(code, ""), f"city:{cities[0]}"
            _attach_site(r, master, code)
            return r
        # City maps to several sites -> ambiguous by design (rule 4).
        r.status = "AMBIGUOUS"
        r.detail = f"city '{cities[0]}' has {len(codes)} sites"
        r.candidates = codes[:]
        return r

    # 5) fuzzy match against site names (difflib), with cutoff + margin.
    cutoff = float(policy.get("fuzzy_match_cutoff", 0.72))
    margin = float(policy.get("fuzzy_match_margin", 0.06))
    scored = []
    for name_norm, code in master.name_to_code.items():
        score = difflib.SequenceMatcher(None, n, name_norm).ratio()
        scored.append((score, code, name_norm))
    scored.sort(reverse=True)
    if scored and scored[0][0] >= cutoff:
        top = scored[0]
        second = scored[1][0] if len(scored) > 1 else 0.0
        if (top[0] - second) >= margin:
            code = top[1]
            r.status, r.code, r.name = "MATCHED", code, master.code_to_name.get(code, "")
            r.detail = f"fuzzy {top[0]:.2f}"
            _attach_site(r, master, code)
            return r
        r.status = "AMBIGUOUS"
        r.detail = f"fuzzy tie {top[0]:.2f} vs {second:.2f}"
        r.candidates = [top[1], scored[1][1]]
        return r

    # 6) nothing matched.
    r.status = "UNMATCHED"
    r.detail = "no alias/code/name/city/fuzzy match"
    if scored:
        r.candidates = [scored[0][1]]  # best (rejected) guess, for the reviewer
    return r


# --------------------------------------------------------------------------
# Input parsing
# --------------------------------------------------------------------------

class Trip:
    def __init__(self):
        self.date = None
        self.origin_raw = ""
        self.dest_raw = ""
        self.notes = ""
        self.odo_start = None
        self.odo_end = None
        self.claimed_miles = None
        self.claimed_reimb = None


class EmployeeFile:
    def __init__(self, filename):
        self.filename = filename
        self.emp_name = ""
        self.emp_id = ""
        self.rate = None
        self.period = ""
        self.trips = []
        self.parse_warnings = []


def _find_label_value(ws_rows, label_substrings):
    """Scan the first ~12 rows for a cell containing any label substring;
    return the next non-empty cell to its right."""
    for row in ws_rows[:12]:
        for ci, cell in enumerate(row):
            if cell is None:
                continue
            txt = normalize(cell)
            if any(sub in txt for sub in label_substrings):
                for cj in range(ci + 1, len(row)):
                    if row[cj] is not None and str(row[cj]).strip() != "":
                        return row[cj]
    return None


def parse_input(path, policy):
    ef = EmployeeFile(os.path.basename(path))
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    # Pick the mileage-log sheet (name may vary slightly month to month).
    sheet = None
    for name in wb.sheetnames:
        if "mileage" in name.lower() or "log" in name.lower() or "expense" in name.lower():
            sheet = wb[name]
            break
    if sheet is None:
        sheet = wb[wb.sheetnames[0]]
    rows = list(sheet.iter_rows(values_only=True))

    # --- header block ---
    ef.emp_name = str(_find_label_value(rows, ["employee name"]) or "").strip()
    ef.emp_id = str(_find_label_value(rows, ["employee id"]) or "").strip()
    if ef.emp_id.endswith(".0"):
        ef.emp_id = ef.emp_id[:-2]
    rate_val = _find_label_value(rows, ["rate per mile", "rate/mile"])
    try:
        ef.rate = float(rate_val) if rate_val is not None else None
    except (TypeError, ValueError):
        ef.rate = None
    ef.period = str(_find_label_value(rows, ["for period", "period"]) or "").strip()

    # --- locate the log header row (has "date" and "destination") ---
    header_idx = None
    for ri, row in enumerate(rows):
        cells = [normalize(c) for c in row]
        if any(c == "date" for c in cells) and any("destination" in c for c in cells):
            header_idx = ri
            break
    if header_idx is None:
        ef.parse_warnings.append("Could not locate log header row (Date/Destination).")
        wb.close()
        return ef

    header = rows[header_idx]
    norm_header = [normalize(c) for c in header]
    date_col = norm_header.index("date")
    # Layout is positional from the Date column:
    #   Date | Origin | Destination | Notes | OdoStart | OdoEnd | Mileage | Reimbursement
    c_origin = date_col + 1
    c_dest = date_col + 2
    c_notes = date_col + 3
    c_odos = date_col + 4
    c_odoe = date_col + 5
    c_miles = date_col + 6
    c_reimb = date_col + 7

    def cell(row, idx):
        return row[idx] if idx < len(row) else None

    for row in rows[header_idx + 1:]:
        d = cell(row, date_col)
        origin = cell(row, c_origin)
        dest = cell(row, c_dest)
        # Stop at the first fully-empty log line (e.g. the totals row has no date).
        if (d is None or str(d).strip() == "") and \
           (origin is None or str(origin).strip() == "") and \
           (dest is None or str(dest).strip() == ""):
            continue  # skip blanks rather than hard-stop (robust to stray gaps)
        if d is None or str(d).strip() == "":
            continue
        t = Trip()
        t.date = d
        t.origin_raw = "" if origin is None else str(origin).strip()
        t.dest_raw = "" if dest is None else str(dest).strip()
        notes = cell(row, c_notes)
        t.notes = "" if notes is None else str(notes).strip()
        t.odo_start = _num(cell(row, c_odos))
        t.odo_end = _num(cell(row, c_odoe))
        t.claimed_miles = _num(cell(row, c_miles))
        t.claimed_reimb = _num(cell(row, c_reimb))
        ef.trips.append(t)

    wb.close()
    return ef


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


# --------------------------------------------------------------------------
# Processing one trip
# --------------------------------------------------------------------------

# Row-level statuses used across the report.
ST_APPROVED = "APPROVED"
ST_REVIEW_HOME = "HOME_ADDRESS_REVIEW"
ST_REVIEW_MULTISTOP = "MULTI-STOP - MANUAL ITEMIZATION"
ST_EXC_AMBIGUOUS = "EXCEPTION - AMBIGUOUS LOCATION"
ST_EXC_UNMATCHED = "EXCEPTION - UNMATCHED LOCATION"
ST_EXC_VARIANCE = "EXCEPTION - MILEAGE VARIANCE"
ST_EXC_NODIST = "EXCEPTION - NO MATRIX DISTANCE"
ST_REVIEW_OFFMATRIX = "REVIEW - OFF-MATRIX LOCATION"

REVIEW_STATUSES = {ST_REVIEW_HOME, ST_REVIEW_MULTISTOP, ST_REVIEW_OFFMATRIX}
EXCEPTION_STATUSES = {ST_EXC_AMBIGUOUS, ST_EXC_UNMATCHED, ST_EXC_VARIANCE, ST_EXC_NODIST}


class ResultRow:
    def __init__(self):
        self.emp_name = ""
        self.emp_id = ""
        self.date = ""
        self.origin_raw = ""
        self.dest_raw = ""
        self.notes = ""
        self.origin_code = ""
        self.origin_name = ""
        self.dest_code = ""
        self.dest_name = ""
        self.match_detail = ""
        self.matrix_miles = None
        self.claimed_miles = None
        self.variance_pct = None
        self.rate = None
        self.reimbursable = None       # APPROVED = matrix$, HOME/MULTISTOP = 0, else None
        self.claimed_reimb = None
        self.status = ""
        self.reason = ""
        self.ref_miles = None          # straight-line reference (never authoritative)
        self.origin_addr = ""
        self.dest_addr = ""


def process_trip(ef, trip, master, aliases, policy, rate, hr_locations=None):
    rr = ResultRow()
    rr.emp_name = ef.emp_name
    rr.emp_id = ef.emp_id
    rr.date = _fmt_date(trip.date)
    rr.origin_raw = trip.origin_raw
    rr.dest_raw = trip.dest_raw
    rr.notes = trip.notes
    rr.claimed_miles = trip.claimed_miles
    rr.claimed_reimb = trip.claimed_reimb
    rr.rate = rate

    o = resolve_location(trip.origin_raw, master, aliases, policy, ef.emp_id, hr_locations)
    d = resolve_location(trip.dest_raw, master, aliases, policy, ef.emp_id, hr_locations)

    rr.origin_code = o.code or ""
    rr.origin_name = o.name or ""
    rr.dest_code = d.code or ""
    rr.dest_name = d.name or ""
    rr.origin_addr = o.address or ""
    rr.dest_addr = d.address or ""
    rr.match_detail = f"O:{o.status}({o.detail}) | D:{d.status}({d.detail})"

    # Straight-line reference distance if both endpoints have coordinates.
    # This is a REFERENCE aid for HR only; Matrix road miles stay authoritative.
    rr.ref_miles = haversine_miles(o.coord, d.coord)

    # --- priority of flags (rules 2 & 3 win over everything) ---
    statuses = {o.status, d.status}

    if "PERSONAL" in statuses:
        # Home/apartment leg: flagged for HR, matrix miles NOT applied,
        # reimbursement forced to $0 by policy (even after the address is filled).
        rr.status = ST_REVIEW_HOME
        rr.matrix_miles = None
        rr.reimbursable = 0.0
        home_note = ""
        home_res = o if o.status == "PERSONAL" else d
        if home_res.address:
            home_note = f" Verified home on file: {home_res.address}."
        elif home_res.hr_source is not None:
            home_note = " Home configured in HR Settings but address not yet filled."
        ref = f" Ref straight-line ~{rr.ref_miles} mi." if rr.ref_miles is not None else ""
        rr.reason = ("Home/apartment endpoint -> HOME_ADDRESS_REVIEW, $0 until HR approval."
                     + home_note + ref + " " + _leg_reason("home/personal", o, d))
        return rr

    if "MULTISTOP" in statuses:
        # Multiple stores on one line: require manual itemization. Route order
        # is never guessed, so NO automatic distance/reimbursement -> $0.
        rr.status = ST_REVIEW_MULTISTOP
        rr.matrix_miles = None
        rr.reimbursable = 0.0
        rr.reason = ("Multiple stores on one line -> requires manual itemization; "
                     "route order not guessed, no auto distance, $0 by policy. "
                     + _leg_reason("multi-stop", o, d))
        return rr

    if "OFFMATRIX" in statuses:
        # A verified HR location that is not on the authoritative Matrix.
        # Show the straight-line reference distance, but do NOT auto-pay -> $0.
        rr.status = ST_REVIEW_OFFMATRIX
        rr.matrix_miles = None
        rr.reimbursable = 0.0
        ref = f" Ref straight-line ~{rr.ref_miles} mi." if rr.ref_miles is not None else ""
        rr.reason = ("Verified HR location not on the official Matrix -> review; "
                     "no auto Matrix distance, $0 until confirmed." + ref
                     + f"  [O:{o.detail} | D:{d.detail}]")
        return rr

    if "AMBIGUOUS" in statuses:
        rr.status = ST_EXC_AMBIGUOUS
        rr.reason = _leg_reason("ambiguous", o, d, master)
        return rr

    if "UNMATCHED" in statuses:
        rr.status = ST_EXC_UNMATCHED
        rr.reason = _leg_reason("unmatched", o, d)
        return rr

    # Both endpoints cleanly MATCHED -> compute matrix distance.
    dist = master.distance(o.code, d.code)
    if dist is None:
        rr.status = ST_EXC_NODIST
        rr.reason = f"No matrix distance between {o.code} and {d.code}."
        return rr
    rr.matrix_miles = dist

    # Variance cross-check (matrix is authoritative; odometer is the check).
    if trip.claimed_miles and dist > 0:
        var = abs(trip.claimed_miles - dist) / dist * 100.0
        rr.variance_pct = round(var, 1)
        if var > float(policy.get("variance_threshold_pct", 20.0)):
            rr.status = ST_EXC_VARIANCE
            rr.reimbursable = round(dist * rate, 2)   # still pay matrix miles, but flag
            rr.reason = (f"Claimed {trip.claimed_miles:g} mi vs matrix {dist:g} mi "
                         f"({rr.variance_pct:g}% > {policy.get('variance_threshold_pct')}%).")
            return rr

    # Clean, in-tolerance, non-personal, single-stop leg -> auto-approve.
    rr.status = ST_APPROVED
    rr.reimbursable = round(dist * rate, 2)
    rr.reason = ""
    return rr


def _leg_reason(kind, o, d, master=None):
    parts = []
    for side, res in (("Origin", o), ("Dest", d)):
        tag = ""
        if kind == "home/personal" and res.status == "PERSONAL":
            tag = "home/personal"
        elif kind == "multi-stop" and res.status == "MULTISTOP":
            tag = res.detail
        elif kind == "ambiguous" and res.status == "AMBIGUOUS":
            cand = ", ".join(res.candidates)
            tag = f"{res.detail} (candidates: {cand})"
        elif kind == "unmatched" and res.status == "UNMATCHED":
            tag = "not found in directory"
        if tag:
            parts.append(f"{side} '{res.raw}': {tag}")
    return " ; ".join(parts) if parts else kind


def _fmt_date(v):
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).split(" ")[0]


# --------------------------------------------------------------------------
# Report writing
# --------------------------------------------------------------------------

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(size=14, bold=True, color="1F4E78")
BAND_FILL = PatternFill("solid", fgColor="FFC000")
STATUS_FILLS = {
    ST_APPROVED: PatternFill("solid", fgColor="C6EFCE"),
    ST_REVIEW_HOME: PatternFill("solid", fgColor="FFEB9C"),
    ST_REVIEW_MULTISTOP: PatternFill("solid", fgColor="FFEB9C"),
    ST_REVIEW_OFFMATRIX: PatternFill("solid", fgColor="FFEB9C"),
    ST_EXC_AMBIGUOUS: PatternFill("solid", fgColor="FFC7CE"),
    ST_EXC_UNMATCHED: PatternFill("solid", fgColor="FFC7CE"),
    ST_EXC_VARIANCE: PatternFill("solid", fgColor="FFC7CE"),
    ST_EXC_NODIST: PatternFill("solid", fgColor="FFC7CE"),
}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _write_header(ws, headers, row=1):
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.border = BORDER
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def _autosize(ws, max_w=48):
    for col in ws.columns:
        length = 0
        letter = None
        for c in col:
            if letter is None:
                letter = c.column_letter
            if c.value is not None:
                length = max(length, len(str(c.value)))
        if letter:
            ws.column_dimensions[letter].width = min(max(10, length + 2), max_w)


def build_report(all_rows, employee_files, master, policy, month, out_path, run_ts,
                 config_source="CSV", hr_locations=None):
    wb = openpyxl.Workbook()

    # ---------- Summary ----------
    ws = wb.active
    ws.title = "Summary"
    ws.cell(row=1, column=1, value=f"Fleet Gas Reimbursement — {month}").font = TITLE_FONT
    ws.cell(row=2, column=1, value=f"Generated {run_ts}  |  Master: {policy['master_workbook']}  |  Pay basis: matrix road miles")
    for cc in range(1, 9):
        ws.cell(row=4, column=cc).fill = BAND_FILL

    sum_hdr = ["Employee", "Emp ID", "Legs", "Approved", "Needs Review",
               "Exceptions", "Approved $ (auto)", "Claimed $ (all legs)", "Pending $ (review/exc)"]
    _write_header(ws, sum_hdr, row=5)

    r = 6
    grand = [0, 0, 0, 0, 0.0, 0.0, 0.0]
    for ef in employee_files:
        rows = [x for x in all_rows if x.emp_id == ef.emp_id and x._file == ef.filename]
        legs = len(rows)
        approved = sum(1 for x in rows if x.status == ST_APPROVED)
        review = sum(1 for x in rows if x.status in REVIEW_STATUSES)
        exc = sum(1 for x in rows if x.status in EXCEPTION_STATUSES)
        appr_amt = sum(x.reimbursable for x in rows if x.status == ST_APPROVED and x.reimbursable)
        claimed_amt = sum(x.claimed_reimb for x in rows if x.claimed_reimb)
        pending_amt = sum((x.claimed_reimb or 0) for x in rows if x.status != ST_APPROVED)
        vals = [ef.emp_name or "(unknown)", ef.emp_id, legs, approved, review, exc,
                round(appr_amt, 2), round(claimed_amt, 2), round(pending_amt, 2)]
        for ci, v in enumerate(vals, start=1):
            ws.cell(row=r, column=ci, value=v).border = BORDER
        grand[0] += legs; grand[1] += approved; grand[2] += review; grand[3] += exc
        grand[4] += appr_amt; grand[5] += claimed_amt; grand[6] += pending_amt
        r += 1
    # totals
    tcell = ws.cell(row=r, column=1, value="TOTAL")
    tcell.font = Font(bold=True)
    for ci, v in enumerate(["", grand[0], grand[1], grand[2], grand[3],
                            round(grand[4], 2), round(grand[5], 2), round(grand[6], 2)], start=2):
        c = ws.cell(row=r, column=ci, value=v)
        c.font = Font(bold=True)
        c.border = BORDER
    _autosize(ws)

    # ---------- Trip Detail ----------
    ws = wb.create_sheet("Trip Detail")
    detail_hdr = ["File", "Employee", "Emp ID", "Date", "Origin (raw)", "Destination (raw)",
                  "Notes", "Origin Code", "Dest Code", "Matrix Miles",
                  "Ref Miles (straight-line)", "Claimed Miles", "Variance %", "Rate",
                  "Reimbursable $ (matrix)", "Claimed $", "Status",
                  "Origin Addr (verified)", "Dest Addr (verified)", "Reason / Match detail"]
    status_col = detail_hdr.index("Status") + 1
    _write_header(ws, detail_hdr)

    def detail_row(x, with_match_detail):
        reason = x.reason
        if with_match_detail and x.match_detail:
            reason = (reason + "  ||  " + x.match_detail).strip()
        return [x._file, x.emp_name, x.emp_id, x.date, x.origin_raw, x.dest_raw, x.notes,
                x.origin_code, x.dest_code, x.matrix_miles, x.ref_miles, x.claimed_miles,
                x.variance_pct, x.rate, x.reimbursable, x.claimed_reimb, x.status,
                x.origin_addr, x.dest_addr, reason]

    for x in all_rows:
        ws.append(detail_row(x, with_match_detail=True))
        rr = ws.max_row
        fill = STATUS_FILLS.get(x.status)
        if fill:
            ws.cell(row=rr, column=status_col).fill = fill
        for cc in range(1, len(detail_hdr) + 1):
            ws.cell(row=rr, column=cc).border = BORDER
    _autosize(ws)

    # ---------- Exceptions (everything not auto-approved) ----------
    ws = wb.create_sheet("Exceptions")
    _write_header(ws, detail_hdr)
    for x in all_rows:
        if x.status == ST_APPROVED:
            continue
        ws.append(detail_row(x, with_match_detail=False))
        rr = ws.max_row
        fill = STATUS_FILLS.get(x.status)
        if fill:
            ws.cell(row=rr, column=status_col).fill = fill
        for cc in range(1, len(detail_hdr) + 1):
            ws.cell(row=rr, column=cc).border = BORDER
    _autosize(ws)

    # ---------- Unmatched / Ambiguous locations (unique) ----------
    ws = wb.create_sheet("Unmatched Locations")
    _write_header(ws, ["Location (raw)", "Normalized", "Reason", "Suggested candidates",
                        "Times seen", "Example file"])
    agg = {}
    for x in all_rows:
        # Scan EVERY row's endpoints (not just exception rows): an ambiguous
        # store hidden behind a flagged home leg is still vocabulary HR may
        # want to add to location_aliases.csv.
        for raw in (x.origin_raw, x.dest_raw):
            res = resolve_location(raw, master, load_aliases_cached(), policy)
            if res.status in ("AMBIGUOUS", "UNMATCHED"):
                key = res.norm
                if key not in agg:
                    cand = ", ".join(f"{c} {master.code_to_name.get(c,'')}".strip()
                                     for c in res.candidates)
                    agg[key] = {"raw": raw, "reason": f"{res.status}: {res.detail}",
                                "cand": cand, "count": 0, "file": x._file}
                agg[key]["count"] += 1
    for key, info in sorted(agg.items()):
        ws.append([info["raw"], key, info["reason"], info["cand"], info["count"], info["file"]])
        for cc in range(1, 7):
            ws.cell(row=ws.max_row, column=cc).border = BORDER
    if not agg:
        ws.append(["(none)", "", "", "", "", ""])
    _autosize(ws)

    # ---------- Audit Trail ----------
    ws = wb.create_sheet("Audit Trail")
    _write_header(ws, ["Item", "Value"])
    audit = [
        ("Run timestamp", run_ts),
        ("Month processed", month),
        ("Master workbook", policy["master_workbook"]),
        ("Pay basis", policy.get("pay_basis")),
        ("Default rate / mile", policy.get("default_rate_per_mile")),
        ("Variance threshold %", policy.get("variance_threshold_pct")),
        ("Fuzzy cutoff / margin", f"{policy.get('fuzzy_match_cutoff')} / {policy.get('fuzzy_match_margin')}"),
        ("Home leg policy", policy.get("home_leg_policy")),
        ("Multi-stop policy", policy.get("multistop_policy")),
        ("Config source", config_source),
        ("HR verified locations loaded", len(hr_locations) if hr_locations else 0),
        ("  of which HR-approved", sum(1 for l in (hr_locations or []) if l.approved)),
        ("Directory sites loaded", len(master.codes)),
        ("Directory sites with coordinates", len(master.code_to_coord)),
        ("Matrix pairs loaded", len(master.matrix)),
        ("Files processed", len(employee_files)),
    ]
    for loc in (hr_locations or []):
        audit.append((f"  HR loc {loc.loc_id}",
                      f"emp {loc.emp_id}, {loc.type}, approved={'YES' if loc.approved else 'NO'}, "
                      f"addr={'(set)' if loc.address else '(blank)'}, "
                      f"coords={'(set)' if loc.coord else '(blank)'}, "
                      f"matrix_code={loc.matrix_code or '-'}"))
    for ef in employee_files:
        rate_note = f"rate {ef.rate}" if ef.rate is not None else "rate MISSING (used policy default)"
        warn = ("; ".join(ef.parse_warnings)) if ef.parse_warnings else "ok"
        audit.append((f"  • {ef.filename}",
                      f"{ef.emp_name} (ID {ef.emp_id}), {len(ef.trips)} legs, {ef.period}, {rate_note} [{warn}]"))
    for item, val in audit:
        ws.append([item, val])
        for cc in range(1, 3):
            ws.cell(row=ws.max_row, column=cc).border = BORDER
    _autosize(ws, max_w=90)

    wb.save(out_path)


# Small cache so the Unmatched tab can re-resolve without reloading the CSV.
_ALIAS_CACHE = None
def load_aliases_cached():
    global _ALIAS_CACHE
    if _ALIAS_CACHE is None:
        _ALIAS_CACHE = load_aliases()
    return _ALIAS_CACHE


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def pick_month(root_inbox, month_arg):
    if month_arg:
        return month_arg
    if not os.path.isdir(root_inbox):
        sys.exit(f"No inbox directory: {root_inbox}")
    subs = sorted(d for d in os.listdir(root_inbox)
                  if os.path.isdir(os.path.join(root_inbox, d)))
    if not subs:
        sys.exit("No month folders found in inbox/.")
    return subs[-1]


def main():
    ap = argparse.ArgumentParser(description="Fleet monthly gas-reimbursement processor.")
    ap.add_argument("--month", help="Month folder under inbox/ (e.g. 2026-05). Default: newest.")
    args = ap.parse_args()

    policy = load_policy()
    aliases = load_aliases_cached()

    # Primary config source: HR_Settings.xlsx (HR-maintained). Fall back to CSV.
    hr_employees, hr_locations = load_hr_settings()
    if hr_employees is not None:
        config_source = HR_SETTINGS_FILE
        employees = hr_employees
    else:
        config_source = "employees.csv"
        employees = load_employees()
        hr_locations = []
    month = pick_month(INBOX_DIR, args.month)

    in_dir = os.path.join(INBOX_DIR, month)
    if not os.path.isdir(in_dir):
        sys.exit(f"Month folder not found: {in_dir}")
    out_dir = os.path.join(OUTPUT_DIR, month)
    os.makedirs(out_dir, exist_ok=True)

    print(f"[info] Config source: {config_source} "
          f"({len(employees)} employees, {len(hr_locations)} HR verified locations).")
    print(f"[info] Loading master data from {policy['master_workbook']} ...")
    master = MasterData(policy)
    print(f"[info] {len(master.codes)} sites, {len(master.matrix)} matrix pairs, "
          f"{len(master.city_to_codes)} cities, {len(master.code_to_coord)} with coords.")

    input_files = sorted(f for f in os.listdir(in_dir)
                         if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"))
    if not input_files:
        sys.exit(f"No input .xlsx files in {in_dir}")

    all_rows = []
    employee_files = []
    default_rate = float(policy.get("default_rate_per_mile", 0.27))

    for fname in input_files:
        path = os.path.join(in_dir, fname)
        print(f"[info] Parsing {fname} ...")
        ef = parse_input(path, policy)

        roster = employees.get(ef.emp_id)
        if roster is None:
            ef.parse_warnings.append("employee not in roster (HR Settings/employees.csv)")
        elif roster.get("active") is False:
            ef.parse_warnings.append("employee marked INACTIVE in roster -> skipped")
            employee_files.append(ef)
            print(f"       {ef.emp_name or '(unknown)'} / ID {ef.emp_id} : SKIPPED (inactive)")
            continue

        # rate precedence: file's own rate -> roster default -> policy default
        roster_rate = roster.get("default_rate") if roster else None
        rate = ef.rate or roster_rate or default_rate
        if not ef.rate:
            ef.parse_warnings.append(
                f"file rate blank -> used {'roster' if roster_rate else 'policy'} default {rate}")
        if roster and roster.get("employee_name") and not ef.emp_name:
            ef.emp_name = roster["employee_name"]

        for trip in ef.trips:
            rr = process_trip(ef, trip, master, aliases, policy, rate, hr_locations)
            rr._file = fname
            all_rows.append(rr)
        employee_files.append(ef)
        print(f"       {ef.emp_name or '(unknown)'} / ID {ef.emp_id} : {len(ef.trips)} legs")

    run_ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    out_name = f"Gas_Reimbursement_Review_{month}.xlsx"
    out_path = os.path.join(out_dir, out_name)
    build_report(all_rows, employee_files, master, policy, month, out_path, run_ts,
                 config_source=config_source, hr_locations=hr_locations)

    # ---- console summary ----
    n_appr = sum(1 for x in all_rows if x.status == ST_APPROVED)
    n_rev = sum(1 for x in all_rows if x.status in REVIEW_STATUSES)
    n_exc = sum(1 for x in all_rows if x.status in EXCEPTION_STATUSES)
    print("\n================ RUN SUMMARY ================")
    print(f"Month           : {month}")
    print(f"Files           : {len(employee_files)}")
    print(f"Total legs      : {len(all_rows)}")
    print(f"  Auto-approved : {n_appr}")
    print(f"  Needs review  : {n_rev}")
    print(f"  Exceptions    : {n_exc}")
    print(f"Report written  : {out_path}")
    print("=============================================")


if __name__ == "__main__":
    main()
