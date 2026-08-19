# Fleet Monthly Gas-Reimbursement Workflow

A reusable system that matches trip locations against the authoritative Fleet site
Directory + distance Matrix, cross-checks claimed odometer mileage, and routes
anything uncertain to human review instead of guessing.

---

## ★★ THE CALCULATOR — simplest, self-contained, no Python at runtime

**`Gas_Reimbursement_Calculator.xlsx`** (in `Reimbursement_Workflow/`) works like
the Fleet distance-calculator workbook: **native Excel formulas** do everything.
No macros, no Run.bat, no config files. HR just uses Excel.

**Tabs:** Dashboard · Monthly Input · Employees · Locations · Monthly Report ·
Settings · MatrixData (embedded authoritative matrix, locked).

**What HR does each month:**
1. **Monthly Input** — enter/paste the month's trips. Pick **Origin** and
   **Destination** from the dropdowns (site names, or `HOME`, `MULTI-STOP`,
   `OTHER (review)`), enter odometer start/end. Every calculated column
   (miles, status, reimbursement) fills in automatically.
2. **Employees** — add people + per-mile rates.
3. **Locations** — add any new verified address/site (Text → Type → Matrix
   Code). New entries are used automatically.
4. **Settings** — add the new month to the month list (e.g. `2026-06`).
5. **Dashboard** — pick the **Selected Month**; press **F9** to recalculate.
   KPIs + Monthly Report update instantly.

**How it calculates (all native Excel):**
- Origin/Destination text → site code via `VLOOKUP` on the **Locations** table.
- Miles → 2-D `INDEX/MATCH` on the embedded **MatrixData** (authoritative).
- Status → `IF` logic: `HOME_ADDRESS_REVIEW` / `MULTI-STOP` / `REVIEW - OTHER`
  (ambiguous or unknown) / `EXCEPTION - …` (unresolved, no matrix distance,
  variance) / `APPROVED`.
- Reimbursement → **Matrix miles × rate, only when APPROVED**; everything else
  is **$0** pending review. Dashboard/Report totals via `SUMIFS`/`COUNTIFS`.
- **Matrix is authoritative; anything not on the Matrix is never auto-paid.**
- All months are kept on Monthly Input; the Dashboard month selector scopes the
  KPIs and Monthly Report. History is retained in the one file.

**To (re)build or refresh** (e.g. after the master Matrix changes) — the only
time Python is used: `python scripts/build_calculator.py --force` (re-embeds the
current master Matrix and rewires the workbook).

**Verified:** opened in Excel and recalculated — embedded Matrix lookup returns
correct road miles (e.g. 0083→0075 = 181), May sample = 31 legs, 0 auto-approved,
$0 auto, $1,320.57 claimed; home legs $0; ambiguous legs flagged for review.

*(The Python-engine versions below — the HR Console and the per-file pipeline —
remain available but are superseded by this calculator for day-to-day HR use.)*

---

## ★ HR CONSOLE (the primary way to use this)

HR works from **one** Excel workbook: **`Gas_Reimbursement_Console.xlsx`**
(in the `Reimbursement_Workflow/` folder). No other files to open, no config
files to maintain.

**Tabs:**

| Tab | Who fills it | What it is |
|---|---|---|
| **Dashboard** | pick month | Month selector + this-month KPIs + all-month history. |
| **Monthly Input** | HR pastes | Every month's trips (all history kept in one table). |
| **Employees** | HR | Roster + per-mile rates + Active flag. |
| **Locations** | HR | Verified HOME / WORK / OTHER addresses + coordinates. |
| **Settings** | locked | Policy knobs (rate, thresholds). Protected against accidental edits. |
| **Monthly Report** | engine | Per-employee totals for the selected month. |
| **Trip Detail** | engine | Every leg: route, codes, Matrix miles, reference miles, status, reimbursement. |
| **Exceptions** | engine | Everything needing review for the selected month. |
| **Audit Trail** | engine | Run metadata, config in effect, locations loaded. |

**HR's routine each month:**
1. On **Monthly Input**, paste the month's trips (columns: Month `YYYY-MM`,
   Employee ID, Date, Origin, Destination, Notes, Odometer Start/End, Claimed Miles).
2. Keep **Employees** and **Locations** current (add people, rates, addresses).
3. On the **Dashboard**, pick the **Selected Month**.
4. **Save and close** the workbook.
5. Double-click **`Run.bat`**.
6. Re-open — Monthly Report, Trip Detail, Exceptions and the Dashboard are updated.

**What the engine does (backend, `scripts/process_console.py`):**
- Reads the master **Directory/Matrix read-only** (a separate, protected file —
  never inside the console) → **Matrix miles are authoritative**.
- Applies the locked rules: HOME → `HOME_ADDRESS_REVIEW` **$0** until HR approval;
  multi-store → **$0**, manual itemisation, route never guessed; a city that maps
  to >1 site → **AMBIGUOUS**; a verified location **not on the Matrix** →
  `REVIEW - OFF-MATRIX LOCATION`, **never auto-paid**.
- Only a clean, single-stop, non-home, in-variance leg with both ends on the
  Matrix is ever auto-approved.
- **Keeps all history**: Monthly Input holds every month; the Dashboard shows a
  per-month history table; each run **backs up** the workbook (`backups/`) and
  **archives** a timestamped copy (`output/<month>/`). The master is never modified.

**Add a new employee / location / month?** Just edit the console and run again —
no code change. Locations you add are used automatically next run.

To (re)create a blank console template: `python scripts/build_console.py`
(refuses to overwrite an existing console unless `--force`).

---

## Legacy per-file pipeline (backend detail)

The original flow below still works and shares the same matching engine
(`scripts/process_month.py`, which the console imports). It processes one raw
file per employee from `inbox/<month>/` into a separate report. The HR Console
supersedes it for day-to-day use.

---

## Folder layout

```
Reimbursement_Workflow/
├── README.md                     ← this file
├── config/
│   ├── HR_Settings.xlsx          ← ★ HR CONTROL PANEL (employees + verified locations)
│   ├── policy.json               ← global rules (rate, thresholds, flag policies)
│   ├── employees.csv             ← fallback roster (used only if HR_Settings.xlsx is absent)
│   └── location_aliases.csv      ← VERIFIED free-text → site-code mappings
├── scripts/
│   ├── process_month.py          ← the engine (run this each month)
│   └── build_hr_settings.py      ← (re)creates a blank HR_Settings.xlsx template
├── inbox/
│   └── <YYYY-MM>/                ← drop each employee's raw file here
│       └── *.xlsx
└── output/
    └── <YYYY-MM>/
        └── Gas_Reimbursement_Review_<YYYY-MM>.xlsx   ← generated report
```

The master distance workbook (`Fleet_Distance_Dashboard_v5.xlsx`, in the
project root) is the read-only source of truth for site codes, addresses and
road miles. **Originals are never modified** — the engine opens every source
file read-only and writes only into `output/`.

---

## How to run

```bash
cd Reimbursement_Workflow/scripts
python process_month.py --month 2026-05      # process a specific month
python process_month.py                       # process the newest month folder
```

Requires Python 3 + `openpyxl` (`pip install openpyxl`). No other dependencies.

Console output prints a run summary; the full report lands in
`output/<month>/Gas_Reimbursement_Review_<month>.xlsx`.

---

## Processing a new month (5–7 employees)

1. Create `inbox/<YYYY-MM>/` (e.g. `inbox/2026-06/`).
2. Drop in each employee's raw gas-reimbursement `.xlsx`. File names don't
   matter; the engine reads the employee name/ID from inside each file.
3. Make sure each employee is listed in **`config/HR_Settings.xlsx`** (see below).
4. Run `python process_month.py --month 2026-06`.
5. Open the generated report and work the **Exceptions** tab.

All employees for the month are combined into **one** workbook with per-employee
summary rows.

---

## HR Settings — the control panel (`config/HR_Settings.xlsx`)

This Excel workbook is where HR maintains everything from their own computer.
The engine **reads it on every run**, so anything HR adds or edits is used
automatically for the next month — no code change, no programmer.

If `HR_Settings.xlsx` is present it is the authoritative config. If it is ever
deleted, the engine falls back to the legacy `employees.csv` + `location_aliases.csv`.
To (re)create a blank template: `python scripts/build_hr_settings.py`
(it refuses to overwrite an existing file unless you pass `--force`).

**Tab: `Employees`** — one row per person. Columns: Employee ID, Employee Name,
Default Rate, Active (YES/NO), Notes. Add new staff by adding rows.

**Tab: `Locations`** — the fully-expandable verified-location table. Add as many
rows as you like, for any employee, now or in the future:

| Column | Purpose |
|---|---|
| Location ID | Any unique label (e.g. LOC001). Blank / `(example)` rows are ignored. |
| Employee ID | Whose location it is. Use `ALL` for a shared location. |
| Type | `HOME`, `WORK`, or `OTHER`. |
| Aliases | Every spelling seen in the raw log, separated by `;`. |
| Verified Address | The confirmed street address. |
| Latitude / Longitude | From **Google Maps** — reference only. |
| Matrix Site Code | *(optional)* link a WORK location to an existing site code (e.g. `0083`) so official Matrix miles are used. |
| HR Approved | YES/NO. |
| Notes | Free text. |

**How locations are used (Matrix stays authoritative):**
- **`HOME`** → always `HOME_ADDRESS_REVIEW` with **$0** until HR approval, *even
  after the address is filled in*. If coordinates are present, the report also
  shows a straight-line **reference** distance to help HR review — never a payment.
- **`WORK`/`OTHER` linked to a Matrix Site Code** → treated as that site, so
  **official Matrix road miles** are used and the leg can auto-calculate normally.
- **`WORK`/`OTHER` with coordinates but no Matrix code** (e.g. a brand-new site
  not yet in the Matrix) → `REVIEW - OFF-MATRIX LOCATION`, **$0**, with a
  straight-line reference distance. It is never auto-paid, because it is not on
  the authoritative Matrix.

The master site **Directory** and distance **Matrix** are never modified by HR
edits here; new verified locations live entirely in `HR_Settings.xlsx`.

---

## Other config files

### `policy.json` — global rules
| Key | Meaning |
|---|---|
| `master_workbook` | Authoritative distance workbook (project root). |
| `default_rate_per_mile` | Used only if a file's own rate cell is blank. |
| `pay_basis` | `matrix` — reimburse on Matrix road miles (authoritative). |
| `variance_threshold_pct` | Claimed-vs-matrix gap above this % → VARIANCE exception. |
| `fuzzy_match_cutoff` / `fuzzy_match_margin` | Name-matching sensitivity. |
| `home_leg_policy` / `multistop_policy` | Both `flag` (see rules below). |

### `employees.csv` — legacy fallback roster
Only used if `HR_Settings.xlsx` is missing. Columns: `employee_id,
employee_name, home_label, home_address, home_lat, home_lon, default_rate, notes`.

### `location_aliases.csv` — verified mappings only
`raw_pattern, site_code, site_name, verified_reason`. **Only add an alias when
the mapping is certain and unique.** Seeded with:
- `123 demo blvd` → `0075` (Demo Store North)
- `west store` → `0083` (Demo Store West)

Deliberately **not** aliased: `region-a store` (0025 vs 0078) and `region-b store`
(0077 vs 0076) — each city has two sites, so they are reported AMBIGUOUS rather
than guessed. Grow this file over time using the **Unmatched Locations** tab.

---

## How a leg is decided (matching pipeline)

For each origin/destination string, in order:
0. **HR verified location** (Locations tab, scoped to the employee, then `ALL`):
   `HOME` → home flag; `WORK`/`OTHER` with a Matrix code → that site; else →
   off-Matrix review.
1. **Personal keyword** (home, demo apt, apart…) → always **flag** `HOME_ADDRESS_REVIEW`, $0.
2. **Multi-stop phrasing** (" and ", "all stores", plural "stores") → always **flag** `MULTI-STOP - MANUAL ITEMIZATION`, $0.
3. **Verified alias** (exact) → matched.
4. **Exact site code / name / dropdown label** → matched.
5. **City keyword**: one city with a single site → matched; a city with several
   sites → **AMBIGUOUS**; two or more cities → **MULTI-STOP**.
6. **Fuzzy name match** (difflib) above cutoff *and* clear of the runner-up by
   the margin → matched; otherwise **AMBIGUOUS**.
7. Nothing → **UNMATCHED**.

A leg is **auto-approved only** when *both* endpoints cleanly match, it is not
personal/multi-stop, and the odometer variance is within threshold. Reimbursable
amount is always **Matrix road miles × rate** (odometer is a cross-check only).

### Locked policy decisions (do not relax without client approval)
1. Matrix road miles are the authoritative reimbursable distance; odometer is a cross-check.
2. **Home / apartment legs** → status `HOME_ADDRESS_REVIEW`, matrix miles **not**
   applied, reimbursement forced to **$0**, flagged for HR **until HR approval**.
   Never auto-approved or auto-rejected. (The employee's home address/coordinates
   are stored in `HR_Settings.xlsx` for reference only, from Google Maps; they
   only produce a straight-line reference distance, never a payment.)
3. **Multi-store rows** → status `MULTI-STOP - MANUAL ITEMIZATION`, **no**
   automatic distance, reimbursement **$0**. Requires manual itemization; route
   order is **never** guessed.
4. A city that maps to more than one site is **AMBIGUOUS** — never auto-picked.

Only cleanly-matched, single-stop, non-home, in-variance legs ever carry a
non-zero auto-computed reimbursement. Everything else is $0 pending human review.

---

## Report tabs

| Tab | Contents |
|---|---|
| **Summary** | Per-employee totals: legs, approved / review / exception counts, auto-approved $, claimed $, pending $, plus a grand total. |
| **Trip Detail** | Every leg with raw + resolved locations, matrix miles, **straight-line reference miles**, claimed miles, variance %, reimbursable $, status, **verified addresses**, and match reasoning. |
| **Exceptions** | Every leg not auto-approved — the reviewer's worklist. |
| **Unmatched Locations** | Unique ambiguous/unresolved location strings with candidate sites and frequency — the to-do list for growing the Locations tab / `location_aliases.csv`. |
| **Audit Trail** | Run timestamp, config source, HR locations loaded (and how many approved), config values in effect, sites/matrix/coords loaded, and per-file parse notes. |

Status colours: green = approved, amber = needs review, red = exception.

---

## Extending

- **New/updated home or work location** → edit `config/HR_Settings.xlsx` (Locations
  tab); used automatically on the next run.
- **New employee** → add a row to the Employees tab of `HR_Settings.xlsx`, drop
  their file in the month folder.
- **Deactivate an employee** → set Active = NO in the Employees tab.
- **Verified global alias** → add a row to `location_aliases.csv` (or a WORK/OTHER
  row in HR_Settings with a Matrix Site Code).
- **Policy change** (rate, variance threshold) → edit `policy.json`; no code change.
- **Master data change** (new site, corrected miles) → handled automatically; the
  engine reads the current master workbook every run.
```
