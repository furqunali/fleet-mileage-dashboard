# Fleet Mileage & Gas-Reimbursement Toolset

A Python + Excel + HTML toolset for a small fleet operation. It computes
**driving distances between company sites and employee homes**, produces a
branded **distance / mileage dashboard**, and runs a **monthly gas-reimbursement
workflow** with an HR approval step.

> **Sanitized public demo.** All employee names, home/street addresses, site
> names, mileage figures and dollar amounts in this repository are **fictional**.
> Real data files (`.xlsx` / `.xlsm` / `.csv` / `.pdf`, `approvals.json`,
> `ledger.csv`, `geocode_cache.json`, the `HR_Settings.xlsx` control panel, and
> the monthly inbox/output/backups) are **not** included and are git-ignored.

---

## What it does

**1. Distance / mileage dashboard (`excel_pipeline/`)**
- A one-time geocoding step (`geocode_sites.py`) resolves each site address to
  coordinates via Nominatim/OpenStreetMap, validated against the trusted
  distance matrix (Texas bounding box + haversine cross-check).
- A series of *surgical XML zip edits* build up the workbook version by version
  (`patch_workbook.py` &rarr; `polish_workbook.py` &rarr; `visual_polish_workbook.py`
  &rarr; `add_homes_v6.py` &rarr; `add_monthly_plan_sheet.py` /
  `restyle_monthly_plan_v6.py` &rarr; `make_v7.py` / `protect_workbook_v7.py` &rarr;
  `add_extra_entries_v8.py` &rarr; `build_v9_buttons.py`). Editing the raw XML
  (instead of re-saving through a spreadsheet library) preserves the slicer,
  pivot cache, Power-Pivot data model, tables, the embedded map image and cached
  formula values.
- `build_html_from_workbook.py` treats the workbook as the single source of
  truth and regenerates the standalone HTML calculator's embedded `DATA` block,
  with a `--verify` mode that fails if the HTML has drifted from the workbook.

**2. Gas-reimbursement workflow (`reimbursement_workflow/`)**
- Reads each employee's monthly mileage workbook, matches messy free-text
  locations against the authoritative site directory + distance matrix, and
  cross-checks claimed odometer miles against the matrix road miles.
- Locked policy rules: matrix road miles are authoritative; **home/apartment
  legs** and **multi-store lines** are always flagged for HR ($0 until approved);
  a city mapping to more than one site is **ambiguous** and never auto-picked.
- Three front-ends share one matching engine (`process_month.py`): a per-file
  pipeline, a single **HR Console** workbook (`build_console.py` /
  `process_console.py`), and a fully self-contained **native-formula Calculator**
  workbook (`build_calculator.py`).
- An HR web **control panel** (`web/hr_control_panel.html`) lets a non-technical
  reviewer load the run output, approve/hold each employee, and draft the
  Director email &mdash; the email is always a human-reviewed draft, never
  auto-sent.

---

## Tech stack

- **Python 3** with **openpyxl** (workbook read/build) and the standard library
  (`zipfile`, `re`, `csv`, `json`, `difflib`, `urllib`). `build_v9_buttons.py`
  optionally uses `pywin32` (Excel COM) on Windows to add macro buttons.
- **Excel** (`.xlsx` / `.xlsm`) as the data + delivery format.
- **HTML/CSS/JS** for the standalone distance calculator and the HR control panel.

---

## Repository layout

```
fleet-mileage-dashboard/
├── excel_pipeline/            # scripts that build the distance/mileage workbook + HTML
├── reimbursement_workflow/
│   ├── scripts/               # matching engine + console/calculator builders
│   ├── config/policy.json     # policy knobs (rate, thresholds, flag rules)
│   ├── config_samples/        # *.sample.csv fictional roster / alias examples
│   ├── Run.bat, Refresh-Matrix.bat
│   └── README.md              # detailed workflow docs
├── web/
│   ├── hr_control_panel.html            # HR approval panel (loads run output at runtime)
│   └── Fleet_Distance_Calculator_template.html  # generator template (fictional DATA)
└── demo/
    └── index.html             # self-contained, deployable demo (fictional data, no calls)
```

---

## Getting started

```bash
python -m venv .venv && . .venv/bin/activate      # or .venv\Scripts\activate on Windows
pip install openpyxl                               # pywin32 only for build_v9_buttons.py
cp .env.example .env                               # then edit values
```

The pipeline scripts expect the real, private data workbooks/config to be
present locally (they are not shipped). Provide your own:
`reimbursement_workflow/config/employees.csv` and `location_aliases.csv` (see the
`config_samples/` schema) or an `HR_Settings.xlsx`, plus the master distance
workbook referenced in `policy.json`.

To preview the UI with fictional data, just open **`demo/index.html`** in a
browser, or deploy the `demo/` folder as a static site.

See **`reimbursement_workflow/README.md`** for the full monthly workflow.
