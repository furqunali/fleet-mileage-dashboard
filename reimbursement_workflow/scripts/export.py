"""Safe export boundaries for reimbursement reports."""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterable, Mapping, Sequence

from reporting import validate_report


def _safe_cell(value: object) -> str:
    """Prevent accidental spreadsheet formula injection in text fields."""
    text = "" if value is None else str(value)
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def write_csv(path: str | Path, rows: Sequence[Mapping[str, object]], fields: Sequence[str]) -> Path:
    """Write a deterministic UTF-8 CSV using an explicit column contract."""
    if not fields:
        raise ValueError("fields must be non-empty")
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(fields), extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({field: _safe_cell(row.get(field)) for field in fields})
    return target


def write_report_csv(path: str | Path, report: Mapping[str, object]) -> Path:
    """Export report rows using the stable reporting schema."""
    validate_report(dict(report))
    rows = report["rows"]
    if not isinstance(rows, list):
        raise ValueError("report rows must be a list")
    fields = (
        "employee_id", "trip_date", "origin", "destination",
        "claimed_miles", "matrix_miles", "variance_miles",
        "status", "approved_miles", "approved_amount", "reason",
    )
    return write_csv(path, rows, fields)


def read_csv(path: str | Path) -> list[dict[str, str]]:
    """Read a UTF-8 CSV without applying business logic."""
    target = Path(path)
    with target.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def validate_export_path(path: str | Path, *, allowed_suffixes: Iterable[str] = (".csv",)) -> Path:
    """Validate an output path before a caller writes to disk."""
    target = Path(path)
    suffixes = {suffix.lower() for suffix in allowed_suffixes}
    if target.suffix.lower() not in suffixes:
        raise ValueError(f"unsupported export suffix: {target.suffix}")
    if target.name in {".", ".."} or not target.name:
        raise ValueError("export path must include a filename")
    return target


def export_rows(
    path: str | Path,
    rows: Sequence[Mapping[str, object]],
    fields: Sequence[str],
) -> Path:
    """Validate and write generic report rows."""
    target = validate_export_path(path)
    return write_csv(target, rows, fields)
