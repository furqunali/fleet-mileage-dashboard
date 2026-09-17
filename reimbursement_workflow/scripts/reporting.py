"""Stable reporting projections for reimbursement review runs."""
from __future__ import annotations

from collections import Counter
from dataclasses import asdict
from datetime import date
from typing import Iterable, Sequence

from domain import ReviewDecision, Trip, TripStatus
from policy_engine import decision_reason, explain_flags


def decision_row(trip: Trip, decision: ReviewDecision) -> dict[str, object]:
    """Project a trip decision into a flat reporting row."""
    return {
        "employee_id": trip.employee_id,
        "trip_date": trip.trip_date.isoformat(),
        "origin": trip.origin,
        "destination": trip.destination,
        "claimed_miles": round(trip.claimed_miles, 2),
        "matrix_miles": round(trip.matrix_miles, 2),
        "variance_miles": trip.variance_miles,
        "status": decision.status.value,
        "approved_miles": round(decision.approved_miles, 2),
        "approved_amount": round(decision.approved_amount, 2),
        "flags": explain_flags(decision.flags),
        "reason": decision_reason(decision),
        "reviewer_note": decision.reviewer_note,
    }


def build_rows(
    trips: Sequence[Trip],
    decisions: Sequence[ReviewDecision],
) -> list[dict[str, object]]:
    """Build report rows while enforcing one decision per trip."""
    if len(trips) != len(decisions):
        raise ValueError("trip and decision counts must match")
    return [decision_row(trip, decision) for trip, decision in zip(trips, decisions)]


def status_summary(decisions: Iterable[ReviewDecision]) -> dict[str, int]:
    """Return counts for every stable workflow status."""
    counts = Counter(decision.status.value for decision in decisions)
    return {status.value: counts.get(status.value, 0) for status in TripStatus}


def flag_summary(decisions: Iterable[ReviewDecision]) -> dict[str, int]:
    """Count review flags across a run."""
    counts: Counter[str] = Counter()
    for decision in decisions:
        for flag in decision.flags:
            counts[flag.code.value] += 1
    return dict(sorted(counts.items()))


def financial_summary(decisions: Iterable[ReviewDecision]) -> dict[str, float]:
    """Return approved mileage and reimbursement totals."""
    approved = [d for d in decisions if d.status is TripStatus.APPROVED]
    return {
        "approved_miles": round(sum(d.approved_miles for d in approved), 2),
        "approved_amount": round(sum(d.approved_amount for d in approved), 2),
    }


def build_month_report(
    *,
    month: str,
    trips: Sequence[Trip],
    decisions: Sequence[ReviewDecision],
) -> dict[str, object]:
    """Build a complete deterministic report payload."""
    if len(trips) != len(decisions):
        raise ValueError("trip and decision counts must match")
    rows = build_rows(trips, decisions)
    return {
        "month": month,
        "generated_for": "hr-review",
        "trip_count": len(rows),
        "status_counts": status_summary(decisions),
        "flag_counts": flag_summary(decisions),
        "financials": financial_summary(decisions),
        "rows": rows,
    }


def employee_report(
    employee_id: str,
    trips: Sequence[Trip],
    decisions: Sequence[ReviewDecision],
) -> dict[str, object]:
    """Return a report limited to one employee."""
    selected = [
        (trip, decision)
        for trip, decision in zip(trips, decisions)
        if trip.employee_id == employee_id
    ]
    return {
        "employee_id": employee_id,
        "trip_count": len(selected),
        "rows": [decision_row(trip, decision) for trip, decision in selected],
    }


def review_queue(
    trips: Sequence[Trip],
    decisions: Sequence[ReviewDecision],
) -> list[dict[str, object]]:
    """Return only rows that need a human reviewer."""
    return [
        decision_row(trip, decision)
        for trip, decision in zip(trips, decisions)
        if decision.requires_human_review
    ]


def approved_rows(
    trips: Sequence[Trip],
    decisions: Sequence[ReviewDecision],
) -> list[dict[str, object]]:
    """Return rows that are safe for approved-total reporting."""
    return [
        decision_row(trip, decision)
        for trip, decision in zip(trips, decisions)
        if decision.status is TripStatus.APPROVED
    ]


def validate_report(report: dict[str, object]) -> None:
    """Validate the minimum schema required by downstream exporters."""
    required = {"month", "trip_count", "status_counts", "flag_counts", "financials", "rows"}
    missing = required.difference(report)
    if missing:
        raise ValueError(f"report missing fields: {sorted(missing)}")
    if not isinstance(report["month"], str) or not report["month"].strip():
        raise ValueError("report month must be non-empty")
    if not isinstance(report["rows"], list):
        raise ValueError("report rows must be a list")


def report_date(month: str) -> date:
    """Parse YYYY-MM into the first day of the reporting month."""
    if len(month) != 7 or month[4] != "-":
        raise ValueError("month must use YYYY-MM format")
    return date.fromisoformat(f"{month}-01")


def redact_report_row(row: dict[str, object]) -> dict[str, object]:
    """Return a shareable row without reviewer notes."""
    sanitized = dict(row)
    sanitized.pop("reviewer_note", None)
    return sanitized


def shareable_report(report: dict[str, object]) -> dict[str, object]:
    """Create a sanitized report suitable for non-HR diagnostics."""
    validate_report(report)
    result = dict(report)
    result["rows"] = [redact_report_row(row) for row in report["rows"]]
    return result
