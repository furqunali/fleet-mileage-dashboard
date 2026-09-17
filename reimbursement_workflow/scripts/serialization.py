"""JSON-safe serialization boundaries for reimbursement domain objects."""
from __future__ import annotations

import json
from datetime import date
from enum import Enum
from typing import Any, Mapping, Sequence

from domain import ReviewDecision, ReviewFlag, Trip


class DomainJSONEncoder(json.JSONEncoder):
    """Encode workflow dates, enums and domain records deterministically."""

    def default(self, value: Any) -> Any:
        if isinstance(value, date):
            return value.isoformat()
        if isinstance(value, Enum):
            return value.value
        if isinstance(value, ReviewFlag):
            return {
                "code": value.code.value,
                "message": value.message,
                "severity": value.severity,
            }
        if isinstance(value, ReviewDecision):
            return {
                "status": value.status.value,
                "flags": list(value.flags),
                "approved_miles": value.approved_miles,
                "approved_amount": value.approved_amount,
                "reviewer_note": value.reviewer_note,
            }
        if isinstance(value, Trip):
            return {
                "employee_id": value.employee_id,
                "trip_date": value.trip_date,
                "origin": value.origin,
                "destination": value.destination,
                "claimed_miles": value.claimed_miles,
                "matrix_miles": value.matrix_miles,
                "rate": value.rate,
                "status": value.status.value,
                "notes": value.notes,
            }
        return super().default(value)


def to_json(value: Any, *, indent: int | None = None) -> str:
    """Serialize a domain value with stable key ordering."""
    return json.dumps(
        value,
        cls=DomainJSONEncoder,
        indent=indent,
        sort_keys=True,
        ensure_ascii=False,
    )


def from_json(payload: str) -> Any:
    """Parse JSON while rejecting empty/non-object payloads."""
    if not payload.strip():
        raise ValueError("payload must be non-empty")
    return json.loads(payload)


def trip_to_dict(trip: Trip) -> dict[str, object]:
    """Return a stable primitive mapping for a Trip."""
    return {
        "employee_id": trip.employee_id,
        "trip_date": trip.trip_date.isoformat(),
        "origin": trip.origin,
        "destination": trip.destination,
        "claimed_miles": trip.claimed_miles,
        "matrix_miles": trip.matrix_miles,
        "rate": trip.rate,
        "status": trip.status.value,
        "notes": trip.notes,
    }


def decision_to_dict(decision: ReviewDecision) -> dict[str, object]:
    """Return a stable primitive mapping for a ReviewDecision."""
    return {
        "status": decision.status.value,
        "flags": [
            {
                "code": flag.code.value,
                "message": flag.message,
                "severity": flag.severity,
            }
            for flag in decision.flags
        ],
        "approved_miles": decision.approved_miles,
        "approved_amount": decision.approved_amount,
        "reviewer_note": decision.reviewer_note,
    }


def trips_to_json(trips: Sequence[Trip]) -> str:
    """Serialize a sequence of trips for deterministic diagnostics."""
    return to_json([trip_to_dict(trip) for trip in trips])


def decisions_to_json(decisions: Sequence[ReviewDecision]) -> str:
    """Serialize a sequence of decisions for deterministic diagnostics."""
    return to_json([decision_to_dict(decision) for decision in decisions])


def json_object(payload: str) -> dict[str, Any]:
    """Parse a JSON object and reject arrays/scalars at config boundaries."""
    value = from_json(payload)
    if not isinstance(value, dict):
        raise ValueError("expected a JSON object")
    return value


def json_array(payload: str) -> list[Any]:
    """Parse a JSON array and reject objects/scalars at batch boundaries."""
    value = from_json(payload)
    if not isinstance(value, list):
        raise ValueError("expected a JSON array")
    return value


def round_trip_json(value: Mapping[str, Any]) -> Mapping[str, Any]:
    """Verify that a mapping survives a JSON serialization cycle."""
    return json_object(to_json(value))
