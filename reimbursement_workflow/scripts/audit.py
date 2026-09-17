"""Audit trail helpers for reimbursement workflow decisions."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Iterable, Mapping

from domain import ReviewDecision, Trip
from reporting import decision_row


@dataclass(frozen=True)
class AuditEvent:
    event_type: str
    employee_id: str
    timestamp: str
    details: Mapping[str, object]

    def __post_init__(self) -> None:
        if not self.event_type.strip():
            raise ValueError("event_type must be non-empty")
        if not self.employee_id.strip():
            raise ValueError("employee_id must be non-empty")
        if not self.timestamp.strip():
            raise ValueError("timestamp must be non-empty")


def utc_timestamp() -> str:
    """Return an explicit UTC timestamp suitable for audit records."""
    return datetime.now(timezone.utc).isoformat()


def trip_event(
    trip: Trip,
    decision: ReviewDecision,
    *,
    event_type: str = "trip_reviewed",
    timestamp: str | None = None,
) -> AuditEvent:
    """Create an audit event from a reviewed trip."""
    row = decision_row(trip, decision)
    return AuditEvent(
        event_type=event_type,
        employee_id=trip.employee_id,
        timestamp=timestamp or utc_timestamp(),
        details=row,
    )


def batch_events(
    trips: Iterable[Trip],
    decisions: Iterable[ReviewDecision],
    *,
    timestamp: str | None = None,
) -> tuple[AuditEvent, ...]:
    """Create paired audit events while rejecting count mismatches."""
    trip_rows = list(trips)
    decision_rows = list(decisions)
    if len(trip_rows) != len(decision_rows):
        raise ValueError("trip and decision counts must match")
    stamp = timestamp or utc_timestamp()
    return tuple(
        trip_event(trip, decision, timestamp=stamp)
        for trip, decision in zip(trip_rows, decision_rows)
    )


def event_to_dict(event: AuditEvent) -> dict[str, object]:
    """Convert an event into JSON-safe primitives."""
    return {
        "event_type": event.event_type,
        "employee_id": event.employee_id,
        "timestamp": event.timestamp,
        "details": dict(event.details),
    }


def validate_event(event: AuditEvent) -> None:
    """Validate basic audit invariants before persistence."""
    if "status" not in event.details:
        raise ValueError("audit event must include decision status")
    if "employee_id" not in event.details:
        raise ValueError("audit event must include employee id")
    if event.details["employee_id"] != event.employee_id:
        raise ValueError("audit employee id must match event owner")


def validate_events(events: Iterable[AuditEvent]) -> None:
    """Validate an entire audit sequence."""
    for event in events:
        validate_event(event)


def append_event(
    events: Iterable[AuditEvent],
    event: AuditEvent,
) -> tuple[AuditEvent, ...]:
    """Append an event without mutating the caller's collection."""
    validate_event(event)
    return tuple(events) + (event,)


def events_for_employee(
    events: Iterable[AuditEvent],
    employee_id: str,
) -> tuple[AuditEvent, ...]:
    """Filter audit events by employee without changing their order."""
    if not employee_id.strip():
        raise ValueError("employee_id must be non-empty")
    return tuple(event for event in events if event.employee_id == employee_id)


def events_by_type(
    events: Iterable[AuditEvent],
) -> dict[str, tuple[AuditEvent, ...]]:
    """Group events by type using deterministic key ordering."""
    grouped: dict[str, list[AuditEvent]] = {}
    for event in events:
        grouped.setdefault(event.event_type, []).append(event)
    return {
        key: tuple(grouped[key])
        for key in sorted(grouped)
    }


def summarize_audit(events: Iterable[AuditEvent]) -> dict[str, object]:
    """Produce compact audit metrics for an operational report."""
    rows = tuple(events)
    validate_events(rows)
    by_type = events_by_type(rows)
    return {
        "event_count": len(rows),
        "employee_count": len({event.employee_id for event in rows}),
        "event_types": {key: len(value) for key, value in by_type.items()},
        "first_timestamp": min((event.timestamp for event in rows), default=None),
        "last_timestamp": max((event.timestamp for event in rows), default=None),
    }


def ensure_monotonic_timestamps(events: Iterable[AuditEvent]) -> None:
    """Reject audit sequences that move backwards in timestamp order."""
    rows = tuple(events)
    timestamps = [event.timestamp for event in rows]
    if timestamps != sorted(timestamps):
        raise ValueError("audit events must be ordered by timestamp")


def redact_audit_event(event: AuditEvent) -> AuditEvent:
    """Remove reviewer note from an event before external sharing."""
    details = dict(event.details)
    details.pop("reviewer_note", None)
    return AuditEvent(
        event_type=event.event_type,
        employee_id=event.employee_id,
        timestamp=event.timestamp,
        details=details,
    )
