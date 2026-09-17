"""Typed domain objects for the fleet reimbursement workflow.

The existing workbook pipeline is intentionally procedural because it must
preserve compatibility with Excel deliverables. This module adds a small,
dependency-free domain boundary so callers can validate and reason about
trip records without opening a workbook.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from enum import Enum
from typing import Iterable, Mapping, Sequence

from validation import normalize_text, reimbursement_amount, validate_rate, validate_trip_distance


class TripStatus(str, Enum):
    """Stable workflow states used by downstream review tooling."""

    PENDING = "pending"
    APPROVED = "approved"
    HOLD = "hold"
    REJECTED = "rejected"


class FlagCode(str, Enum):
    """Reasons a trip requires human review."""

    HOME_LEG = "home_leg"
    MULTI_STOP = "multi_stop"
    AMBIGUOUS = "ambiguous"
    MILEAGE_VARIANCE = "mileage_variance"
    INVALID_INPUT = "invalid_input"


@dataclass(frozen=True)
class Location:
    """Normalized location reference used during matching."""

    code: str
    name: str
    aliases: tuple[str, ...] = ()
    latitude: float | None = None
    longitude: float | None = None

    def __post_init__(self) -> None:
        if not self.code.strip():
            raise ValueError("location code must be non-empty")
        if not self.name.strip():
            raise ValueError("location name must be non-empty")

    @property
    def normalized_aliases(self) -> tuple[str, ...]:
        """Return deterministic normalized aliases including the name."""
        values = {normalize_text(self.name)}
        values.update(normalize_text(a) for a in self.aliases)
        return tuple(sorted(v for v in values if v))


@dataclass(frozen=True)
class Employee:
    """Minimal employee identity required by the reimbursement engine."""

    employee_id: str
    name: str
    rate: float
    active: bool = True

    def __post_init__(self) -> None:
        if not self.employee_id.strip():
            raise ValueError("employee_id must be non-empty")
        if not self.name.strip():
            raise ValueError("employee name must be non-empty")
        validate_rate(self.rate)


@dataclass(frozen=True)
class Trip:
    """One normalized reimbursement claim before policy evaluation."""

    employee_id: str
    trip_date: date
    origin: str
    destination: str
    claimed_miles: float
    matrix_miles: float
    rate: float
    status: TripStatus = TripStatus.PENDING
    notes: str = ""

    def __post_init__(self) -> None:
        if not self.employee_id.strip():
            raise ValueError("employee_id must be non-empty")
        if not self.origin.strip() or not self.destination.strip():
            raise ValueError("origin and destination must be non-empty")
        validate_trip_distance(self.claimed_miles)
        validate_trip_distance(self.matrix_miles)
        validate_rate(self.rate)

    @property
    def variance_miles(self) -> float:
        """Difference between claimed and authoritative matrix distance."""
        return round(self.claimed_miles - self.matrix_miles, 2)

    @property
    def reimbursable_miles(self) -> float:
        """Return authoritative road miles, never the claimed value."""
        return self.matrix_miles

    @property
    def reimbursement(self) -> float:
        """Calculate the proposed reimbursement from matrix miles."""
        return round(reimbursement_amount(self.matrix_miles, self.rate), 2)

    @property
    def normalized_origin(self) -> str:
        return normalize_text(self.origin)

    @property
    def normalized_destination(self) -> str:
        return normalize_text(self.destination)


@dataclass(frozen=True)
class ReviewFlag:
    """A deterministic review finding attached to a trip."""

    code: FlagCode
    message: str
    severity: str = "warning"

    def __post_init__(self) -> None:
        if not self.message.strip():
            raise ValueError("flag message must be non-empty")
        if self.severity not in {"info", "warning", "error"}:
            raise ValueError("severity must be info, warning, or error")


@dataclass(frozen=True)
class ReviewDecision:
    """Policy outcome for one trip."""

    status: TripStatus
    flags: tuple[ReviewFlag, ...] = ()
    approved_miles: float = 0.0
    approved_amount: float = 0.0
    reviewer_note: str = ""

    def __post_init__(self) -> None:
        validate_trip_distance(self.approved_miles)
        if self.approved_amount < 0:
            raise ValueError("approved_amount must be non-negative")
        if self.status is TripStatus.APPROVED and self.flags:
            raise ValueError("approved trips cannot retain blocking flags")

    @property
    def requires_human_review(self) -> bool:
        return self.status in {TripStatus.PENDING, TripStatus.HOLD}


@dataclass(frozen=True)
class MonthlySummary:
    """Immutable monthly totals suitable for reporting and reconciliation."""

    employee_id: str
    trip_count: int
    approved_count: int
    held_count: int
    rejected_count: int
    approved_miles: float
    approved_amount: float

    def __post_init__(self) -> None:
        for name, value in (
            ("trip_count", self.trip_count),
            ("approved_count", self.approved_count),
            ("held_count", self.held_count),
            ("rejected_count", self.rejected_count),
        ):
            if value < 0:
                raise ValueError(f"{name} must be non-negative")
        if self.approved_count + self.held_count + self.rejected_count > self.trip_count:
            raise ValueError("decision counts cannot exceed trip count")
        validate_trip_distance(self.approved_miles)
        if self.approved_amount < 0:
            raise ValueError("approved_amount must be non-negative")


def build_employee(record: Mapping[str, object]) -> Employee:
    """Build an Employee from CSV-like data with explicit defaults."""
    active_value = str(record.get("active", "YES")).strip().upper()
    return Employee(
        employee_id=str(record.get("employee_id", "")).strip(),
        name=str(record.get("employee_name", "")).strip(),
        rate=float(record.get("default_rate", 0) or 0),
        active=active_value not in {"NO", "FALSE", "0"},
    )


def build_trip(record: Mapping[str, object], rate: float) -> Trip:
    """Build a Trip from a row mapping used by workbook importers."""
    raw_date = record.get("trip_date")
    if isinstance(raw_date, date):
        trip_date = raw_date
    else:
        trip_date = date.fromisoformat(str(raw_date))
    return Trip(
        employee_id=str(record.get("employee_id", "")).strip(),
        trip_date=trip_date,
        origin=str(record.get("origin", "")).strip(),
        destination=str(record.get("destination", "")).strip(),
        claimed_miles=float(record.get("claimed_miles", 0) or 0),
        matrix_miles=float(record.get("matrix_miles", 0) or 0),
        rate=rate,
    )


def summarize(employee_id: str, decisions: Iterable[ReviewDecision]) -> MonthlySummary:
    """Aggregate review decisions without mutating the source records."""
    rows = list(decisions)
    approved = [r for r in rows if r.status is TripStatus.APPROVED]
    held = [r for r in rows if r.status is TripStatus.HOLD]
    rejected = [r for r in rows if r.status is TripStatus.REJECTED]
    return MonthlySummary(
        employee_id=employee_id,
        trip_count=len(rows),
        approved_count=len(approved),
        held_count=len(held),
        rejected_count=len(rejected),
        approved_miles=round(sum(r.approved_miles for r in approved), 2),
        approved_amount=round(sum(r.approved_amount for r in approved), 2),
    )


def group_by_employee(trips: Sequence[Trip]) -> dict[str, tuple[Trip, ...]]:
    """Group immutable trip records by employee in deterministic order."""
    groups: dict[str, list[Trip]] = {}
    for trip in trips:
        groups.setdefault(trip.employee_id, []).append(trip)
    return {
        key: tuple(sorted(value, key=lambda item: (item.trip_date, item.origin, item.destination)))
        for key, value in sorted(groups.items())
    }
