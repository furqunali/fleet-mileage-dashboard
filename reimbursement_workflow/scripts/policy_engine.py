"""Deterministic reimbursement policy evaluation.

This module centralizes the locked business rules documented by the project:
home/apartment legs, multi-stop claims, and ambiguous locations are never
auto-approved. Matrix road miles remain the authoritative reimbursable value.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Sequence

from domain import FlagCode, ReviewDecision, ReviewFlag, Trip, TripStatus


@dataclass(frozen=True)
class Policy:
    """Configurable thresholds that do not change the locked safety rules."""

    mileage_variance_limit: float = 10.0
    require_positive_distance: bool = False
    hold_on_zero_matrix_miles: bool = False

    def __post_init__(self) -> None:
        if self.mileage_variance_limit < 0:
            raise ValueError("mileage_variance_limit must be non-negative")


def _flag(code: FlagCode, message: str, severity: str = "warning") -> ReviewFlag:
    return ReviewFlag(code=code, message=message, severity=severity)


def detect_home_leg(trip: Trip) -> bool:
    """Detect personal/home legs using conservative normalized keywords."""
    values = (trip.normalized_origin, trip.normalized_destination)
    keywords = ("home", "apartment", "apt", "residence")
    return any(any(keyword in value for keyword in keywords) for value in values)


def detect_multi_stop(trip: Trip) -> bool:
    """Detect free-text indications that route order cannot be inferred."""
    text = f" {trip.normalized_origin} {trip.normalized_destination} "
    phrases = (" and ", "all stores", "all store", "multiple stops", "& ")
    return any(phrase in text for phrase in phrases)


def detect_ambiguous(trip: Trip, ambiguous_locations: Iterable[str]) -> bool:
    """Return true when either endpoint has more than one candidate site."""
    candidates = {value.strip().lower() for value in ambiguous_locations if value.strip()}
    return trip.normalized_origin in candidates or trip.normalized_destination in candidates


def evaluate_trip(
    trip: Trip,
    *,
    policy: Policy = Policy(),
    ambiguous_locations: Iterable[str] = (),
) -> ReviewDecision:
    """Evaluate a trip without guessing route order or replacing matrix miles."""
    flags: list[ReviewFlag] = []

    if detect_home_leg(trip):
        flags.append(_flag(
            FlagCode.HOME_LEG,
            "home/apartment leg requires HR review",
        ))

    if detect_multi_stop(trip):
        flags.append(_flag(
            FlagCode.MULTI_STOP,
            "multi-stop claim requires manual itemization",
        ))

    if detect_ambiguous(trip, ambiguous_locations):
        flags.append(_flag(
            FlagCode.AMBIGUOUS,
            "location maps to multiple possible sites",
        ))

    if abs(trip.variance_miles) > policy.mileage_variance_limit:
        flags.append(_flag(
            FlagCode.MILEAGE_VARIANCE,
            "claimed mileage differs materially from matrix mileage",
        ))

    if policy.require_positive_distance and trip.matrix_miles <= 0:
        flags.append(_flag(
            FlagCode.INVALID_INPUT,
            "matrix distance must be positive",
            "error",
        ))

    if policy.hold_on_zero_matrix_miles and trip.matrix_miles == 0:
        flags.append(_flag(
            FlagCode.INVALID_INPUT,
            "zero matrix distance requires review",
        ))

    if flags:
        return ReviewDecision(
            status=TripStatus.HOLD,
            flags=tuple(flags),
            approved_miles=0.0,
            approved_amount=0.0,
        )

    return ReviewDecision(
        status=TripStatus.APPROVED,
        flags=(),
        approved_miles=trip.reimbursable_miles,
        approved_amount=trip.reimbursement,
    )


def evaluate_batch(
    trips: Sequence[Trip],
    *,
    policy: Policy = Policy(),
    ambiguous_locations: Iterable[str] = (),
) -> tuple[ReviewDecision, ...]:
    """Evaluate a batch while preserving input order."""
    ambiguous = tuple(ambiguous_locations)
    return tuple(
        evaluate_trip(trip, policy=policy, ambiguous_locations=ambiguous)
        for trip in trips
    )


def policy_summary(policy: Policy) -> dict[str, object]:
    """Expose stable policy settings for audit/report output."""
    return {
        "mileage_variance_limit": policy.mileage_variance_limit,
        "require_positive_distance": policy.require_positive_distance,
        "hold_on_zero_matrix_miles": policy.hold_on_zero_matrix_miles,
        "matrix_distance_authoritative": True,
        "home_legs_auto_approved": False,
        "multi_stop_auto_approved": False,
        "ambiguous_locations_auto_approved": False,
    }


def decision_reason(decision: ReviewDecision) -> str:
    """Create a concise deterministic explanation for an HR reviewer."""
    if decision.status is TripStatus.APPROVED:
        return "approved using authoritative matrix miles"
    if decision.flags:
        return "; ".join(flag.message for flag in decision.flags)
    return "manual review required"


def filter_review_required(
    decisions: Iterable[ReviewDecision],
) -> tuple[ReviewDecision, ...]:
    """Return decisions requiring human action."""
    return tuple(
        decision
        for decision in decisions
        if decision.status in {TripStatus.PENDING, TripStatus.HOLD}
    )


def assert_policy_safe(decision: ReviewDecision) -> None:
    """Guard against accidental approval of a flagged decision."""
    if decision.status is TripStatus.APPROVED and decision.flags:
        raise ValueError("flagged trip cannot be auto-approved")


def apply_human_decision(
    decision: ReviewDecision,
    *,
    approved: bool,
    reviewer_note: str = "",
) -> ReviewDecision:
    """Convert a held decision into an explicit human approval/rejection."""
    if decision.status not in {TripStatus.PENDING, TripStatus.HOLD}:
        raise ValueError("only pending or held decisions can be reviewed")
    if not reviewer_note.strip():
        raise ValueError("reviewer note is required for a human decision")
    if approved:
        return ReviewDecision(
            status=TripStatus.APPROVED,
            flags=(),
            approved_miles=decision.approved_miles,
            approved_amount=decision.approved_amount,
            reviewer_note=reviewer_note.strip(),
        )
    return ReviewDecision(
        status=TripStatus.REJECTED,
        flags=decision.flags,
        approved_miles=0.0,
        approved_amount=0.0,
        reviewer_note=reviewer_note.strip(),
    )


def explain_flags(flags: Sequence[ReviewFlag]) -> list[str]:
    """Return stable human-readable flag messages."""
    return [f"{flag.code.value}: {flag.message}" for flag in flags]


def policy_has_blocking_flags(flags: Iterable[ReviewFlag]) -> bool:
    """Return true when any flag is an error-level finding."""
    return any(flag.severity == "error" for flag in flags)


def validate_policy_inputs(
    *,
    mileage_variance_limit: float,
    require_positive_distance: bool,
    hold_on_zero_matrix_miles: bool,
) -> Policy:
    """Validate raw policy values before constructing a Policy object."""
    if mileage_variance_limit < 0:
        raise ValueError("mileage_variance_limit must be non-negative")
    return Policy(
        mileage_variance_limit=float(mileage_variance_limit),
        require_positive_distance=bool(require_positive_distance),
        hold_on_zero_matrix_miles=bool(hold_on_zero_matrix_miles),
    )
