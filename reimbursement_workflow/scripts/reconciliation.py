"""Reconciliation services for mileage and reimbursement totals.

The matrix remains authoritative. These helpers deliberately never replace
matrix miles with employee-entered odometer values.
"""
from __future__ import annotations

from dataclasses import dataclass
from math import isfinite
from typing import Iterable, Sequence

from domain import ReviewDecision, Trip, TripStatus, summarize


@dataclass(frozen=True)
class ReconciliationIssue:
    """One discrepancy found while checking a reimbursement run."""

    code: str
    message: str
    expected: float | None = None
    actual: float | None = None

    def __post_init__(self) -> None:
        if not self.code.strip() or not self.message.strip():
            raise ValueError("reconciliation issue requires code and message")


@dataclass(frozen=True)
class ReconciliationResult:
    """Immutable result of a reconciliation pass."""

    ok: bool
    issues: tuple[ReconciliationIssue, ...] = ()

    @property
    def issue_count(self) -> int:
        return len(self.issues)


def compare_numbers(
    expected: float,
    actual: float,
    *,
    tolerance: float = 0.01,
    code: str = "numeric_mismatch",
) -> ReconciliationIssue | None:
    """Return an issue when two finite numbers differ beyond tolerance."""
    if not isfinite(expected) or not isfinite(actual):
        return ReconciliationIssue(code, "values must be finite", expected, actual)
    if tolerance < 0:
        raise ValueError("tolerance must be non-negative")
    if abs(expected - actual) > tolerance:
        return ReconciliationIssue(code, "values differ beyond tolerance", expected, actual)
    return None


def check_trip_authority(trip: Trip, *, max_claim_variance: float = 10.0) -> tuple[ReconciliationIssue, ...]:
    """Check that authoritative matrix miles are present and claim variance is visible."""
    if max_claim_variance < 0:
        raise ValueError("max_claim_variance must be non-negative")
    issues: list[ReconciliationIssue] = []
    if trip.matrix_miles < 0:
        issues.append(ReconciliationIssue(
            "negative_matrix_miles",
            "matrix miles cannot be negative",
            0.0,
            trip.matrix_miles,
        ))
    if abs(trip.variance_miles) > max_claim_variance:
        issues.append(ReconciliationIssue(
            "claimed_matrix_variance",
            "claimed miles differ materially from authoritative matrix miles",
            trip.matrix_miles,
            trip.claimed_miles,
        ))
    return tuple(issues)


def reconcile_decisions(
    decisions: Sequence[ReviewDecision],
    *,
    expected_miles: float,
    expected_amount: float,
    tolerance: float = 0.01,
) -> ReconciliationResult:
    """Reconcile approved mileage and amount against expected totals."""
    summary = summarize("reconciliation", decisions)
    issues: list[ReconciliationIssue] = []
    miles_issue = compare_numbers(
        expected_miles,
        summary.approved_miles,
        tolerance=tolerance,
        code="approved_miles_mismatch",
    )
    amount_issue = compare_numbers(
        expected_amount,
        summary.approved_amount,
        tolerance=tolerance,
        code="approved_amount_mismatch",
    )
    if miles_issue:
        issues.append(miles_issue)
    if amount_issue:
        issues.append(amount_issue)
    return ReconciliationResult(ok=not issues, issues=tuple(issues))


def reconcile_trips(trips: Iterable[Trip]) -> ReconciliationResult:
    """Validate every trip's authoritative-distance boundary."""
    issues: list[ReconciliationIssue] = []
    for index, trip in enumerate(trips):
        for issue in check_trip_authority(trip):
            issues.append(
                ReconciliationIssue(
                    code=f"trip_{index}_{issue.code}",
                    message=issue.message,
                    expected=issue.expected,
                    actual=issue.actual,
                )
            )
    return ReconciliationResult(ok=not issues, issues=tuple(issues))


def ensure_approved_only(decisions: Iterable[ReviewDecision]) -> tuple[ReviewDecision, ...]:
    """Return approved decisions, preserving deterministic input order."""
    return tuple(d for d in decisions if d.status is TripStatus.APPROVED)


def approved_totals(decisions: Iterable[ReviewDecision]) -> tuple[float, float]:
    """Return approved miles and amount as rounded reporting totals."""
    approved = ensure_approved_only(decisions)
    return (
        round(sum(d.approved_miles for d in approved), 2),
        round(sum(d.approved_amount for d in approved), 2),
    )


def assert_reconciled(result: ReconciliationResult) -> None:
    """Raise a concise error at a reporting boundary when reconciliation fails."""
    if not result.ok:
        details = "; ".join(issue.code for issue in result.issues)
        raise ValueError(f"reconciliation failed: {details}")


def status_counts(decisions: Iterable[ReviewDecision]) -> dict[str, int]:
    """Count workflow statuses using stable string keys."""
    counts = {status.value: 0 for status in TripStatus}
    for decision in decisions:
        counts[decision.status.value] += 1
    return counts


def validate_decision(decision: ReviewDecision) -> None:
    """Validate decision invariants before a decision is persisted."""
    if decision.status is TripStatus.APPROVED and decision.requires_human_review:
        raise ValueError("approved decision cannot require human review")
    if decision.status is TripStatus.REJECTED and decision.approved_amount:
        raise ValueError("rejected decision cannot contain approved amount")
    if decision.status is TripStatus.HOLD and decision.approved_amount:
        raise ValueError("held decision cannot contain approved amount")
