"""Stable workflow constants and schema identifiers."""

WORKFLOW_NAME = "fleet-mileage-reimbursement"
WORKFLOW_VERSION = "1.0"
AUTHORITATIVE_DISTANCE_FIELD = "matrix_miles"
CLAIMED_DISTANCE_FIELD = "claimed_miles"
APPROVED_DISTANCE_FIELD = "approved_miles"
APPROVED_AMOUNT_FIELD = "approved_amount"

REVIEW_REQUIRED_STATUSES = frozenset({"pending", "hold"})
FINAL_STATUSES = frozenset({"approved", "rejected"})
EXPORT_FIELDS = (
    "employee_id",
    "trip_date",
    "origin",
    "destination",
    "claimed_miles",
    "matrix_miles",
    "variance_miles",
    "status",
    "approved_miles",
    "approved_amount",
    "reason",
)


def is_final_status(status: str) -> bool:
    """Return whether a status represents a completed review decision."""
    return status in FINAL_STATUSES


def requires_review(status: str) -> bool:
    """Return whether a status requires human action."""
    return status in REVIEW_REQUIRED_STATUSES


def schema_version() -> str:
    """Return the public export schema version."""
    return WORKFLOW_VERSION
