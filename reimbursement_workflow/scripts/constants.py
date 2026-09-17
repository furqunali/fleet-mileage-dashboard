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


def required_export_fields() -> tuple[str, ...]:
    """Return the immutable export field contract."""
    return EXPORT_FIELDS


def distance_field_contract() -> dict[str, str]:
    """Describe which mileage field is authoritative at each boundary."""
    return {
        "claimed": CLAIMED_DISTANCE_FIELD,
        "authoritative": AUTHORITATIVE_DISTANCE_FIELD,
        "approved": APPROVED_DISTANCE_FIELD,
    }


def amount_field_contract() -> str:
    """Return the field used for approved reimbursement totals."""
    return APPROVED_AMOUNT_FIELD


def workflow_identity() -> dict[str, str]:
    """Return stable identity metadata for logs and audit records."""
    return {
        "name": WORKFLOW_NAME,
        "version": WORKFLOW_VERSION,
    }


def validate_status(status: str) -> str:
    """Validate and return a workflow status string."""
    normalized = str(status).strip().lower()
    if normalized not in REVIEW_REQUIRED_STATUSES | FINAL_STATUSES:
        raise ValueError(f"unsupported workflow status: {status!r}")
    return normalized
