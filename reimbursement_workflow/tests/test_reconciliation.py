from datetime import date

from reimbursement_workflow.scripts.domain import Trip
from reimbursement_workflow.scripts.reconciliation import reconcile_trips


def make_trip(claimed, matrix):
    return Trip(
        employee_id="E1",
        trip_date=date(2026, 9, 1),
        origin="Store A",
        destination="Store B",
        claimed_miles=claimed,
        matrix_miles=matrix,
        rate=0.75,
    )


def test_reconcile_trips_accepts_authoritative_matrix_distance():
    result = reconcile_trips([make_trip(12, 10)])
    assert result.ok is True
    assert result.issue_count == 0


def test_reconcile_trips_surfaces_large_claim_variance():
    result = reconcile_trips([make_trip(40, 10)])
    assert result.ok is False
    assert result.issues[0].code.endswith("claimed_matrix_variance")
