from datetime import date

from reimbursement_workflow.scripts.domain import Trip, TripStatus
from reimbursement_workflow.scripts.policy_engine import evaluate_trip


def make_trip(origin="Store A", destination="Store B", claimed=10.0, matrix=10.0):
    return Trip(
        employee_id="E1",
        trip_date=date(2026, 9, 1),
        origin=origin,
        destination=destination,
        claimed_miles=claimed,
        matrix_miles=matrix,
        rate=0.75,
    )


def test_clean_trip_uses_matrix_miles():
    decision = evaluate_trip(make_trip(claimed=12, matrix=10))
    assert decision.status is TripStatus.APPROVED
    assert decision.approved_miles == 10
    assert decision.approved_amount == 7.5


def test_home_leg_is_held():
    decision = evaluate_trip(make_trip(origin="Employee Home"))
    assert decision.status is TripStatus.HOLD


def test_multi_stop_is_held():
    decision = evaluate_trip(make_trip(origin="Store A and Store B"))
    assert decision.status is TripStatus.HOLD


def test_ambiguous_location_is_held():
    decision = evaluate_trip(make_trip(), ambiguous_locations=["store a"])
    assert decision.status is TripStatus.HOLD


def test_material_mileage_variance_is_held():
    decision = evaluate_trip(make_trip(claimed=30, matrix=10))
    assert decision.status is TripStatus.HOLD
