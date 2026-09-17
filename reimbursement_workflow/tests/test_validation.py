from __future__ import annotations

import math

import pytest

from reimbursement_workflow.scripts.validation import (
    is_finite_non_negative,
    normalize_text,
    reimbursement_amount,
    validate_rate,
    validate_trip_distance,
)


def test_normalize_text_is_deterministic():
    assert normalize_text("  Main   Street ") == "main street"


def test_non_negative_validation_rejects_nan_and_negative_values():
    assert is_finite_non_negative(0.0)
    assert not is_finite_non_negative(-0.1)
    assert not is_finite_non_negative(math.nan)


def test_trip_distance_and_rate_are_validated():
    assert validate_trip_distance("12.5") == 12.5
    assert validate_rate(0) == 0.0

    with pytest.raises(ValueError):
        validate_trip_distance(-1)
    with pytest.raises(ValueError):
        validate_rate(math.inf)


def test_reimbursement_amount():
    assert reimbursement_amount(10, 0.75) == 7.5
