"""Pure validation helpers for reimbursement workflow inputs."""

from __future__ import annotations

import math
from typing import Any


def normalize_text(value: Any) -> str:
    """Normalize arbitrary input for deterministic matching."""
    return " ".join(str(value or "").strip().lower().split())


def is_finite_non_negative(value: float) -> bool:
    """Return whether a numeric distance/value is finite and non-negative."""
    return math.isfinite(value) and value >= 0


def validate_trip_distance(distance: float) -> float:
    """Validate and return a usable trip distance in miles."""
    distance = float(distance)
    if not is_finite_non_negative(distance):
        raise ValueError("trip distance must be finite and non-negative")
    return distance


def validate_rate(rate: float) -> float:
    """Validate a reimbursement rate."""
    rate = float(rate)
    if not is_finite_non_negative(rate):
        raise ValueError("reimbursement rate must be finite and non-negative")
    return rate


def reimbursement_amount(distance: float, rate: float) -> float:
    """Calculate reimbursement from validated distance and rate."""
    return validate_trip_distance(distance) * validate_rate(rate)
