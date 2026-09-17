"""Deterministic location matching helpers for reimbursement imports."""
from __future__ import annotations

from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Iterable, Sequence

from domain import Location
from validation import normalize_text


@dataclass(frozen=True)
class MatchCandidate:
    location: Location
    score: float
    reason: str


@dataclass(frozen=True)
class MatchResult:
    query: str
    candidates: tuple[MatchCandidate, ...]

    @property
    def best(self) -> MatchCandidate | None:
        if len(self.candidates) != 1:
            return None
        return self.candidates[0]

    @property
    def ambiguous(self) -> bool:
        return len(self.candidates) > 1


def exact_candidates(query: str, locations: Iterable[Location]) -> tuple[MatchCandidate, ...]:
    normalized = normalize_text(query)
    found = []
    for location in locations:
        if normalized in location.normalized_aliases:
            found.append(MatchCandidate(location, 1.0, "exact alias"))
    return tuple(sorted(found, key=lambda item: item.location.code))


def fuzzy_candidates(
    query: str,
    locations: Sequence[Location],
    *,
    threshold: float = 0.82,
) -> tuple[MatchCandidate, ...]:
    if not 0 <= threshold <= 1:
        raise ValueError("threshold must be between 0 and 1")
    normalized = normalize_text(query)
    matches = []
    for location in locations:
        score = max(
            SequenceMatcher(None, normalized, alias).ratio()
            for alias in location.normalized_aliases
        )
        if score >= threshold:
            matches.append(MatchCandidate(location, round(score, 4), "fuzzy alias"))
    return tuple(sorted(matches, key=lambda item: (-item.score, item.location.code)))


def match_location(
    query: str,
    locations: Sequence[Location],
    *,
    fuzzy_threshold: float = 0.92,
) -> MatchResult:
    """Prefer exact matches; return all candidates instead of guessing."""
    exact = exact_candidates(query, locations)
    if exact:
        return MatchResult(normalize_text(query), exact)
    fuzzy = fuzzy_candidates(query, locations, threshold=fuzzy_threshold)
    return MatchResult(normalize_text(query), fuzzy)


def require_unique_match(result: MatchResult) -> Location:
    """Return a unique match or raise instead of silently selecting one."""
    if not result.candidates:
        raise LookupError(f"no location match for {result.query!r}")
    if len(result.candidates) > 1:
        codes = ", ".join(item.location.code for item in result.candidates)
        raise LookupError(f"ambiguous location {result.query!r}: {codes}")
    return result.candidates[0].location


def candidate_scores(result: MatchResult) -> list[dict[str, object]]:
    return [
        {"code": item.location.code, "name": item.location.name, "score": item.score, "reason": item.reason}
        for item in result.candidates
    ]
