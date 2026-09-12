from __future__ import annotations

from dataclasses import dataclass
from difflib import SequenceMatcher
from hashlib import sha256
import re
import unicodedata
from typing import Iterable, Mapping

from .models import ExtractedRecord


@dataclass(frozen=True)
class Destination:
    """A declared destination and every user-approved name for it."""

    key: str
    aliases: tuple[str, ...]


@dataclass(frozen=True)
class MatchCandidate:
    destination_key: str
    matched_alias: str
    rule: str
    confidence: float


@dataclass(frozen=True)
class MatchResult:
    record_id: str
    record: ExtractedRecord
    status: str
    destination_key: str | None
    rule: str | None
    confidence: float
    candidates: tuple[MatchCandidate, ...] = ()
    reasons: tuple[str, ...] = ()


def stable_record_id(record: ExtractedRecord) -> str:
    """Return an identity that survives row moves and workbook reordering."""

    parts = (
        record.source,
        record.identifier,
        record.event_date,
        record.location,
        record.category,
    )
    payload = "\x1f".join(_canonical_identity_part(part) for part in parts)
    return f"rec_{sha256(payload.encode('utf-8')).hexdigest()[:20]}"


def preallocate_records(
    records: Iterable[ExtractedRecord],
    destinations: Iterable[Destination],
    *,
    strict: bool = True,
    confirmations: Mapping[str, str] | None = None,
    fuzzy_threshold: float = 0.78,
) -> list[MatchResult]:
    """Assign each record to at most one destination before any workbook write.

    Exact, case-insensitive and normalized matches may be accepted. Fuzzy matches
    are suggestions only and always remain in review, including in non-strict
    mode. Reusable human confirmations are keyed by ``stable_record_id``.
    """

    destination_list = tuple(destinations)
    _validate_destinations(destination_list)
    destination_keys = {item.key for item in destination_list}
    saved_confirmations = confirmations or {}
    seen_record_ids: set[str] = set()
    results: list[MatchResult] = []

    for record in records:
        record_id = stable_record_id(record)
        if record_id in seen_record_ids:
            results.append(
                MatchResult(
                    record_id,
                    record,
                    "duplicate",
                    None,
                    None,
                    0.0,
                    reasons=("duplicate_record_id",),
                )
            )
            continue
        seen_record_ids.add(record_id)

        confirmed_key = saved_confirmations.get(record_id)
        if confirmed_key is not None:
            if confirmed_key in destination_keys:
                results.append(
                    MatchResult(
                        record_id,
                        record,
                        "matched",
                        confirmed_key,
                        "human_confirmation",
                        1.0,
                    )
                )
            else:
                results.append(
                    MatchResult(
                        record_id,
                        record,
                        "review",
                        None,
                        None,
                        0.0,
                        reasons=("confirmation_target_missing",),
                    )
                )
            continue

        results.append(
            _match_one(record_id, record, destination_list, strict, fuzzy_threshold)
        )

    return results


def review_results(results: Iterable[MatchResult]) -> list[MatchResult]:
    """Return every record that must not be written automatically."""

    return [result for result in results if result.status != "matched"]


def _match_one(
    record_id: str,
    record: ExtractedRecord,
    destinations: tuple[Destination, ...],
    strict: bool,
    fuzzy_threshold: float,
) -> MatchResult:
    value = record.location
    rules = (
        ("configured_exact", lambda alias: alias == value, 1.0),
        ("casefold_exact", lambda alias: alias.casefold() == value.casefold(), 0.99),
        ("normalized_exact", lambda alias: _normalize(alias) == _normalize(value), 0.97),
    )
    for rule, predicate, confidence in rules:
        matches = _matching_destinations(destinations, predicate, rule, confidence)
        if matches:
            return _resolved_result(record_id, record, matches, rule, confidence)

    if strict:
        return MatchResult(
            record_id,
            record,
            "unmatched",
            None,
            None,
            0.0,
            reasons=("no_strict_match",),
        )

    candidates = _fuzzy_candidates(value, destinations, fuzzy_threshold)
    return MatchResult(
        record_id,
        record,
        "review" if candidates else "unmatched",
        None,
        "fuzzy_candidate" if candidates else None,
        candidates[0].confidence if candidates else 0.0,
        candidates=candidates,
        reasons=("fuzzy_requires_review",) if candidates else ("no_match",),
    )


def _matching_destinations(
    destinations: tuple[Destination, ...], predicate, rule: str, confidence: float
) -> tuple[MatchCandidate, ...]:
    matches: list[MatchCandidate] = []
    for destination in destinations:
        aliases = (destination.key, *destination.aliases)
        matched_alias = next((alias for alias in aliases if predicate(alias)), None)
        if matched_alias is not None:
            matches.append(MatchCandidate(destination.key, matched_alias, rule, confidence))
    return tuple(matches)


def _resolved_result(
    record_id: str,
    record: ExtractedRecord,
    matches: tuple[MatchCandidate, ...],
    rule: str,
    confidence: float,
) -> MatchResult:
    if len(matches) > 1:
        return MatchResult(
            record_id,
            record,
            "conflict",
            None,
            rule,
            confidence,
            candidates=matches,
            reasons=("multiple_destinations",),
        )
    match = matches[0]
    return MatchResult(
        record_id,
        record,
        "matched",
        match.destination_key,
        rule,
        confidence,
        candidates=matches,
    )


def _fuzzy_candidates(
    value: str, destinations: tuple[Destination, ...], threshold: float
) -> tuple[MatchCandidate, ...]:
    normalized_value = _normalize(value)
    candidates: list[MatchCandidate] = []
    for destination in destinations:
        scored = [
            (SequenceMatcher(None, normalized_value, _normalize(alias)).ratio(), alias)
            for alias in (destination.key, *destination.aliases)
        ]
        score, alias = max(scored, default=(0.0, ""))
        if score >= threshold:
            candidates.append(
                MatchCandidate(destination.key, alias, "fuzzy_candidate", score)
            )
    return tuple(sorted(candidates, key=lambda item: (-item.confidence, item.destination_key)))


def _normalize(value: str) -> str:
    text = unicodedata.normalize("NFKC", value).casefold()
    return "".join(character for character in text if character.isalnum())


def _canonical_identity_part(value: str) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", value).strip()).casefold()


def _validate_destinations(destinations: tuple[Destination, ...]) -> None:
    keys = [item.key for item in destinations]
    if len(keys) != len(set(keys)):
        raise ValueError("Destination keys must be unique")
    if any(not key.strip() for key in keys):
        raise ValueError("Destination keys must not be empty")
