"""Delivery-destination ambiguities and their Recipe confirmations."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Mapping, Sequence

from .ambiguity import (
    UNKNOWN,
    Ambiguity,
    ConfirmationBatch,
    RecipeDecision,
    load_project_recipe,
)
from .delivery_models import (
    REJECTED,
    REVIEW,
    AmbiguityOutcome,
    DeliveryFailure,
    DeliveryTarget,
    RecordOutcome,
)
from .matching import MatchResult


_AMBIGUITY_FIELD = "destination"


def _ambiguity_shape(match: MatchResult) -> tuple[str, tuple[str, ...]] | None:
    """Return the record-independent shape of a match's open question.

    ``Ambiguity.key`` is ``field:question``, and ``group_ambiguities`` merges
    every ambiguity sharing a key into one confirmation item.  The key must
    therefore describe the *kind* of question — which candidates are competing
    for which field — and must not embed the individual record's source value,
    or nothing would ever group and a saved Recipe decision would only ever
    match the one run that produced it.
    """

    if match.status in {"review", "conflict"} and match.candidates:
        candidates = tuple(dict.fromkeys(item.destination_key for item in match.candidates))
        return match.status, candidates
    return None


def _ambiguity_question(status: str, candidates: Sequence[str]) -> str:
    joined = ", ".join(candidates)
    if status == "conflict":
        return (
            f"Which declared destination owns source values claimed by "
            f"more than one of [{joined}]?"
        )
    return f"Which declared destination owns source values fuzzily matching [{joined}]?"


def _ambiguity_key(match: MatchResult) -> str | None:
    shape = _ambiguity_shape(match)
    if shape is None:
        return None
    status, candidates = shape
    return f"{_AMBIGUITY_FIELD}:{_ambiguity_question(status, candidates)}"


def _ambiguities_from_matches(
    match_results: Sequence[MatchResult],
) -> tuple[tuple[Ambiguity, ...], dict[str, tuple[str, ...]]]:
    items: list[Ambiguity] = []
    records: dict[str, list[str]] = {}
    for match in match_results:
        shape = _ambiguity_shape(match)
        if shape is None:
            continue
        status, candidates = shape
        question = _ambiguity_question(status, candidates)
        key = f"{_AMBIGUITY_FIELD}:{question}"
        recommendation = candidates[0] if status == "review" and len(candidates) == 1 else None
        items.append(
            Ambiguity(
                _AMBIGUITY_FIELD,
                question,
                (match.record.location,),
                candidates,
                1,
                recommendation,
                "fuzzy candidate requires confirmation"
                if status == "review"
                else "several declared destinations claim this value",
                match.confidence,
                (match.record.source,),
            )
        )
        records.setdefault(key, []).append(match.record_id)
    return tuple(items), {key: tuple(value) for key, value in records.items()}


def _load_recipe(
    recipe_path: str | Path | None,
) -> tuple[dict[str, RecipeDecision], DeliveryFailure | None]:
    if recipe_path is None or not Path(recipe_path).is_file():
        return {}, None
    try:
        return load_project_recipe(recipe_path), None
    except (OSError, ValueError) as error:
        return {}, DeliveryFailure(
            "unreadable_recipe",
            f"{Path(recipe_path).name}: {error}",
            "Repair or remove the project Recipe before re-running the delivery.",
        )


def _confirmations(
    batch: ConfirmationBatch,
    ambiguity_records: Mapping[str, tuple[str, ...]],
    targets: Sequence[DeliveryTarget],
) -> dict[str, str]:
    keys = {item.key for item in targets}
    confirmations: dict[str, str] = {}
    for item in batch.items:
        if item.status != "resolved" or not item.selected or item.selected == UNKNOWN:
            continue
        if item.selected not in keys:
            continue
        for record_id in ambiguity_records.get(item.ambiguity.key, ()):
            confirmations[record_id] = item.selected
    return confirmations


def _block_unresolved(
    outcomes: Sequence[RecordOutcome],
    batch: ConfirmationBatch,
    ambiguity_records: Mapping[str, tuple[str, ...]],
) -> list[RecordOutcome]:
    """Never let a record with an unresolved question reach accepted."""

    blocked: dict[str, str] = {}
    for key in batch.unresolved_blockers:
        for record_id in ambiguity_records.get(key, ()):
            blocked[record_id] = key
    updated: list[RecordOutcome] = []
    for item in outcomes:
        if item.record_id in blocked and item.status != REJECTED:
            updated.append(
                replace(
                    item,
                    status=REVIEW,
                    destination_key=None,
                    reasons=tuple(dict.fromkeys((*item.reasons, "unresolved_ambiguity"))),
                )
            )
        else:
            updated.append(item)
    return updated


def _ambiguity_outcomes(
    batch: ConfirmationBatch, ambiguity_records: Mapping[str, tuple[str, ...]]
) -> list[AmbiguityOutcome]:
    return [
        AmbiguityOutcome(
            item.ambiguity.key,
            item.ambiguity.field,
            item.ambiguity.question,
            item.status,
            item.selected,
            item.decision_source,
            item.ambiguity.candidates,
            ambiguity_records.get(item.ambiguity.key, ()),
        )
        for item in batch.items
    ]
