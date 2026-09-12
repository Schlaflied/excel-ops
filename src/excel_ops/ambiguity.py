from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
from typing import Any, Iterable, Mapping

from .review_pack import ReviewPackRow


VALID_SCOPES = ("this-run", "project")
UNKNOWN = "unknown"


@dataclass(frozen=True)
class Ambiguity:
    field: str
    question: str
    samples: tuple[Any, ...]
    candidates: tuple[str, ...]
    affected_count: int
    recommendation: str | None
    recommendation_reason: str
    confidence: float
    source_locations: tuple[str, ...] = ()

    @property
    def key(self) -> str:
        return f"{self.field}:{self.question}"


@dataclass(frozen=True)
class RecipeDecision:
    field: str
    question: str
    selected: str
    scope: str
    source: str
    decided_at: str
    candidates: tuple[str, ...]
    custom: bool = False

    @property
    def key(self) -> str:
        return f"{self.field}:{self.question}"


@dataclass(frozen=True)
class ConfirmationItem:
    ambiguity: Ambiguity
    status: str
    selected: str | None = None
    decision_source: str | None = None
    conflict_reason: str | None = None


@dataclass(frozen=True)
class ConfirmationBatch:
    items: tuple[ConfirmationItem, ...]

    @property
    def unresolved_blockers(self) -> tuple[str, ...]:
        return tuple(item.ambiguity.key for item in self.items if item.status in {"pending", "unknown", "conflict"})

    @property
    def delivery_blocked(self) -> bool:
        return bool(self.unresolved_blockers)

    def manifest_decisions(self) -> tuple[dict[str, Any], ...]:
        """Return provenance records consumable by the future #20 manifest."""
        return tuple(
            {
                "field": item.ambiguity.field,
                "question": item.ambiguity.question,
                "selected": item.selected or UNKNOWN,
                "status": item.status,
                "source": item.decision_source or "unresolved",
            }
            for item in self.items
        )


def group_ambiguities(items: Iterable[Ambiguity]) -> tuple[Ambiguity, ...]:
    """Aggregate repeated field/rule questions into one confirmation item."""
    grouped: dict[str, list[Ambiguity]] = {}
    for item in items:
        grouped.setdefault(item.key, []).append(item)
    output: list[Ambiguity] = []
    for group in grouped.values():
        first = group[0]
        candidates = tuple(dict.fromkeys(value for item in group for value in item.candidates))
        samples = tuple(dict.fromkeys(str(value) for item in group for value in item.samples))
        locations = tuple(dict.fromkeys(value for item in group for value in item.source_locations))
        recommendations = {item.recommendation for item in group}
        recommendation = first.recommendation if len(recommendations) == 1 else None
        reason = first.recommendation_reason if recommendation else "输入之间存在不同推荐，需要人工确认"
        output.append(
            Ambiguity(
                first.field,
                first.question,
                samples,
                candidates,
                sum(item.affected_count for item in group),
                recommendation,
                reason,
                min(item.confidence for item in group),
                locations,
            )
        )
    return tuple(output)


def build_confirmation_batch(
    ambiguities: Iterable[Ambiguity],
    *,
    run_decisions: Mapping[str, RecipeDecision] | None = None,
    project_recipe: Mapping[str, RecipeDecision] | None = None,
) -> ConfirmationBatch:
    run_decisions = run_decisions or {}
    project_recipe = project_recipe or {}
    results: list[ConfirmationItem] = []
    for ambiguity in group_ambiguities(ambiguities):
        saved = run_decisions.get(ambiguity.key) or project_recipe.get(ambiguity.key)
        if saved is None:
            results.append(ConfirmationItem(ambiguity, "pending"))
        elif saved.selected == UNKNOWN:
            results.append(
                ConfirmationItem(ambiguity, "unknown", UNKNOWN, f"{saved.scope}:{saved.source}")
            )
        elif set(saved.candidates) != set(ambiguity.candidates) or ambiguity.recommendation is None:
            results.append(
                ConfirmationItem(
                    ambiguity,
                    "conflict",
                    None,
                    saved.source,
                    "当前候选或输入证据与保存决定时冲突",
                )
            )
        else:
            provenance = f"{saved.scope}:{saved.source}"
            results.append(ConfirmationItem(ambiguity, "resolved", saved.selected, provenance))
    return ConfirmationBatch(tuple(results))


def decide(
    ambiguity: Ambiguity,
    selected: str | None,
    *,
    scope: str,
    source: str = "user",
    decided_at: str | None = None,
    allow_custom: bool = False,
) -> RecipeDecision:
    if scope not in VALID_SCOPES:
        raise ValueError(f"Invalid Recipe scope: {scope}")
    normalized = str(selected).strip() if selected is not None else UNKNOWN
    if not normalized or normalized.casefold() in {"unknown", "cannot determine", "不知道"}:
        normalized = UNKNOWN
    custom = normalized != UNKNOWN and normalized not in ambiguity.candidates
    if custom and not allow_custom:
        raise ValueError(f"Decision for {ambiguity.key} must be one of the current candidates")
    return RecipeDecision(
        ambiguity.field,
        ambiguity.question,
        normalized,
        scope,
        source,
        decided_at or datetime.now(timezone.utc).isoformat(),
        ambiguity.candidates,
        custom,
    )


def save_project_recipe(decisions: Iterable[RecipeDecision], path: str | Path) -> Path:
    project = [item for item in decisions if item.scope == "project"]
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    payload = {"format": "excel-ops-recipe-v1", "decisions": [asdict(item) for item in project]}
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output


def load_project_recipe(path: str | Path) -> dict[str, RecipeDecision]:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if payload.get("format") != "excel-ops-recipe-v1":
        raise ValueError("Unsupported Recipe format")
    decisions: dict[str, RecipeDecision] = {}
    for raw in payload.get("decisions", []):
        item = RecipeDecision(**{**raw, "candidates": tuple(raw.get("candidates", ()))})
        if item.scope != "project":
            raise ValueError("Persisted Recipe decisions must use project scope")
        decisions[item.key] = item
    return decisions


def review_rows_from_ambiguities(items: Iterable[Ambiguity]) -> list[ReviewPackRow]:
    """Adapt grouped questions to the existing #18 offline review contract."""
    rows: list[ReviewPackRow] = []
    for item in group_ambiguities(items):
        samples = ", ".join(str(value) for value in item.samples[:5])
        locations = ", ".join(item.source_locations[:5]) or f"field:{item.field}"
        rows.append(
            ReviewPackRow(
                record_id=f"ambiguity:{item.key}",
                source_value=samples,
                source_location=locations,
                proposed_value=item.recommendation,
                review_reason=(
                    f"{item.question}; affected={item.affected_count}; "
                    f"reason={item.recommendation_reason}"
                ),
                confidence=item.confidence,
                candidates=item.candidates,
            )
        )
    return rows
