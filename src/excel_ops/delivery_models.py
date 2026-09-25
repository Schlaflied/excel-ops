"""Result and target types shared by the delivery pipeline modules."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Mapping, Sequence

from .ambiguity import ConfirmationBatch
from .delivery_manifest import DeliveryManifest
from .delivery_verification import PeriodExpectation, VerificationFinding
from .formula_delivery import FormulaDeliveryRule
from .formula_verification import FormulaVerifier
from .idempotency import RunDecision
from .matching import Destination
from .models import ExtractedRecord
from .template_writer import TemplateMapping


#: Every field a delivery row may map.  A mapping naming anything else is a
#: blocking planning error rather than a silent empty cell.
CONTRACT_FIELDS = (
    "record_id",
    "location",
    "event_date",
    "identifier",
    "category",
    "confidence",
    "amount",
    "currency",
    "source",
    "source_file",
    "source_sheet",
    "source_row",
    "source_region",
)

WRITTEN = "written"
ACCEPTED = "accepted"
REVIEW = "review"
REJECTED = "rejected"
SKIPPED_EXISTING = "skipped_existing"


class DeliveryPlanError(ValueError):
    """Raised when a declared delivery target cannot be planned at all."""


@dataclass(frozen=True)
class DeliveryTarget:
    """One declared destination plus the template it is written into."""

    destination: Destination
    template_path: str | Path
    mapping: TemplateMapping
    required_fields: tuple[str, ...] = ()
    record_id_field: str = "record_id"
    period_expectations: tuple[PeriodExpectation, ...] = field(default_factory=tuple)
    formula_verifier: FormulaVerifier | None = None
    delivery_name: str | None = None
    formula_rules: tuple[FormulaDeliveryRule, ...] = field(default_factory=tuple)

    @property
    def key(self) -> str:
        return self.destination.key


@dataclass(frozen=True)
class SourceTrace:
    """Where a record came from, before any workbook was touched."""

    locator: str
    source_file: str
    source_sheet: str
    source_row: int | None
    source_region: str

    @classmethod
    def of(cls, record: ExtractedRecord) -> "SourceTrace":
        return cls(
            record.source,
            record.source_file,
            record.source_sheet,
            record.source_row,
            record.source_region,
        )


@dataclass(frozen=True)
class WrittenCell:
    """A delivered cell traced back to its source and stable record ID."""

    record_id: str
    sheet: str
    cell: str
    field: str
    value: Any
    source: SourceTrace

    def to_dict(self) -> dict[str, Any]:
        return {
            "record_id": self.record_id,
            "sheet": self.sheet,
            "cell": self.cell,
            "field": self.field,
            "value": self.value,
            "source": asdict(self.source),
        }


@dataclass(frozen=True)
class RecordOutcome:
    """The single terminal state of one input record."""

    record_id: str
    status: str
    destination_key: str | None
    match_status: str
    rule: str | None
    confidence: float
    reasons: tuple[str, ...]
    source: SourceTrace
    values: Mapping[str, Any]
    candidates: tuple[str, ...] = field(default_factory=tuple)
    ambiguity_key: str | None = None
    cells: tuple[WrittenCell, ...] = field(default_factory=tuple)

    def to_dict(self) -> dict[str, Any]:
        return {
            "record_id": self.record_id,
            "status": self.status,
            "destination_key": self.destination_key,
            "match_status": self.match_status,
            "rule": self.rule,
            "confidence": self.confidence,
            "reasons": list(self.reasons),
            "source": asdict(self.source),
            "values": dict(self.values),
            "candidates": list(self.candidates),
            "ambiguity_key": self.ambiguity_key,
            "cells": [item.to_dict() for item in self.cells],
        }


@dataclass(frozen=True)
class PlannedTarget:
    """What a target is expected to do, produced before any file is touched."""

    destination_key: str
    template_path: str
    sheet: str
    field_mapping: Mapping[str, str | int]
    staged_path: str
    delivery_path: str
    expected_written: int
    expected_review: int
    expected_skipped_existing: int
    #: Rows the writer is predicted to refuse because a mapped target cell
    #: already holds a formula or is a merged-range follower.  Those rows are
    #: excluded from ``expected_written`` so the plan never overstates the write.
    expected_skipped_protected: int = 0
    protected_cells: tuple[str, ...] = field(default_factory=tuple)
    blocking_items: tuple[str, ...] = field(default_factory=tuple)
    formula_rules: tuple[Mapping[str, Any], ...] = field(default_factory=tuple)

    def to_dict(self) -> dict[str, Any]:
        return {
            "destination_key": self.destination_key,
            "template_path": self.template_path,
            "sheet": self.sheet,
            "field_mapping": {name: str(column) for name, column in self.field_mapping.items()},
            "staged_path": self.staged_path,
            "delivery_path": self.delivery_path,
            "expected_written": self.expected_written,
            "expected_review": self.expected_review,
            "expected_skipped_existing": self.expected_skipped_existing,
            "expected_skipped_protected": self.expected_skipped_protected,
            "protected_cells": list(self.protected_cells),
            "blocking_items": list(self.blocking_items),
            "formula_rules": [dict(item) for item in self.formula_rules],
        }


@dataclass(frozen=True)
class DeliveryPlan:
    """The checkable dry-run contract for a run."""

    inputs: tuple[str, ...]
    counts: Mapping[str, int]
    targets: tuple[PlannedTarget, ...]
    unresolved_ambiguities: tuple[str, ...] = field(default_factory=tuple)
    blocking_items: tuple[str, ...] = field(default_factory=tuple)

    @property
    def writable(self) -> bool:
        return not self.blocking_items

    def to_dict(self) -> dict[str, Any]:
        return {
            "inputs": list(self.inputs),
            "counts": dict(self.counts),
            "targets": [item.to_dict() for item in self.targets],
            "unresolved_ambiguities": list(self.unresolved_ambiguities),
            "blocking_items": list(self.blocking_items),
            "writable": self.writable,
        }


@dataclass(frozen=True)
class DeliveryFailure:
    """A machine-readable reason a target or the run did not deliver."""

    code: str
    message: str
    suggestion: str
    destination_key: str | None = None
    findings: tuple[VerificationFinding, ...] = field(default_factory=tuple)

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "message": self.message,
            "suggestion": self.suggestion,
            "destination_key": self.destination_key,
            "findings": [asdict(item) for item in self.findings],
        }


@dataclass(frozen=True)
class TargetOutcome:
    """What actually happened to one target, judged by the persisted file."""

    destination_key: str
    template_path: str
    delivered: bool
    staged_path: str | None = None
    delivery_path: str | None = None
    change_log_path: str | None = None
    verification_report_path: str | None = None
    written_record_ids: tuple[str, ...] = field(default_factory=tuple)
    written_cells: int = 0
    skipped_writes: tuple[Mapping[str, Any], ...] = field(default_factory=tuple)
    findings: tuple[VerificationFinding, ...] = field(default_factory=tuple)
    status: str = "planned"
    format_policy: Mapping[str, Any] = field(default_factory=dict)
    formula_evidence: Mapping[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "destination_key": self.destination_key,
            "template_path": self.template_path,
            "delivered": self.delivered,
            "status": self.status,
            "staged_path": self.staged_path,
            "delivery_path": self.delivery_path,
            "change_log_path": self.change_log_path,
            "verification_report_path": self.verification_report_path,
            "written_record_ids": list(self.written_record_ids),
            "written_cells": self.written_cells,
            "skipped_writes": [dict(item) for item in self.skipped_writes],
            "findings": [asdict(item) for item in self.findings],
            "format_policy": dict(self.format_policy),
            "formula_evidence": dict(self.formula_evidence),
        }


@dataclass(frozen=True)
class AmbiguityOutcome:
    key: str
    field_name: str
    question: str
    status: str
    selected: str | None
    decision_source: str | None
    candidates: tuple[str, ...]
    record_ids: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "field": self.field_name,
            "question": self.question,
            "status": self.status,
            "selected": self.selected,
            "decision_source": self.decision_source,
            "candidates": list(self.candidates),
            "record_ids": list(self.record_ids),
        }


@dataclass(frozen=True)
class DeliveryRun:
    """The single JSON-serializable result of one delivery attempt."""

    delivered: bool
    dry_run: bool
    plan: DeliveryPlan
    counts: Mapping[str, int]
    records: tuple[RecordOutcome, ...]
    targets: tuple[TargetOutcome, ...]
    ambiguities: tuple[AmbiguityOutcome, ...] = field(default_factory=tuple)
    failures: tuple[DeliveryFailure, ...] = field(default_factory=tuple)
    recipe_path: str | None = None
    #: The live confirmation batch, so callers can build reusable Recipe
    #: decisions with ``ambiguity.decide(...)``. Not part of ``to_dict()``.
    confirmation_batch: ConfirmationBatch | None = field(default=None, repr=False)
    #: True when whole-run idempotency short-circuited this call: nothing
    #: changed since a previously successful run, so no work was redone.
    no_op: bool = False
    #: The idempotency verdict for this call, when ``idempotency`` was supplied.
    run_decision: RunDecision | None = field(default=None, repr=False)
    #: One source-and-verification Manifest per *delivered* output (issue #20),
    #: generated from this run's own outcomes and from the persisted file.  A run
    #: that delivered nothing carries none.
    manifests: tuple[DeliveryManifest, ...] = field(default_factory=tuple)

    @property
    def manifest_paths(self) -> tuple[str, ...]:
        return tuple(str(item.manifest_path()) for item in self.manifests)

    @property
    def fingerprint(self) -> str | None:
        """The run fingerprint, when whole-run idempotency was enabled."""

        return self.run_decision.fingerprint.digest if self.run_decision else None

    @property
    def delivery_paths(self) -> tuple[str, ...]:
        return tuple(
            item.delivery_path for item in self.targets if item.delivered and item.delivery_path
        )

    def records_by_status(self, status: str) -> tuple[RecordOutcome, ...]:
        return tuple(item for item in self.records if item.status == status)

    def to_dict(self) -> dict[str, Any]:
        return {
            "delivered": self.delivered,
            "dry_run": self.dry_run,
            "plan": self.plan.to_dict(),
            "counts": dict(self.counts),
            "records": [item.to_dict() for item in self.records],
            "targets": [item.to_dict() for item in self.targets],
            "ambiguities": [item.to_dict() for item in self.ambiguities],
            "failures": [item.to_dict() for item in self.failures],
            "recipe_path": self.recipe_path,
            "delivery_paths": list(self.delivery_paths),
            "no_op": self.no_op,
            "run_decision": self.run_decision.to_dict() if self.run_decision else None,
            "manifests": [item.to_dict() for item in self.manifests],
            "manifest_paths": list(self.manifest_paths),
        }


def _delivery_path(target: DeliveryTarget, delivery: Path) -> Path:
    """The single place a target's delivered filename is derived."""

    template = Path(target.template_path)
    return delivery / (
        target.delivery_name or f"{template.stem}-{target.key}{template.suffix}"
    )


def _counts(outcomes: Sequence[RecordOutcome]) -> dict[str, int]:
    counts = {
        "input": len(outcomes),
        WRITTEN: 0,
        ACCEPTED: 0,
        REVIEW: 0,
        REJECTED: 0,
        SKIPPED_EXISTING: 0,
    }
    for item in outcomes:
        counts[item.status] += 1
    return counts
