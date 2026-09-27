"""Whole-run idempotency guard and no-op result for delivery runs."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

from .ambiguity import RecipeDecision
from .delivery_models import (
    DeliveryFailure,
    DeliveryPlan,
    DeliveryRun,
    DeliveryTarget,
    PlannedTarget,
    TargetOutcome,
    _counts,
    _delivery_path,
)
from .idempotency import (
    CHANGED,
    NO_OP,
    STARTED,
    ConnectorTarget,
    IdempotencyError,
    IdempotencyOptions,
    RunDecision,
    RunFingerprint,
    RunRecord,
    compute_fingerprint,
    evaluate_run,
    load_run_record,
    record_run,
)
from .number_formats import policy_manifest


def _period_payload(targets: Sequence[DeliveryTarget], period: Any) -> dict[str, Any]:
    """The declared report period: the explicit one plus the in-workbook banners."""

    return {
        "declared": period,
        "expectations": sorted(
            f"{item.field}@{item.sheet}!{item.cell}={item.expected}"
            for target in targets
            for item in target.period_expectations
        ),
    }


def _mapping_payload(targets: Sequence[DeliveryTarget]) -> list[dict[str, Any]]:
    """The declared mapping version: what each target promises to write, where."""

    return [
        {
            "key": target.key,
            "aliases": sorted(target.destination.aliases),
            "sheet": target.mapping.sheet,
            "field_columns": {
                name: str(column) for name, column in sorted(target.mapping.field_columns.items())
            },
            "header_row": target.mapping.header_row,
            "data_start_row": target.mapping.data_start_row,
            "template_type": target.mapping.template_type,
            "style_source_row": target.mapping.style_source_row,
            "max_rows": target.mapping.max_rows,
            "overwrite_formulas": target.mapping.overwrite_formulas,
            "format_policy": policy_manifest(target.mapping.format_policy),
            "record_id_field": target.record_id_field,
            "required_fields": sorted(target.required_fields),
            "formula_rules": [
                rule.fingerprint_payload() for rule in target.formula_rules
            ],
        }
        for target in sorted(targets, key=lambda item: item.key)
    ]


@dataclass
class _RunGuard:
    """The whole-run idempotency state for one ``run_delivery`` call."""

    options: IdempotencyOptions
    state_path: Path
    task_key: str
    connector: ConnectorTarget
    decision: RunDecision
    previous: RunRecord | None
    #: Recomputes the fingerprint against the files as they are *now*, so the
    #: record left behind describes the delivered state rather than the
    #: pre-delivery one.  Without this, the output content hash recorded before
    #: the write would never match the next run and no run could ever be a no-op.
    recompute: Callable[[], RunFingerprint]

    @classmethod
    def of(
        cls,
        inputs: Sequence[str | Path],
        targets: Sequence[DeliveryTarget],
        delivery: Path,
        project_recipe: Mapping[str, RecipeDecision],
        decisions: Sequence[RecipeDecision],
        options: IdempotencyOptions | None,
        artifact_fingerprint: Mapping[str, Any] | None,
    ) -> "_RunGuard | None":
        if options is None:
            return None
        outputs = [_delivery_path(target, delivery) for target in targets]
        connector = options.connector or ConnectorTarget.local(outputs)
        templates = [target.template_path for target in targets]
        mapping: Any = _mapping_payload(targets)
        if artifact_fingerprint is not None:
            mapping = {"targets": mapping, "artifacts": dict(artifact_fingerprint)}
        period = _period_payload(targets, options.period)

        def recompute() -> RunFingerprint:
            return compute_fingerprint(
                inputs=inputs,
                templates=templates,
                template_profile_version=options.template_profile_version,
                mapping=mapping,
                recipe_decisions=project_recipe,
                confirmations=decisions,
                period=period,
                connector=connector,
                revision_source=options.revision_source,
            )

        fingerprint = recompute()
        path = options.resolved_state_path(delivery)
        task_key = options.resolved_task_key(connector)
        try:
            previous = load_run_record(path, task_key)
        except IdempotencyError:
            # An unreadable or foreign run-state file must never authorize a
            # no-op; the run proceeds and rewrites the record.
            return cls(
                options,
                path,
                task_key,
                connector,
                RunDecision(CHANGED, "unreadable_run_state", fingerprint),
                None,
                recompute,
            )
        decision = evaluate_run(
            fingerprint,
            previous,
            connector=connector,
            revision_source=options.revision_source,
        )
        return cls(options, path, task_key, connector, decision, previous, recompute)

    def start(self) -> None:
        record_run(
            self.decision.fingerprint,
            self.state_path,
            task_key=self.task_key,
            status=STARTED,
            delivered=False,
            detail={"stage": "write_and_verify"},
            previous=self.previous,
        )

    def finish(
        self,
        status: str,
        delivered: bool,
        counts: Mapping[str, int],
        failures: Sequence[DeliveryFailure],
        target_outcomes: Sequence[TargetOutcome],
    ) -> RunRecord:
        # Only aggregate, non-identifying detail is persisted: counts, failure
        # codes, and how many targets were delivered.  No path, no destination
        # name, no cell value ever reaches the run record.
        detail = {
            "counts": dict(counts),
            "failure_codes": sorted({item.code for item in failures}),
            "targets": len(target_outcomes),
            "delivered_targets": sum(1 for item in target_outcomes if item.delivered),
        }
        return record_run(
            self.recompute(),
            self.state_path,
            task_key=self.task_key,
            status=status,
            delivered=delivered,
            detail=detail,
            previous=self.previous,
        )


def _no_op_run(
    inputs: Sequence[str | Path],
    targets: Sequence[DeliveryTarget],
    delivery: Path,
    decision: RunDecision,
    recipe_path: str | Path | None,
) -> DeliveryRun:
    """Return "nothing changed" without redoing matching, writing, or verifying."""

    planned: list[PlannedTarget] = []
    outcomes: list[TargetOutcome] = []
    for target in targets:
        template = Path(target.template_path)
        delivered_path = _delivery_path(target, delivery)
        planned.append(
            PlannedTarget(
                destination_key=target.key,
                template_path=str(template),
                sheet=target.mapping.sheet,
                field_mapping=dict(target.mapping.field_columns),
                staged_path="",
                delivery_path=str(delivered_path),
                expected_written=0,
                expected_review=0,
                expected_skipped_existing=0,
            )
        )
        outcomes.append(
            TargetOutcome(
                target.key,
                str(template),
                False,
                delivery_path=str(delivered_path) if delivered_path.is_file() else None,
                status=NO_OP,
            )
        )
    plan = DeliveryPlan(tuple(str(item) for item in inputs), _counts(()), tuple(planned))
    return DeliveryRun(
        False,
        False,
        plan,
        _counts(()),
        (),
        tuple(outcomes),
        (),
        (),
        str(recipe_path) if recipe_path else None,
        None,
        True,
        decision,
    )
