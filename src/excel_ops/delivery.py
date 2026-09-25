"""End-to-end delivery orchestration for declared workbook templates.

This module contains no new matching, ambiguity, write-back or verification
logic.  It only sequences the already-merged modules into one agent-callable
pipeline and returns a single JSON-serializable result:

``ingest -> validate -> strict match -> ambiguity/Recipe -> plan ->
staged template write -> independent verification of the persisted file ->
delivery``

A successful ``write_template`` call is never reported as a delivery.  Only a
target whose reopened file passed :func:`verify_and_deliver` is published.

Passing ``idempotency=IdempotencyOptions(...)`` additionally turns on the
whole-run short-circuit from :mod:`excel_ops.idempotency`: the run fingerprint is
computed *before* any matching, writing or verification, and a fingerprint that
matches a previously successful run returns a ``no_op`` result immediately.  That
is a different mechanism from the per-record deduplication below
(``_existing_record_ids`` / ``_build_plan`` / ``_mark_existing``), which stops one
row being appended twice inside a target workbook; both stay in force.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

from .ambiguity import RecipeDecision, build_confirmation_batch, save_project_recipe
from .delivery_ambiguity import (
    _ambiguities_from_matches,
    _ambiguity_outcomes,
    _block_unresolved,
    _confirmations,
    _load_recipe,
)
from .delivery_manifest import (
    DeliveryManifest,
    build_delivery_manifests,
    read_delivery_manifest,
    recorded_exports_are_current,
    write_delivery_manifests,
)
from .delivery_models import (
    ACCEPTED,
    CONTRACT_FIELDS,
    REJECTED,
    REVIEW,
    SKIPPED_EXISTING,
    WRITTEN,
    AmbiguityOutcome,
    DeliveryFailure,
    DeliveryPlan,
    DeliveryPlanError,
    DeliveryRun,
    DeliveryTarget,
    PlannedTarget,
    RecordOutcome,
    SourceTrace,
    TargetOutcome,
    WrittenCell,
    _counts,
    _delivery_path,
)
from .delivery_planning import _build_plan, _classify, _mark_existing, _planned_outcome
from .delivery_verification import PeriodExpectation
from .delivery_write import _write_and_verify
from .formula_delivery import FormulaDeliveryError, formula_rule_from_mapping
from .idempotency import (
    CHANGED,
    FAILED,
    NO_OP,
    RETRY,
    STARTED,
    SUCCEEDED,
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
from .ingestion import LayoutDetectionError, load_input_records
from .matching import Destination
from .models import ExtractedRecord
from .number_formats import FormatPolicy, NumberFormatPolicyError, policy_manifest
from .export_formats import ExportFormatError, parse_export_formats
from .periods import PeriodResult
from .template_writer import TemplateMapping

# ``excel_ops.delivery`` stays the one public import path for the pipeline.
__all__ = [
    "ACCEPTED",
    "CONTRACT_FIELDS",
    "REJECTED",
    "REVIEW",
    "SKIPPED_EXISTING",
    "WRITTEN",
    "AmbiguityOutcome",
    "DeliveryFailure",
    "DeliveryPlan",
    "DeliveryPlanError",
    "DeliveryRun",
    "DeliveryTarget",
    "PlannedTarget",
    "RecordOutcome",
    "SourceTrace",
    "TargetOutcome",
    "WrittenCell",
    "load_delivery_targets",
    "plan_delivery",
    "run_delivery",
]


def run_delivery(
    inputs: Sequence[str | Path],
    targets: Sequence[DeliveryTarget],
    *,
    staging_dir: str | Path,
    delivery_dir: str | Path,
    confidence_threshold: float = 0.85,
    recipe_path: str | Path | None = None,
    decisions: Sequence[RecipeDecision] = (),
    dry_run: bool = False,
    post_stage_hook: Callable[[Path], None] | None = None,
    artifact_hook: Callable[
        [DeliveryRun, tuple[DeliveryManifest, ...]], tuple[DeliveryManifest, ...]
    ]
    | None = None,
    artifact_fingerprint: Mapping[str, Any] | None = None,
    idempotency: IdempotencyOptions | None = None,
    period: PeriodResult | None = None,
    write_manifest: bool = True,
) -> DeliveryRun:
    """Run ingest to verified delivery for declared targets and return one result.

    ``post_stage_hook`` receives each staged workbook path after the write and
    before verification.  It exists so callers and tests can inspect or corrupt
    the staged artifact and prove that verification, not the writer, decides
    whether a file is delivered.

    ``artifact_hook`` runs after the verified workbooks and their in-memory
    Manifests exist, but before evidence is written and idempotency is marked
    successful.  It lets format adapters attach their evidence atomically to
    the delivery result: an export failure makes the whole run retryable.

    ``idempotency`` opts this call into the whole-run short-circuit.  The run
    fingerprint is computed before any matching, write or verification; when it
    matches a previously *successful* run the call returns a ``no_op`` result
    without redoing the work.  A dry run never short-circuits -- it still returns
    the full plan -- but it does report the verdict in ``run_decision``.

    ``period`` is the resolved report period this run delivers for.  It is
    recorded in the #20 Manifest as ``period_start``/``period_end``; the period
    values written *into* a workbook are still declared per target as
    ``period_expectations`` and verified by #4.

    Every delivered output gets one Manifest in ``result.manifests``, built from
    this run's own outcomes and from the persisted file.  ``write_manifest=True``
    (the default) also writes ``<delivered file>.manifest.json`` and
    ``.manifest.txt`` beside it.  A dry run and a no-op produce no Manifest,
    because no file was delivered.
    """

    if not targets:
        raise DeliveryPlanError("at least one delivery target is required")
    _validate_targets(targets)
    staging = Path(staging_dir)
    delivery = Path(delivery_dir)

    failures: list[DeliveryFailure] = []
    project_recipe, recipe_failure = _load_recipe(recipe_path)
    guard = _RunGuard.of(
        inputs,
        targets,
        delivery,
        project_recipe,
        decisions,
        idempotency,
        artifact_fingerprint,
    )
    if guard is not None and guard.decision.no_op and not dry_run:
        outputs = tuple(
            output
            for target in targets
            for output in (_delivery_path(target, delivery),)
            if output.is_file()
        )
        if artifact_hook is None or recorded_exports_are_current(outputs):
            return _no_op_run(inputs, targets, delivery, guard.decision, recipe_path)
        guard.decision = RunDecision(
            RETRY,
            "recorded_export_missing_or_changed",
            guard.decision.fingerprint,
            guard.previous,
        )

    records, record_inputs, ingest_failures = _ingest(inputs)
    failures.extend(ingest_failures)

    outcomes, match_results = _classify(records, targets, confidence_threshold)
    ambiguities, ambiguity_records = _ambiguities_from_matches(match_results)
    if recipe_failure is not None:
        failures.append(recipe_failure)
    run_decisions = {item.key: item for item in decisions if item.scope == "this-run"}
    project_decisions = {item.key: item for item in decisions if item.scope == "project"}
    batch = build_confirmation_batch(
        ambiguities,
        run_decisions=run_decisions,
        project_recipe={**project_recipe, **project_decisions},
    )
    confirmations = _confirmations(batch, ambiguity_records, targets)
    if confirmations:
        # Re-run matching with the reusable human decisions. The first batch is
        # kept as the provenance of those decisions.
        outcomes, _ = _classify(
            records, targets, confidence_threshold, confirmations=confirmations
        )

    outcomes = _block_unresolved(outcomes, batch, ambiguity_records)
    # A dry run must not touch a single file, and a real run must never drop a
    # decision an earlier run already recorded: merge the loaded Recipe with
    # this call's decisions so the file only ever grows or updates in place.
    if not dry_run and recipe_path is not None and project_decisions:
        save_project_recipe({**project_recipe, **project_decisions}.values(), recipe_path)

    plan, existing_ids = _build_plan(inputs, outcomes, targets, staging, delivery, batch)
    outcomes = _mark_existing(outcomes, existing_ids)
    plan = replace(plan, counts=_counts(outcomes))

    if dry_run or not plan.writable or failures:
        if not plan.writable:
            failures.append(
                DeliveryFailure(
                    "plan_blocked",
                    f"The delivery plan is blocked by {len(plan.blocking_items)} item(s).",
                    "Resolve every blocking item, then re-run the delivery.",
                )
            )
        if guard is not None and not dry_run:
            # A blocked or failed run is recorded as a failure, never as a
            # baseline a later identical run could short-circuit against.
            guard.finish(FAILED, False, _counts(outcomes), failures, ())
        return DeliveryRun(
            False,
            dry_run,
            plan,
            _counts(outcomes),
            tuple(outcomes),
            tuple(_planned_outcome(item, plan) for item in targets),
            tuple(_ambiguity_outcomes(batch, ambiguity_records)),
            tuple(failures),
            str(recipe_path) if recipe_path else None,
            batch,
            False,
            guard.decision if guard else None,
        )

    if guard is not None:
        # Mark the attempt before the first write, so a run interrupted between
        # here and the end is found as ``started`` and retried rather than being
        # mistaken for a completed no-op baseline.
        guard.start()
    outcomes, target_outcomes, write_failures = _write_and_verify(
        outcomes, targets, plan, post_stage_hook
    )
    failures.extend(write_failures)
    delivered = bool(target_outcomes) and all(
        item.delivered or item.status == "no_new_records" for item in target_outcomes
    ) and not failures
    run = DeliveryRun(
        delivered,
        False,
        plan,
        _counts(outcomes),
        tuple(outcomes),
        tuple(target_outcomes),
        tuple(_ambiguity_outcomes(batch, ambiguity_records)),
        tuple(failures),
        str(recipe_path) if recipe_path else None,
        batch,
        False,
        guard.decision if guard else None,
    )
    # The Manifest is generated last, from the run that just finished and from
    # the file that is now on disk.  It is evidence about a completed delivery,
    # never a parallel bookkeeping system that could drift from it.
    manifests = build_delivery_manifests(
        run,
        targets,
        period=period,
        recipe_decisions={
            **project_recipe,
            **project_decisions,
            **run_decisions,
        },
        record_input_paths=record_inputs,
    )
    if artifact_hook is not None:
        built_outputs = {Path(manifest.output).resolve() for manifest in manifests}
        recovered = tuple(
            manifest
            for outcome in target_outcomes
            if outcome.delivery_path
            for manifest in (read_delivery_manifest(outcome.delivery_path),)
            if manifest is not None and Path(manifest.output).resolve() not in built_outputs
        )
        manifests = (*manifests, *recovered)
    if artifact_hook is not None and delivered:
        manifest_outputs = {Path(manifest.output).resolve() for manifest in manifests}
        missing_manifests = tuple(
            outcome.delivery_path
            for outcome in target_outcomes
            if outcome.delivery_path
            and Path(outcome.delivery_path).is_file()
            and Path(outcome.delivery_path).resolve() not in manifest_outputs
        )
        if missing_manifests:
            failures.append(
                DeliveryFailure(
                    "manifest_recovery_failed",
                    "A Manifest could not be recovered for an existing workbook.",
                    "Restore or rebuild the Manifest, then re-run the delivery.",
                )
            )
            if guard is not None:
                guard.finish(FAILED, False, _counts(outcomes), failures, target_outcomes)
            return replace(
                run,
                delivered=False,
                failures=tuple(failures),
                manifests=(),
            )
    if artifact_hook is not None and delivered:
        try:
            manifests = artifact_hook(run, manifests)
        except (OSError, ValueError) as error:
            artifact_failure = DeliveryFailure(
                "format_export_failed",
                f"A selected delivery format could not be exported: {error}",
                "Correct the export scope or install a supported PDF renderer, then re-run.",
            )
            failures = [*failures, artifact_failure]
            if write_manifest and manifests:
                try:
                    write_delivery_manifests(manifests)
                except (OSError, ValueError) as manifest_error:
                    failures.append(
                        DeliveryFailure(
                            "manifest_write_failed",
                            f"The recovery manifest could not be written: {manifest_error}",
                            "Check the delivery directory's permissions and re-run.",
                        )
                    )
            if guard is not None:
                guard.finish(FAILED, False, _counts(outcomes), failures, target_outcomes)
            return replace(
                run,
                delivered=False,
                failures=tuple(failures),
                manifests=(),
            )
    if write_manifest:
        try:
            write_delivery_manifests(manifests)
        except (OSError, ValueError) as error:
            # A delivered workbook with an unwritten manifest is not a
            # completed delivery.  The idempotency record must never say
            # SUCCEEDED here: a matching fingerprint on the next run would
            # otherwise short-circuit to a no-op and never regenerate the
            # missing manifest.  Report it the same way every other failure
            # in this function is reported, and let ``guard.finish`` record
            # the failure so the next identical run is a ``retry``.
            manifest_failure = DeliveryFailure(
                "manifest_write_failed",
                f"The delivery manifest could not be written: {error}",
                "Check the delivery directory's permissions and re-run the delivery.",
            )
            failures = [*failures, manifest_failure]
            if guard is not None:
                guard.finish(FAILED, False, _counts(outcomes), failures, target_outcomes)
            return replace(
                run,
                delivered=False,
                failures=tuple(failures),
                manifests=(),
            )
    if guard is not None:
        guard.finish(
            SUCCEEDED if delivered else FAILED,
            delivered,
            _counts(outcomes),
            failures,
            target_outcomes,
        )
    return replace(run, manifests=manifests)


def plan_delivery(
    inputs: Sequence[str | Path],
    targets: Sequence[DeliveryTarget],
    *,
    staging_dir: str | Path,
    delivery_dir: str | Path,
    confidence_threshold: float = 0.85,
    recipe_path: str | Path | None = None,
    decisions: Sequence[RecipeDecision] = (),
    idempotency: IdempotencyOptions | None = None,
) -> DeliveryRun:
    """Return the checkable plan without touching a single file."""

    return run_delivery(
        inputs,
        targets,
        staging_dir=staging_dir,
        delivery_dir=delivery_dir,
        confidence_threshold=confidence_threshold,
        recipe_path=recipe_path,
        decisions=decisions,
        dry_run=True,
        idempotency=idempotency,
    )


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


def _validate_targets(targets: Sequence[DeliveryTarget]) -> None:
    keys = [item.key for item in targets]
    if len(keys) != len(set(keys)):
        raise DeliveryPlanError("delivery target destination keys must be unique")


def _ingest(
    inputs: Sequence[str | Path],
) -> tuple[list[ExtractedRecord], list[str], list[DeliveryFailure]]:
    """Ingest every input and remember which input each record came from.

    The third return value is aligned with the records: entry *i* is the input
    path record *i* was read from.  A record's own ``source_file`` names the
    provenance it *declares* — an extraction JSON's records name the image, not
    the JSON — so only the ingester knows which declared input produced it.  The
    #20 Manifest needs that to hash the right file per source.
    """

    records: list[ExtractedRecord] = []
    origins: list[str] = []
    failures: list[DeliveryFailure] = []
    for item in inputs:
        path = Path(item)
        try:
            loaded = load_input_records(path)
        except LayoutDetectionError as error:
            failures.append(
                DeliveryFailure(
                    "unknown_layout",
                    f"{path.name}: {error}",
                    "Confirm the header row or declare a mapping before re-running.",
                )
            )
        except (OSError, ValueError) as error:
            failures.append(
                DeliveryFailure(
                    "unreadable_input",
                    f"{path.name}: {error}",
                    "Supply a readable .xlsx, .xlsm, .csv, or extraction .json input.",
                )
            )
        else:
            records.extend(loaded)
            origins.extend(str(item) for _ in loaded)
    return records, origins, failures


def load_delivery_targets(
    payload: Mapping[str, Any], *, base_dir: str | Path = "."
) -> tuple[list[DeliveryTarget], dict[str, Any]]:
    """Build targets and run options from a declarative JSON configuration."""

    base = Path(base_dir)
    raw_targets = payload.get("targets")
    if not isinstance(raw_targets, list) or not raw_targets:
        raise DeliveryPlanError("configuration must contain a non-empty targets array")
    targets: list[DeliveryTarget] = []
    workbook_policy_supplied = "format_policy" in payload
    workbook_format_policy = payload.get("format_policy")
    for raw in raw_targets:
        if not isinstance(raw, Mapping):
            raise DeliveryPlanError("each target must be an object")
        key = str(raw.get("key") or "").strip()
        if not key:
            raise DeliveryPlanError("each target needs a key")
        field_columns = raw.get("field_columns")
        if not isinstance(field_columns, Mapping) or not field_columns:
            raise DeliveryPlanError(f"target {key} needs field_columns")
        target_policy_supplied = "format_policy" in raw
        policy_supplied = target_policy_supplied or workbook_policy_supplied
        raw_format_policy = raw.get("format_policy", workbook_format_policy)
        try:
            if not policy_supplied:
                format_policy = None
            elif isinstance(raw_format_policy, Mapping):
                format_policy = FormatPolicy.from_mapping(raw_format_policy)
            else:
                raise DeliveryPlanError(f"target {key} format_policy must be an object")
        except NumberFormatPolicyError as error:
            raise DeliveryPlanError(f"target {key} format_policy: {error}") from error
        mapping = TemplateMapping(
            sheet=str(raw.get("sheet") or ""),
            field_columns=dict(field_columns),
            header_row=int(raw.get("header_row", 1)),
            data_start_row=int(raw.get("data_start_row", 2)),
            template_type=str(raw.get("template_type", "table")),
            style_source_row=raw.get("style_source_row"),
            max_rows=raw.get("max_rows"),
            format_policy=format_policy,
        )
        expectations = tuple(
            PeriodExpectation(
                str(item.get("sheet") or mapping.sheet),
                str(item.get("cell")),
                item.get("expected"),
                str(item.get("field", "report_period")),
            )
            for item in raw.get("period_expectations", ())
            if isinstance(item, Mapping)
        )
        raw_formulas = raw.get("formulas", ())
        if not isinstance(raw_formulas, (list, tuple)):
            raise DeliveryPlanError(f"target {key} formulas must be an array")
        try:
            formula_rules = tuple(
                formula_rule_from_mapping(item)
                for item in raw_formulas
                if isinstance(item, Mapping)
            )
        except FormulaDeliveryError as error:
            raise DeliveryPlanError(f"target {key} formulas: {error}") from error
        if len(formula_rules) != len(raw_formulas):
            raise DeliveryPlanError(f"target {key} formulas must contain objects")
        targets.append(
            DeliveryTarget(
                destination=Destination(key, tuple(str(item) for item in raw.get("aliases", ()))),
                template_path=base / str(raw.get("template") or ""),
                mapping=mapping,
                required_fields=tuple(str(item) for item in raw.get("required_fields", ())),
                record_id_field=str(raw.get("record_id_field", "record_id")),
                period_expectations=expectations,
                delivery_name=raw.get("delivery_name"),
                formula_rules=formula_rules,
            )
        )
    inputs = payload.get("inputs")
    if not isinstance(inputs, list) or not inputs:
        raise DeliveryPlanError("configuration must contain a non-empty inputs array")
    try:
        export_selection = parse_export_formats(payload)
    except ExportFormatError as error:
        raise DeliveryPlanError(str(error)) from error
    options: dict[str, Any] = {
        "inputs": [base / str(item) for item in inputs],
        "staging_dir": base / str(payload.get("staging_dir") or "staging"),
        "delivery_dir": base / str(payload.get("delivery_dir") or "delivery"),
        "confidence_threshold": float(payload.get("confidence_threshold", 0.85)),
        "export_formats": export_selection.formats,
        "export_selection": export_selection.to_dict(),
    }
    if payload.get("recipe"):
        options["recipe_path"] = base / str(payload["recipe"])
    return targets, options
