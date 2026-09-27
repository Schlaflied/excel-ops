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

from .ambiguity import (
    ConfirmationBatch,
    RecipeDecision,
    build_confirmation_batch,
    save_project_recipe,
)
from .delivery_ambiguity import (
    _ambiguities_from_matches,
    _ambiguity_outcomes,
    _block_unresolved,
    _confirmations,
    _load_recipe,
)
from .delivery_config import load_delivery_targets
from .delivery_guard import _no_op_run, _RunGuard
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
from .delivery_write import _write_and_verify
from .idempotency import (
    FAILED,
    RETRY,
    SUCCEEDED,
    IdempotencyOptions,
    RunDecision,
)
from .ingestion import LayoutDetectionError, load_input_records
from .models import ExtractedRecord
from .periods import PeriodResult

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
    delivery = Path(delivery_dir)
    project_recipe, recipe_failure = _load_recipe(recipe_path)
    guard, no_op = _open_guard(
        inputs, targets, delivery, project_recipe, decisions,
        idempotency, artifact_fingerprint, artifact_hook, dry_run,
    )
    if no_op:
        return _no_op_run(inputs, targets, delivery, guard.decision, recipe_path)

    prepared = _prepare(
        inputs, targets, Path(staging_dir), delivery, confidence_threshold,
        recipe_path, recipe_failure, project_recipe, decisions, dry_run,
    )
    if dry_run or not prepared.plan.writable or prepared.failures:
        return _stopped_run(prepared, targets, dry_run, guard, recipe_path)

    if guard is not None:
        # Mark the attempt before the first write, so a run interrupted between
        # here and the end is found as ``started`` and retried rather than being
        # mistaken for a completed no-op baseline.
        guard.start()
    outcomes, target_outcomes, write_failures = _write_and_verify(
        prepared.outcomes, targets, prepared.plan, post_stage_hook
    )
    failures = [*prepared.failures, *write_failures]
    delivered = bool(target_outcomes) and all(
        item.delivered or item.status == "no_new_records" for item in target_outcomes
    ) and not failures
    run = _run_result(
        prepared, delivered, False, outcomes, target_outcomes, failures, recipe_path, guard
    )
    return _finish_run(
        run,
        guard,
        targets,
        period=period,
        recipe_decisions=prepared.recipe_decisions,
        record_inputs=prepared.record_inputs,
        artifact_hook=artifact_hook,
        write_manifest=write_manifest,
    )


def _open_guard(
    inputs: Sequence[str | Path],
    targets: Sequence[DeliveryTarget],
    delivery: Path,
    project_recipe: Mapping[str, RecipeDecision],
    decisions: Sequence[RecipeDecision],
    idempotency: IdempotencyOptions | None,
    artifact_fingerprint: Mapping[str, Any] | None,
    artifact_hook: Callable[..., Any] | None,
    dry_run: bool,
) -> tuple[_RunGuard | None, bool]:
    """Evaluate whole-run idempotency; the flag says the run may short-circuit."""

    guard = _RunGuard.of(
        inputs,
        targets,
        delivery,
        project_recipe,
        decisions,
        idempotency,
        artifact_fingerprint,
    )
    if guard is None or not guard.decision.no_op or dry_run:
        return guard, False
    if _recorded_outputs_current(targets, delivery, artifact_hook):
        return guard, True
    guard.decision = RunDecision(
        RETRY,
        "recorded_export_missing_or_changed",
        guard.decision.fingerprint,
        guard.previous,
    )
    return guard, False


@dataclass
class _Prepared:
    """Everything decided before the first write: routed records and the plan."""

    outcomes: list[RecordOutcome]
    plan: DeliveryPlan
    batch: ConfirmationBatch
    ambiguity_records: dict[str, tuple[str, ...]]
    failures: list[DeliveryFailure]
    record_inputs: list[str]
    recipe_decisions: dict[str, RecipeDecision]


def _prepare(
    inputs: Sequence[str | Path],
    targets: Sequence[DeliveryTarget],
    staging: Path,
    delivery: Path,
    confidence_threshold: float,
    recipe_path: str | Path | None,
    recipe_failure: DeliveryFailure | None,
    project_recipe: Mapping[str, RecipeDecision],
    decisions: Sequence[RecipeDecision],
    dry_run: bool,
) -> _Prepared:
    """Ingest, route, and plan every record without writing any workbook."""

    records, record_inputs, failures = _ingest(inputs)
    if recipe_failure is not None:
        failures.append(recipe_failure)

    run_decisions = {item.key: item for item in decisions if item.scope == "this-run"}
    project_decisions = {item.key: item for item in decisions if item.scope == "project"}
    outcomes, batch, ambiguity_records = _route_records(
        records,
        targets,
        confidence_threshold,
        run_decisions,
        {**project_recipe, **project_decisions},
    )
    # A dry run must not touch a single file, and a real run must never drop a
    # decision an earlier run already recorded: merge the loaded Recipe with
    # this call's decisions so the file only ever grows or updates in place.
    if not dry_run and recipe_path is not None and project_decisions:
        save_project_recipe({**project_recipe, **project_decisions}.values(), recipe_path)

    plan, existing_ids = _build_plan(inputs, outcomes, targets, staging, delivery, batch)
    outcomes = _mark_existing(outcomes, existing_ids)
    return _Prepared(
        outcomes,
        replace(plan, counts=_counts(outcomes)),
        batch,
        ambiguity_records,
        failures,
        record_inputs,
        {**project_recipe, **project_decisions, **run_decisions},
    )


def _stopped_run(
    prepared: _Prepared,
    targets: Sequence[DeliveryTarget],
    dry_run: bool,
    guard: _RunGuard | None,
    recipe_path: str | Path | None,
) -> DeliveryRun:
    """Return a dry run, or a run stopped before any write by a blocker or failure."""

    failures = list(prepared.failures)
    plan = prepared.plan
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
        guard.finish(FAILED, False, _counts(prepared.outcomes), failures, ())
    planned = [_planned_outcome(item, plan) for item in targets]
    return _run_result(
        prepared, False, dry_run, prepared.outcomes, planned, failures, recipe_path, guard
    )


def _run_result(
    prepared: _Prepared,
    delivered: bool,
    dry_run: bool,
    outcomes: Sequence[RecordOutcome],
    target_outcomes: Sequence[TargetOutcome],
    failures: Sequence[DeliveryFailure],
    recipe_path: str | Path | None,
    guard: _RunGuard | None,
) -> DeliveryRun:
    return DeliveryRun(
        delivered,
        dry_run,
        prepared.plan,
        _counts(outcomes),
        tuple(outcomes),
        tuple(target_outcomes),
        tuple(_ambiguity_outcomes(prepared.batch, prepared.ambiguity_records)),
        tuple(failures),
        str(recipe_path) if recipe_path else None,
        prepared.batch,
        False,
        guard.decision if guard else None,
    )


def _recorded_outputs_current(
    targets: Sequence[DeliveryTarget],
    delivery: Path,
    artifact_hook: Callable[..., Any] | None,
) -> bool:
    """Whether a no-op may trust the exports recorded beside the delivered files."""

    if artifact_hook is None:
        return True
    outputs = tuple(
        output
        for target in targets
        for output in (_delivery_path(target, delivery),)
        if output.is_file()
    )
    return recorded_exports_are_current(outputs)


def _route_records(
    records: Sequence[ExtractedRecord],
    targets: Sequence[DeliveryTarget],
    confidence_threshold: float,
    run_decisions: Mapping[str, RecipeDecision],
    project_recipe: Mapping[str, RecipeDecision],
) -> tuple[list[RecordOutcome], ConfirmationBatch, dict[str, tuple[str, ...]]]:
    """Classify every record and apply the human decisions for open ambiguities."""

    outcomes, match_results = _classify(records, targets, confidence_threshold)
    ambiguities, ambiguity_records = _ambiguities_from_matches(match_results)
    batch = build_confirmation_batch(
        ambiguities,
        run_decisions=run_decisions,
        project_recipe=project_recipe,
    )
    confirmations = _confirmations(batch, ambiguity_records, targets)
    if confirmations:
        # Re-run matching with the reusable human decisions. The first batch is
        # kept as the provenance of those decisions.
        outcomes, _ = _classify(
            records, targets, confidence_threshold, confirmations=confirmations
        )
    return _block_unresolved(outcomes, batch, ambiguity_records), batch, ambiguity_records


def _finish_run(
    run: DeliveryRun,
    guard: _RunGuard | None,
    targets: Sequence[DeliveryTarget],
    *,
    period: PeriodResult | None,
    recipe_decisions: Mapping[str, RecipeDecision],
    record_inputs: Sequence[str],
    artifact_hook: Callable[
        [DeliveryRun, tuple[DeliveryManifest, ...]], tuple[DeliveryManifest, ...]
    ]
    | None,
    write_manifest: bool,
) -> DeliveryRun:
    """Attach Manifests and export evidence, then record the run's outcome.

    The Manifest is generated last, from the run that just finished and from
    the file that is now on disk.  It is evidence about a completed delivery,
    never a parallel bookkeeping system that could drift from it.
    """

    failures = list(run.failures)
    manifests = build_delivery_manifests(
        run,
        targets,
        period=period,
        recipe_decisions=recipe_decisions,
        record_input_paths=record_inputs,
    )
    if artifact_hook is not None:
        manifests, missing_manifests = _with_recovered_manifests(run, manifests)
        if missing_manifests:
            failures.append(
                DeliveryFailure(
                    "manifest_recovery_failed",
                    "A Manifest could not be recovered for an existing workbook.",
                    "Restore or rebuild the Manifest, then re-run the delivery.",
                )
            )
            return _failed_run(run, guard, failures)
    if artifact_hook is not None and run.delivered:
        try:
            manifests = artifact_hook(run, manifests)
        except (OSError, ValueError) as error:
            failures.append(
                DeliveryFailure(
                    "format_export_failed",
                    f"A selected delivery format could not be exported: {error}",
                    "Correct the export scope or install a supported PDF renderer, then re-run.",
                )
            )
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
            return _failed_run(run, guard, failures)
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
            failures.append(
                DeliveryFailure(
                    "manifest_write_failed",
                    f"The delivery manifest could not be written: {error}",
                    "Check the delivery directory's permissions and re-run the delivery.",
                )
            )
            return _failed_run(run, guard, failures)
    if guard is not None:
        guard.finish(
            SUCCEEDED if run.delivered else FAILED,
            run.delivered,
            run.counts,
            failures,
            run.targets,
        )
    return replace(run, manifests=manifests)


def _with_recovered_manifests(
    run: DeliveryRun, manifests: tuple[DeliveryManifest, ...]
) -> tuple[tuple[DeliveryManifest, ...], bool]:
    """Add on-disk Manifests for outputs not rebuilt this run; flag any still missing.

    Only a delivered run can be missing a Manifest: the flag is always False
    otherwise.
    """

    built_outputs = {Path(manifest.output).resolve() for manifest in manifests}
    recovered = tuple(
        manifest
        for outcome in run.targets
        if outcome.delivery_path
        for manifest in (read_delivery_manifest(outcome.delivery_path),)
        if manifest is not None and Path(manifest.output).resolve() not in built_outputs
    )
    manifests = (*manifests, *recovered)
    if not run.delivered:
        return manifests, False
    manifest_outputs = {Path(manifest.output).resolve() for manifest in manifests}
    missing_manifests = tuple(
        outcome.delivery_path
        for outcome in run.targets
        if outcome.delivery_path
        and Path(outcome.delivery_path).is_file()
        and Path(outcome.delivery_path).resolve() not in manifest_outputs
    )
    return manifests, bool(missing_manifests)


def _failed_run(
    run: DeliveryRun, guard: _RunGuard | None, failures: Sequence[DeliveryFailure]
) -> DeliveryRun:
    """Record a failed run and return it without Manifests, so it is retried."""

    if guard is not None:
        guard.finish(FAILED, False, run.counts, failures, run.targets)
    return replace(run, delivered=False, failures=tuple(failures), manifests=())



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
