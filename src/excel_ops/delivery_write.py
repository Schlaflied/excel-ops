"""Staged template writes and independent verification before delivery."""

from __future__ import annotations

from dataclasses import asdict, replace
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

from .delivery_manifest import read_delivery_manifest
from .delivery_models import (
    ACCEPTED,
    WRITTEN,
    DeliveryFailure,
    DeliveryPlan,
    DeliveryTarget,
    PlannedTarget,
    RecordOutcome,
    TargetOutcome,
    WrittenCell,
)
from .delivery_verification import (
    DeliveryContract,
    DeliveryVerificationError,
    StageCounts,
    verify_and_deliver,
)
from .formula_delivery import (
    FormulaDeliveryError,
    apply_formula_delivery,
    formula_contract_fingerprint,
)
from .template_writer import TemplateWriteError, TemplateWriteResult, write_template


def _contract(target: DeliveryTarget, accepted: Sequence[RecordOutcome], withheld: int) -> DeliveryContract:
    verifiers = (target.formula_verifier,) if target.formula_verifier is not None else ()
    return DeliveryContract(
        counts=StageCounts(
            input=len(accepted) + withheld,
            accepted=len(accepted),
            review=withheld,
            written=len(accepted),
        ),
        required_fields=target.required_fields,
        expected_record_ids=tuple(item.record_id for item in accepted),
        record_id_field=target.record_id_field,
        period_expectations=target.period_expectations,
        verifiers=verifiers,
    )


def _write_and_verify(
    outcomes: Sequence[RecordOutcome],
    targets: Sequence[DeliveryTarget],
    plan: DeliveryPlan,
    post_stage_hook: Callable[[Path], None] | None,
) -> tuple[list[RecordOutcome], list[TargetOutcome], list[DeliveryFailure]]:
    # Keyed by position: a duplicate shares its stable record ID with the original.
    final = list(outcomes)
    target_outcomes: list[TargetOutcome] = []
    failures: list[DeliveryFailure] = []

    for target in targets:
        planned = next(item for item in plan.targets if item.destination_key == target.key)
        positions = [
            index
            for index, item in enumerate(outcomes)
            if item.status == ACCEPTED and item.destination_key == target.key
        ]
        accepted = [outcomes[index] for index in positions]
        if not accepted:
            outcome, failure = _no_new_records(target, planned)
        else:
            outcome, failure, traced = _write_target(
                target, planned, accepted, post_stage_hook
            )
            for offset, position in enumerate(positions):
                if offset in traced:
                    final[position] = replace(
                        final[position], status=WRITTEN, cells=traced[offset]
                    )
        if failure is not None:
            failures.append(failure)
        target_outcomes.append(outcome)

    return final, target_outcomes, failures


def _no_new_records(
    target: DeliveryTarget, planned: PlannedTarget
) -> tuple[TargetOutcome, DeliveryFailure | None]:
    """Settle a target that received no accepted records this run."""

    existing_path = planned.delivery_path if Path(planned.delivery_path).is_file() else None
    if not target.formula_rules:
        return (
            TargetOutcome(
                target.key,
                planned.template_path,
                False,
                delivery_path=existing_path,
                status="no_new_records",
            ),
            None,
        )

    # Formula rules still have to hold for the file already delivered: only a
    # verified Manifest for the same formula contract lets this target rest.
    existing_manifest = read_delivery_manifest(planned.delivery_path)
    expected_contract = formula_contract_fingerprint(target.formula_rules)
    if (
        existing_manifest is not None
        and existing_manifest.verification.passed
        and existing_manifest.reconciled
        and existing_manifest.formulas.get("status") == "verified"
        and existing_manifest.formulas.get("contract_fingerprint") == expected_contract
    ):
        return (
            TargetOutcome(
                target.key,
                planned.template_path,
                False,
                delivery_path=planned.delivery_path,
                status="no_new_records",
                formula_evidence=dict(existing_manifest.formulas),
            ),
            None,
        )
    return (
        TargetOutcome(
            target.key,
            planned.template_path,
            False,
            delivery_path=existing_path,
            status="formula_delivery_failed",
            formula_evidence={"status": "failed"},
        ),
        DeliveryFailure(
            "formula_delivery_requires_records",
            f"{target.key}: formula rules changed, but no new records were accepted for a staged delivery.",
            "Provide a new delivery input, or migrate the existing workbook in a separate reviewed operation.",
            target.key,
        ),
    )


def _write_target(
    target: DeliveryTarget,
    planned: PlannedTarget,
    accepted: Sequence[RecordOutcome],
    post_stage_hook: Callable[[Path], None] | None,
) -> tuple[TargetOutcome, DeliveryFailure | None, dict[int, tuple[WrittenCell, ...]]]:
    """Stage-write one target, apply its formulas, and deliver only if verified.

    The third value maps each accepted row's position to its delivered cells;
    it is empty unless the target was delivered.
    """

    rows = [dict(item.values) for item in accepted]
    try:
        write_result = write_template(
            target.template_path, rows, target.mapping, output_path=planned.staged_path
        )
    except TemplateWriteError as error:
        return (
            TargetOutcome(target.key, planned.template_path, False, status="write_failed"),
            DeliveryFailure(
                "template_write_failed",
                f"{target.key}: {error}",
                "Correct the template mapping or the staging path, then re-run.",
                target.key,
            ),
            {},
        )

    staged = Path(write_result.output_path)
    record_ids = tuple(item.record_id for item in accepted)
    skipped = tuple(asdict(item) for item in write_result.skipped)
    if write_result.skipped:
        # A mapped cell the writer refused to touch means an accepted record
        # was not fully written. Fail closed instead of delivering a partial row.
        reasons = sorted({item.reason for item in write_result.skipped})
        return (
            TargetOutcome(
                target.key,
                planned.template_path,
                False,
                str(staged),
                None,
                str(write_result.change_log_path),
                None,
                record_ids,
                len(write_result.changes),
                skipped,
                (),
                "incomplete_write",
            ),
            DeliveryFailure(
                "incomplete_write",
                f"{target.key}: the writer skipped {len(write_result.skipped)} mapped cell(s): {', '.join(reasons)}.",
                "Adjust the template mapping or unprotect the declared region; never deliver a partially written row.",
                target.key,
            ),
            {},
        )

    formula_evidence: Mapping[str, Any] = {}
    if target.formula_rules:
        try:
            formula_evidence = apply_formula_delivery(staged, target.formula_rules)
        except FormulaDeliveryError as error:
            return (
                TargetOutcome(
                    target.key,
                    planned.template_path,
                    False,
                    str(staged),
                    None,
                    str(write_result.change_log_path),
                    status="formula_delivery_failed",
                    formula_evidence={"status": "failed"},
                ),
                DeliveryFailure(
                    "formula_delivery_failed",
                    f"{target.key}: {error}",
                    "Review the formula plan, expectations, and calculation engine, then re-run.",
                    target.key,
                ),
                {},
            )

    if post_stage_hook is not None:
        post_stage_hook(staged)

    contract = _contract(target, accepted, planned.expected_review)
    try:
        verification = verify_and_deliver(write_result, planned.delivery_path, contract)
    except DeliveryVerificationError as error:
        return (
            TargetOutcome(
                target.key,
                planned.template_path,
                False,
                str(staged),
                None,
                str(write_result.change_log_path),
                str(error.result.report_path),
                record_ids,
                len(write_result.changes),
                skipped,
                error.result.findings,
                "verification_failed",
            ),
            DeliveryFailure(
                "verification_failed",
                f"{target.key}: {error}",
                "Read the verification report, fix the cause, and deliver again.",
                target.key,
                error.result.findings,
            ),
            {},
        )

    return (
        TargetOutcome(
            target.key,
            planned.template_path,
            True,
            str(staged),
            str(verification.delivery_path),
            str(write_result.change_log_path),
            str(verification.report_path),
            record_ids,
            len(write_result.changes),
            skipped,
            verification.findings,
            "delivered",
            write_result.format_policy,
            formula_evidence,
        ),
        None,
        _trace(accepted, write_result),
    )


def _trace(
    accepted: Sequence[RecordOutcome], write_result: TemplateWriteResult
) -> dict[int, tuple[WrittenCell, ...]]:
    """Map each accepted row's position to the delivered cells it produced."""

    traced: dict[int, list[WrittenCell]] = {index: [] for index in range(len(accepted))}
    for change in write_result.changes:
        record = accepted[change.row_index]
        traced[change.row_index].append(
            WrittenCell(
                record.record_id,
                change.sheet,
                change.cell,
                change.field,
                change.new_value,
                record.source,
            )
        )
    return {key: tuple(value) for key, value in traced.items()}
