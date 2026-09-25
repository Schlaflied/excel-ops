"""Record classification and the pre-write delivery plan."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Any, Mapping, Sequence

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .ambiguity import ConfirmationBatch
from .cells import column_number, is_merged_non_anchor, is_protected_formula_value
from .delivery_ambiguity import _ambiguity_key
from .delivery_models import (
    ACCEPTED,
    CONTRACT_FIELDS,
    REJECTED,
    REVIEW,
    SKIPPED_EXISTING,
    DeliveryPlan,
    DeliveryPlanError,
    DeliveryTarget,
    PlannedTarget,
    RecordOutcome,
    SourceTrace,
    TargetOutcome,
    _counts,
    _delivery_path,
)
from .matching import MatchResult, preallocate_records, stable_record_id
from .models import ExtractedRecord
from .number_formats import NumberFormatPolicyError, resolve_format_policy
from .review import review_record


_REJECT_REASON_PREFIXES = ("missing:", "invalid:")


def _row_values(record_id: str, record: ExtractedRecord) -> dict[str, Any]:
    return {
        "record_id": record_id,
        "location": record.location,
        "event_date": record.event_date,
        "identifier": record.identifier,
        "category": record.category,
        "confidence": record.confidence,
        "amount": record.amount,
        "currency": record.currency,
        "source": record.source,
        "source_file": record.source_file,
        "source_sheet": record.source_sheet,
        "source_row": record.source_row,
        "source_region": record.source_region,
    }


def _classify(
    records: Sequence[ExtractedRecord],
    targets: Sequence[DeliveryTarget],
    confidence_threshold: float,
    *,
    confirmations: Mapping[str, str] | None = None,
) -> tuple[list[RecordOutcome], list[MatchResult]]:
    """Route every record to exactly one non-written state before any write."""

    reviewed = [(record, review_record(record, confidence_threshold)) for record in records]
    rejected_reasons: dict[int, tuple[str, ...]] = {}
    matchable: list[ExtractedRecord] = []
    order: list[tuple[int, ExtractedRecord, tuple[str, ...]]] = []
    for index, (record, verdict) in enumerate(reviewed):
        if any(reason.startswith(_REJECT_REASON_PREFIXES) for reason in verdict.reasons):
            rejected_reasons[index] = verdict.reasons
        else:
            matchable.append(record)
        order.append((index, record, verdict.reasons))

    match_results = preallocate_records(
        matchable,
        [item.destination for item in targets],
        strict=False,
        confirmations=confirmations,
    )
    # preallocate_records returns exactly one result per input record, in order.
    by_record = dict(zip((index for index, _, _ in order if index not in rejected_reasons), match_results))

    outcomes: list[RecordOutcome] = []
    for index, record, reasons in order:
        if index in rejected_reasons:
            record_id = stable_record_id(record)
            outcomes.append(
                RecordOutcome(
                    record_id,
                    REJECTED,
                    None,
                    "not_matched",
                    None,
                    record.confidence,
                    rejected_reasons[index],
                    SourceTrace.of(record),
                    _row_values(record_id, record),
                )
            )
            continue
        match = by_record[index]
        status = ACCEPTED if match.status == "matched" and not reasons else REVIEW
        combined = tuple(dict.fromkeys((*reasons, *match.reasons)))
        outcomes.append(
            RecordOutcome(
                match.record_id,
                status,
                match.destination_key if status == ACCEPTED else None,
                match.status,
                match.rule,
                match.confidence,
                combined,
                SourceTrace.of(record),
                _row_values(match.record_id, record),
                tuple(item.destination_key for item in match.candidates),
                _ambiguity_key(match),
            )
        )
    return outcomes, match_results


def _column_number(value: str | int) -> int:
    """Convert the shared parser's ``ValueError`` into this module's contract."""

    try:
        return column_number(value)
    except ValueError as exc:
        raise DeliveryPlanError(str(exc)) from exc


def _existing_record_ids(path: Path, target: DeliveryTarget) -> set[str]:
    """Read the record IDs already present in a previously delivered file."""

    column = target.mapping.field_columns.get(target.record_id_field)
    if not path.is_file() or column is None:
        return set()
    index = _column_number(column)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[target.mapping.sheet] if target.mapping.sheet in workbook.sheetnames else None
        if sheet is None:
            return set()
        found: set[str] = set()
        for row in sheet.iter_rows(
            min_row=target.mapping.data_start_row, min_col=index, max_col=index, values_only=True
        ):
            value = str(row[0] or "").strip()
            if value:
                found.add(value)
        return found
    finally:
        workbook.close()


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for revision in range(2, 1000):
        candidate = path.with_name(f"{path.stem}-r{revision}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise DeliveryPlanError(f"unable to allocate a staging path next to {path}")


def _target_blockers(target: DeliveryTarget) -> list[str]:
    blockers: list[str] = []
    template = Path(target.template_path)
    if not template.is_file():
        blockers.append(f"{target.key}: missing_template:{template.name}")
    unknown = [name for name in target.mapping.field_columns if name not in CONTRACT_FIELDS]
    for name in sorted(unknown):
        blockers.append(f"{target.key}: unmapped_contract_field:{name}")
    if target.record_id_field not in target.mapping.field_columns:
        blockers.append(f"{target.key}: record_id_not_mapped:{target.record_id_field}")
    missing_required = [
        name for name in target.required_fields if name not in target.mapping.field_columns
    ]
    for name in sorted(missing_required):
        blockers.append(f"{target.key}: required_field_not_mapped:{name}")
    if target.mapping.max_rows is not None and target.mapping.max_rows == 0:
        blockers.append(f"{target.key}: template_accepts_no_rows")
    return blockers


def _predicted_protected_cells(
    template: Path, target: DeliveryTarget, row_count: int
) -> tuple[tuple[str, ...], int]:
    """Predict the mapped cells ``write_template`` will refuse to touch.

    ``write_template`` only discovers a protected cell at write time, which
    would leave the plan promising rows that can never be written.  Reading the
    template's mapped cells here and applying the *same* shared predicates
    (:func:`cells.is_protected_formula_value`, :func:`cells.is_merged_non_anchor`)
    gives the plan the writer's own verdict ahead of time.

    Returns the predicted cell coordinates and the number of distinct rows they
    span, because one protected cell fails the whole row closed.
    """

    mapping = target.mapping
    if row_count <= 0 or not template.is_file():
        return (), 0
    try:
        workbook = load_workbook(template)
    except (OSError, ValueError, KeyError):
        # An unreadable template is already reported by ``_target_blockers`` and
        # by the writer; a failed pre-check must not invent a second verdict.
        return (), 0
    try:
        if mapping.sheet not in workbook.sheetnames:
            return (), 0
        worksheet = workbook[mapping.sheet]
        columns: list[int] = []
        for column in mapping.field_columns.values():
            try:
                columns.append(_column_number(column))
            except DeliveryPlanError:
                continue
        protected: list[str] = []
        affected_rows = 0
        for offset in range(row_count):
            row = mapping.data_start_row + offset
            hit = False
            for index in columns:
                cell = worksheet.cell(row, index)
                if is_merged_non_anchor(cell) or is_protected_formula_value(
                    cell.value, overwrite_formulas=mapping.overwrite_formulas
                ):
                    protected.append(f"{get_column_letter(index)}{row}")
                    hit = True
            if hit:
                affected_rows += 1
        return tuple(protected), affected_rows
    finally:
        workbook.close()


def _build_plan(
    inputs: Sequence[str | Path],
    outcomes: Sequence[RecordOutcome],
    targets: Sequence[DeliveryTarget],
    staging: Path,
    delivery: Path,
    batch: ConfirmationBatch,
) -> tuple[DeliveryPlan, dict[str, set[str]]]:
    planned: list[PlannedTarget] = []
    blocking: list[str] = []
    format_ambiguities: list[str] = []
    existing: dict[str, set[str]] = {}
    for target in targets:
        template = Path(target.template_path)
        accepted = [item for item in outcomes if item.destination_key == target.key]
        withheld = [
            item
            for item in outcomes
            if item.status == REVIEW and target.key in item.candidates
        ]
        delivery_path = _delivery_path(target, delivery)
        known = _existing_record_ids(delivery_path, target)
        existing[target.key] = known
        skipped = [item for item in accepted if item.record_id in known]
        blockers = list(_target_blockers(target))
        if target.mapping.format_policy is not None:
            try:
                resolve_format_policy(
                    target.mapping.format_policy,
                    [item.values for item in accepted if item.status == ACCEPTED],
                )
            except NumberFormatPolicyError as error:
                item = f"format_policy:{target.key}:{error.code}"
                if error.field_name:
                    item += f":{error.field_name}"
                blockers.append(item)
                format_ambiguities.append(item)
        blocking.extend(blockers)
        writable_rows = len(accepted) - len(skipped)
        protected_cells, protected_rows = _predicted_protected_cells(
            template, target, writable_rows
        )
        planned.append(
            PlannedTarget(
                destination_key=target.key,
                template_path=str(template),
                sheet=target.mapping.sheet,
                field_mapping=dict(target.mapping.field_columns),
                staged_path=str(
                    _unique_path(staging / f"{template.stem}-{target.key}{template.suffix}")
                ),
                delivery_path=str(delivery_path),
                expected_written=writable_rows - protected_rows,
                expected_review=len(withheld),
                expected_skipped_existing=len(skipped),
                expected_skipped_protected=protected_rows,
                protected_cells=protected_cells,
                blocking_items=tuple(blockers),
                formula_rules=tuple(
                    rule.manifest_payload() for rule in target.formula_rules
                ),
            )
        )
    unresolved = (*batch.unresolved_blockers, *format_ambiguities)
    plan = DeliveryPlan(
        tuple(str(item) for item in inputs),
        _counts(outcomes),
        tuple(planned),
        unresolved,
        tuple(blocking),
    )
    return plan, existing


def _mark_existing(
    outcomes: Sequence[RecordOutcome], existing: Mapping[str, set[str]]
) -> list[RecordOutcome]:
    """Withhold records the target already holds instead of appending them twice."""

    updated: list[RecordOutcome] = []
    for item in outcomes:
        known = existing.get(item.destination_key or "", set())
        if item.status == ACCEPTED and item.record_id in known:
            updated.append(
                replace(
                    item,
                    status=SKIPPED_EXISTING,
                    reasons=tuple(dict.fromkeys((*item.reasons, "already_in_target"))),
                )
            )
        else:
            updated.append(item)
    return updated


def _planned_outcome(target: DeliveryTarget, plan: DeliveryPlan) -> TargetOutcome:
    planned = next(item for item in plan.targets if item.destination_key == target.key)
    return TargetOutcome(
        target.key,
        planned.template_path,
        False,
        status="planned",
    )
