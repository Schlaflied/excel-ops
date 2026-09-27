"""Declarative JSON configuration parsing for delivery targets."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from .delivery_models import DeliveryPlanError, DeliveryTarget
from .delivery_verification import PeriodExpectation
from .export_formats import ExportFormatError, parse_export_formats
from .formula_delivery import FormulaDeliveryError, formula_rule_from_mapping
from .matching import Destination
from .number_formats import FormatPolicy, NumberFormatPolicyError
from .template_writer import TemplateMapping


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
