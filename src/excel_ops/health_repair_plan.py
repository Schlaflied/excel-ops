"""Deterministic, source-bound dry-run planning for workbook health findings."""

from __future__ import annotations

from collections.abc import Mapping as MappingABC
from hashlib import sha256
import json
from pathlib import Path
from typing import Any, Literal, Mapping

from .health_models import (
    HealthActionDecision,
    HealthFinding,
    HealthRepairAction,
    HealthRepairPlan,
    HealthRepairPlanError,
)
from .health_scan import scan_workbook_health

ActionDisposition = Literal["selected", "rejected"]
ActionBaseline = Mapping[str, Any]

_OPERATIONS = {
    "abnormal_merged_cells": "unmerge_cells",
    "blank_data_rows": "remove_blank_rows",
    "blank_zero_scan_budget_exceeded": "report_scan_budget",
    "blank_zero_semantics": "declare_blank_zero_policy",
    "content_beyond_working_area": "remove_outlying_content",
    "default_empty_sheet": "remove_empty_sheet",
    "duplicate_header": "rename_duplicate_header",
    "empty_workbook": "report_empty_workbook",
    "extra_whitespace": "normalize_whitespace",
    "filter_range_mismatch": "align_filter_range",
    "freeze_panes_mismatch": "align_freeze_panes",
    "inconsistent_cell_style": "declare_style_baseline",
    "invisible_characters": "remove_invisible_characters",
    "mixed_date_formats": "declare_date_format",
    "numeric_text": "convert_numeric_text",
    "print_area_mismatch": "align_print_area",
}


def _canonical(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _digest(value: Any) -> str:
    return sha256(_canonical(value).encode("utf-8")).hexdigest()


def _current_state(finding: HealthFinding) -> dict[str, Any]:
    summary = dict(finding.details) or {
        "code": finding.code,
        "location": finding.location,
        "evidence_digest": sha256(finding.evidence.encode("utf-8")).hexdigest(),
    }
    return {"summary": summary, "digest": _digest(summary)}


def _action_id(
    source_sha256: str,
    finding: HealthFinding,
    operation: str,
    current_state: Mapping[str, Any],
) -> str:
    identity = {
        "source_sha256": source_sha256,
        "finding_id": finding.finding_id,
        "operation": operation,
        "current_state_digest": current_state["digest"],
    }
    return f"health-{sha256(_canonical(identity).encode('utf-8')).hexdigest()[:20]}"


def _validate_style_baseline(finding: HealthFinding, baseline: ActionBaseline) -> dict[str, str]:
    digest = baseline.get("target_signature_digest")
    available = {
        item["signature_digest"] for item in finding.details.get("signatures", [])
    }
    if not isinstance(digest, str) or digest not in available:
        raise HealthRepairPlanError("style baseline must name a current target_signature_digest")
    return {"target_signature_digest": digest}


def _validate_date_baseline(baseline: ActionBaseline) -> dict[str, str]:
    number_format = baseline.get("target_number_format")
    locale = baseline.get("locale")
    if not isinstance(number_format, str) or not number_format.strip():
        raise HealthRepairPlanError("date baseline requires target_number_format")
    if not isinstance(locale, str) or not locale.strip():
        raise HealthRepairPlanError("date baseline requires locale")
    return {"target_number_format": number_format, "locale": locale}


def _validate_blank_zero_baseline(baseline: ActionBaseline) -> dict[str, str]:
    blank_action = baseline.get("blank_action")
    zero_action = baseline.get("zero_action")
    if blank_action not in {"preserve", "convert_to_zero"}:
        raise HealthRepairPlanError("blank_action must be preserve or convert_to_zero")
    if zero_action not in {"preserve", "convert_to_blank"}:
        raise HealthRepairPlanError("zero_action must be preserve or convert_to_blank")
    return {"blank_action": blank_action, "zero_action": zero_action}


def _normalize_header(value: str) -> str:
    return " ".join(value.split()).casefold()


def _validate_header_baseline(
    finding: HealthFinding, baseline: ActionBaseline
) -> dict[str, str]:
    target = baseline.get("target_header")
    if not isinstance(target, str) or not target.strip():
        raise HealthRepairPlanError("header baseline requires target_header")
    normalized = _normalize_header(target)
    if normalized == finding.details.get("current_normalized"):
        raise HealthRepairPlanError("target_header must differ from the current duplicate header")
    occupied = {
        item["normalized"]
        for item in finding.details.get("headers", [])
        if item.get("coordinate") != finding.location
    }
    if normalized in occupied:
        raise HealthRepairPlanError("target_header conflicts with another header in the same row")
    return {"target_header": " ".join(target.split()), "normalized": normalized}


def _validate_numeric_baseline(
    finding: HealthFinding, baseline: ActionBaseline
) -> dict[str, str]:
    if baseline.get("target_representation") != "preserve_text":
        raise HealthRepairPlanError("unsafe numeric text must use target_representation=preserve_text")
    return {
        "target_representation": "preserve_text",
        "value": str(finding.details["current_value"]),
    }


def _validate_baseline(
    finding: HealthFinding, operation: str, baseline: ActionBaseline
) -> dict[str, Any]:
    if operation == "declare_style_baseline":
        return _validate_style_baseline(finding, baseline)
    if operation == "declare_date_format":
        return _validate_date_baseline(baseline)
    if operation == "declare_blank_zero_policy":
        return _validate_blank_zero_baseline(baseline)
    if operation == "rename_duplicate_header":
        return _validate_header_baseline(finding, baseline)
    if operation == "convert_numeric_text":
        return _validate_numeric_baseline(finding, baseline)
    raise HealthRepairPlanError(f"operation does not accept a baseline: {operation}")


def _candidate_target(finding: HealthFinding) -> dict[str, Any] | None:
    for key in (
        "candidate_value",
        "candidate_ref",
        "candidate_coordinate",
        "candidate_refs",
        "candidate_operation",
    ):
        if key in finding.details:
            return {key: finding.details[key]}
    return None


def _proposed_state(
    finding: HealthFinding,
    operation: str,
    baseline: ActionBaseline | None,
) -> dict[str, Any]:
    if finding.fixability == "unsupported":
        return {"status": "report_only", "target": None}
    if finding.details.get("required_baseline"):
        if baseline is None:
            return {
                "status": "requires_declared_baseline",
                "required": finding.details["required_baseline"],
                "target": None,
            }
        return {"status": "ready", "target": _validate_baseline(finding, operation, baseline)}
    if baseline is not None:
        raise HealthRepairPlanError(f"operation does not accept a baseline: {operation}")
    candidate = _candidate_target(finding)
    if candidate is not None:
        return {"status": "ready", "target": candidate}
    return {"status": "requires_declared_baseline", "required": "explicit_target", "target": None}


def _decision(
    action_id: str,
    classification: str,
    proposed_state: Mapping[str, Any],
    dispositions: Mapping[str, ActionDisposition],
) -> HealthActionDecision:
    decision = dispositions.get(action_id, "proposed")
    if decision == "selected" and classification == "unsupported":
        raise HealthRepairPlanError(f"unsupported action cannot be selected: {action_id}")
    if decision == "selected" and proposed_state["status"] != "ready":
        raise HealthRepairPlanError(f"action requires a validated baseline: {action_id}")
    return decision


def _build_action(
    source_sha256: str,
    finding: HealthFinding,
    dispositions: Mapping[str, ActionDisposition],
    baselines: Mapping[str, ActionBaseline],
) -> HealthRepairAction:
    operation = _OPERATIONS.get(finding.code, "report_only")
    current = _current_state(finding)
    action_id = _action_id(source_sha256, finding, operation, current)
    proposed = _proposed_state(finding, operation, baselines.get(action_id))
    return HealthRepairAction(
        action_id=action_id,
        finding_id=finding.finding_id,
        finding_code=finding.code,
        classification=finding.fixability,
        decision=_decision(action_id, finding.fixability, proposed, dispositions),
        operation=operation,
        sheet=finding.sheet,
        location=finding.location,
        impact_scope=f"{finding.sheet}!{finding.location}",
        current_state=current,
        proposed_state=proposed,
    )


def _validate_selected_header_targets(actions: tuple[HealthRepairAction, ...]) -> None:
    targets: dict[tuple[str, int, str], str] = {}
    for action in actions:
        if action.decision != "selected" or action.operation != "rename_duplicate_header":
            continue
        summary = action.current_state["summary"]
        target = action.proposed_state["target"]
        key = (action.sheet, int(summary["header_row"]), str(target["normalized"]))
        conflicting = targets.get(key)
        if conflicting is not None:
            raise HealthRepairPlanError(
                "selected header targets conflict within the same sheet and header row: "
                f"{conflicting}, {action.action_id}"
            )
        targets[key] = action.action_id


def plan_workbook_health_repairs(
    source: str | Path,
    *,
    dispositions: Mapping[str, ActionDisposition] | None = None,
    baselines: Mapping[str, ActionBaseline] | None = None,
) -> HealthRepairPlan:
    """Create a source-bound review plan without writing or authorizing changes."""

    requested = dict(dispositions or {})
    declared = dict(baselines or {})
    invalid = sorted(value for value in requested.values() if value not in {"selected", "rejected"})
    if invalid:
        raise HealthRepairPlanError(f"invalid action disposition(s): {invalid!r}")
    invalid_baselines = sorted(key for key, value in declared.items() if not isinstance(value, MappingABC))
    if invalid_baselines:
        raise HealthRepairPlanError(f"invalid action baseline(s): {invalid_baselines!r}")
    scan = scan_workbook_health(source)
    actions = tuple(
        _build_action(scan.source_sha256, finding, requested, declared)
        for finding in scan.findings
    )
    known = {action.action_id for action in actions}
    unknown = sorted((set(requested) | set(declared)) - known)
    if unknown:
        raise HealthRepairPlanError(f"unknown or stale action id(s): {unknown!r}")
    _validate_selected_header_targets(actions)
    return HealthRepairPlan(source=scan.source, source_sha256=scan.source_sha256, actions=actions)
