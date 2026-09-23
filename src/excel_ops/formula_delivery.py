"""Formula application and recalculation as one delivery-stage contract."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Literal, Mapping, Sequence

from .formula_application import FormulaApplicationError, apply_formula_plan
from .formula_planning import FormulaPlan
from .formula_recalculation import (
    FormulaRecalculationError,
    FormulaValueExpectation,
    Recalculator,
    verify_formula_recalculation,
)
from .formula_verification import FormulaRegion


class FormulaDeliveryError(ValueError):
    """Raised when a declared formula step cannot be delivered safely."""


@dataclass(frozen=True)
class FormulaDeliveryRule:
    """One explicit formula plan, destination range, and verification contract."""

    plan: FormulaPlan
    sheet: str
    target_range: str
    expectations: tuple[FormulaValueExpectation, ...] = ()
    overwrite_formulas: bool = False
    engine: Literal["auto", "excel", "libreoffice"] = "auto"

    def fingerprint_payload(self) -> dict[str, Any]:
        """Return complete rule identity for the in-memory idempotency hash."""

        return {
            "plan": self.plan.to_dict(),
            "sheet": self.sheet,
            "target_range": self.target_range,
            "expectations": [
                {
                    "sheet": item.sheet,
                    "cell": item.cell,
                    "expected": item.expected,
                    "tolerance": str(item.tolerance),
                }
                for item in sorted(
                    self.expectations,
                    key=lambda item: (item.sheet, item.cell, str(item.tolerance)),
                )
            ],
            "overwrite_formulas": self.overwrite_formulas,
            "engine": self.engine,
        }

    def manifest_payload(self) -> dict[str, Any]:
        """Return auditable formula policy without copying expected cell values."""

        return {
            "business_rule": self.plan.business_rule,
            "operation": self.plan.operation,
            "output_mode": self.plan.output_mode,
            "target_excel_version": self.plan.target_excel_version,
            "function": self.plan.function,
            "compatibility_strategy": self.plan.compatibility_strategy,
            "sheet": self.sheet,
            "target_range": self.target_range,
            "overwrite_formulas": self.overwrite_formulas,
            "expectation_cells": sorted(
                f"{item.sheet}!{item.cell}" for item in self.expectations
            ),
        }


def formula_contract_fingerprint(rules: Sequence[FormulaDeliveryRule]) -> str:
    """Hash the complete formula contract without persisting its raw values."""

    payload = [rule.fingerprint_payload() for rule in rules]
    canonical = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        default=str,
    )
    return f"sha256:{hashlib.sha256(canonical.encode('utf-8')).hexdigest()}"


def formula_rule_from_mapping(payload: Mapping[str, Any]) -> FormulaDeliveryRule:
    """Parse one Agent-supplied rule and reject incomplete formula evidence."""

    raw_plan = payload.get("plan")
    if not isinstance(raw_plan, Mapping):
        raise FormulaDeliveryError("formula rule needs a plan object")
    requires_recalculation = raw_plan.get("requires_independent_recalculation")
    if not isinstance(requires_recalculation, bool):
        raise FormulaDeliveryError(
            "requires_independent_recalculation must be a boolean"
        )
    try:
        plan = FormulaPlan(
            business_rule=str(raw_plan["business_rule"]).strip(),
            operation=str(raw_plan["operation"]).strip(),
            output_mode=str(raw_plan["output_mode"]),  # type: ignore[arg-type]
            target_excel_version=str(raw_plan["target_excel_version"]),  # type: ignore[arg-type]
            value=raw_plan.get("value"),
            function=raw_plan.get("function"),
            compatibility_strategy=str(raw_plan["compatibility_strategy"]).strip(),
            requires_independent_recalculation=requires_recalculation,
        )
    except (KeyError, TypeError, ValueError) as error:
        raise FormulaDeliveryError("formula plan is incomplete") from error
    _validate_plan(plan)

    expectations: list[FormulaValueExpectation] = []
    for raw in payload.get("expectations", ()):
        if not isinstance(raw, Mapping):
            raise FormulaDeliveryError("formula expectations must be objects")
        try:
            tolerance = Decimal(str(raw.get("tolerance", "0")))
            expectations.append(
                FormulaValueExpectation(
                    str(raw["sheet"]),
                    str(raw["cell"]),
                    raw["expected"],
                    tolerance,
                )
            )
        except (InvalidOperation, KeyError, TypeError, ValueError) as error:
            raise FormulaDeliveryError("formula expectation is incomplete") from error
    if plan.requires_independent_recalculation and not expectations:
        raise FormulaDeliveryError(
            "formula mode requires independently derived expectations"
        )
    if not plan.requires_independent_recalculation and expectations:
        raise FormulaDeliveryError(
            "static mode expectations are not verified; remove them"
        )
    engine = str(payload.get("engine", "auto"))
    if engine not in {"auto", "excel", "libreoffice"}:
        raise FormulaDeliveryError("formula engine must be auto, excel, or libreoffice")
    sheet = str(payload.get("sheet") or "").strip()
    target_range = str(payload.get("target_range") or "").strip()
    if not sheet or not target_range:
        raise FormulaDeliveryError("formula rule needs sheet and target_range")
    overwrite_formulas = payload.get("overwrite_formulas", False)
    if not isinstance(overwrite_formulas, bool):
        raise FormulaDeliveryError("overwrite_formulas must be a boolean")
    return FormulaDeliveryRule(
        plan,
        sheet,
        target_range,
        tuple(expectations),
        overwrite_formulas,
        engine,  # type: ignore[arg-type]
    )


def apply_formula_delivery(
    workbook_path: str | Path,
    rules: Sequence[FormulaDeliveryRule],
    *,
    recalculator: Recalculator | None = None,
) -> dict[str, Any]:
    """Apply declared rules in place, then independently recalculate once."""

    workbook = Path(workbook_path).resolve()
    application_evidence: list[dict[str, Any]] = []
    for rule in rules:
        output = _temporary_workbook(workbook)
        try:
            result = apply_formula_plan(
                workbook,
                output,
                rule.plan,
                sheet=rule.sheet,
                target_range=rule.target_range,
                confirmed=True,
                overwrite_formulas=rule.overwrite_formulas,
            )
            if result.skipped:
                reasons = ", ".join(sorted({item.reason for item in result.skipped}))
                raise FormulaDeliveryError(
                    f"formula application skipped {len(result.skipped)} cell(s): {reasons}"
                )
            os.replace(output, workbook)
            application_evidence.append(
                {
                    **rule.manifest_payload(),
                    "written_cells": len(result.changes),
                    "status": "applied",
                }
            )
        except FormulaApplicationError as error:
            raise FormulaDeliveryError(str(error)) from error
        finally:
            output.unlink(missing_ok=True)

    formula_rules = tuple(
        rule for rule in rules if rule.plan.requires_independent_recalculation
    )
    if not formula_rules:
        return {
            "status": "verified",
            "engine": None,
            "contract_fingerprint": formula_contract_fingerprint(rules),
            "rules": application_evidence,
            "finding_codes": [],
        }
    engines = {rule.engine for rule in formula_rules}
    if len(engines) != 1:
        raise FormulaDeliveryError("all formula rules for one target must use one engine")
    recalculated = _temporary_workbook(workbook)
    try:
        result = verify_formula_recalculation(
            workbook,
            recalculated,
            expectations=tuple(
                item for rule in formula_rules for item in rule.expectations
            ),
            formula_regions=tuple(
                FormulaRegion(
                    rule.sheet,
                    rule.target_range
                    if ":" in rule.target_range
                    else f"{rule.target_range}:{rule.target_range}",
                )
                for rule in formula_rules
            ),
            engine=next(iter(engines)),  # type: ignore[arg-type]
            recalculator=recalculator,
        )
        if not result.passed or result.recalculated_path is None:
            codes = ", ".join(item.code for item in result.findings) or result.status
            raise FormulaDeliveryError(f"formula recalculation was not verified: {codes}")
        os.replace(result.recalculated_path, workbook)
        return {
            "status": result.status,
            "engine": result.engine,
            "contract_fingerprint": formula_contract_fingerprint(rules),
            "rules": application_evidence,
            "finding_codes": sorted({item.code for item in result.findings}),
        }
    except FormulaRecalculationError as error:
        raise FormulaDeliveryError(str(error)) from error
    finally:
        recalculated.unlink(missing_ok=True)


def _validate_plan(plan: FormulaPlan) -> None:
    if not plan.business_rule or not plan.operation or not plan.compatibility_strategy:
        raise FormulaDeliveryError("formula plan text fields must not be empty")
    if plan.output_mode not in {"formula", "static"}:
        raise FormulaDeliveryError("formula output_mode must be formula or static")
    if plan.target_excel_version not in {"2016", "2019", "2021", "365"}:
        raise FormulaDeliveryError("unsupported target Excel version")
    if plan.output_mode == "formula":
        if not isinstance(plan.value, str) or not plan.value.startswith("="):
            raise FormulaDeliveryError("formula mode requires formula text")
        if not plan.requires_independent_recalculation:
            raise FormulaDeliveryError("formula mode must require independent recalculation")
    elif plan.requires_independent_recalculation:
        raise FormulaDeliveryError("static mode must not request formula recalculation")


def _temporary_workbook(workbook: Path) -> Path:
    descriptor, name = tempfile.mkstemp(
        prefix=f".{workbook.stem}-formula-",
        suffix=workbook.suffix,
        dir=workbook.parent,
    )
    os.close(descriptor)
    path = Path(name)
    path.unlink()
    return path
