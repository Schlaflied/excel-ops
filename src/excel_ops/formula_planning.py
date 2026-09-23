"""Version-aware, reviewable formula planning for agent-interpreted goals."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from typing import Any, Literal


ExcelVersion = Literal["2016", "2019", "2021", "365"]
OutputMode = Literal["formula", "static"]
_A1 = re.compile(r"\$?[A-Z]{1,3}\$?[1-9][0-9]*")
_RANGE = re.compile(r"\$?[A-Z]{1,3}\$?[1-9][0-9]*:\$?[A-Z]{1,3}\$?[1-9][0-9]*")
_MISSING = object()


class FormulaPlanError(ValueError):
    """Raised when a formula request is incomplete or unsafe."""


@dataclass(frozen=True)
class LookupFormulaSpec:
    business_rule: str
    lookup_cell: str
    source_sheet: str
    lookup_range: str
    return_range: str


@dataclass(frozen=True)
class ConditionalSumFormulaSpec:
    business_rule: str
    criteria_cell: str
    source_sheet: str
    criteria_range: str
    sum_range: str


@dataclass(frozen=True)
class DateAddFormulaSpec:
    business_rule: str
    date_cell: str
    months: int


FormulaSpec = LookupFormulaSpec | ConditionalSumFormulaSpec | DateAddFormulaSpec


@dataclass(frozen=True)
class FormulaPlan:
    """A formula or approved static value plus compatibility evidence."""

    business_rule: str
    operation: str
    output_mode: OutputMode
    target_excel_version: ExcelVersion
    value: Any
    function: str | None
    compatibility_strategy: str
    requires_independent_recalculation: bool

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def plan_formula(
    spec: FormulaSpec,
    *,
    target_excel_version: ExcelVersion,
    output_mode: OutputMode = "formula",
    computed_value: Any = _MISSING,
) -> FormulaPlan:
    """Plan a supported formula without interpreting free-form text in Python.

    The calling agent translates the user's natural-language goal into a typed
    spec. Static mode is fail-closed: an independently computed value must be
    supplied explicitly rather than inferred from a formula string.
    """

    if target_excel_version not in {"2016", "2019", "2021", "365"}:
        raise FormulaPlanError("target_excel_version must be 2016, 2019, 2021, or 365")
    if output_mode not in {"formula", "static"}:
        raise FormulaPlanError("output_mode must be formula or static")
    if not spec.business_rule.strip():
        raise FormulaPlanError("business_rule must explain the intended result")
    if output_mode == "static":
        if computed_value is _MISSING:
            raise FormulaPlanError("static mode requires an independently computed value")
        return FormulaPlan(
            spec.business_rule.strip(),
            _operation(spec),
            output_mode,
            target_excel_version,
            computed_value,
            None,
            "static value supplied by an independent calculation path",
            False,
        )

    formula, function, strategy = _formula(spec, target_excel_version)
    return FormulaPlan(
        spec.business_rule.strip(),
        _operation(spec),
        output_mode,
        target_excel_version,
        formula,
        function,
        strategy,
        True,
    )


def _formula(spec: FormulaSpec, version: ExcelVersion) -> tuple[str, str, str]:
    if isinstance(spec, LookupFormulaSpec):
        lookup_cell = _cell(spec.lookup_cell, "lookup_cell")
        lookup_range = _range(spec.lookup_range, "lookup_range")
        return_range = _range(spec.return_range, "return_range")
        sheet = _sheet(spec.source_sheet)
        if version in {"2021", "365"}:
            return (
                f'=XLOOKUP({lookup_cell},{sheet}!{lookup_range},{sheet}!{return_range},"",0)',
                "XLOOKUP",
                "XLOOKUP is available in the selected Excel version",
            )
        return (
            f'=IFERROR(INDEX({sheet}!{return_range},MATCH({lookup_cell},{sheet}!{lookup_range},0)),"")',
            "INDEX/MATCH",
            "INDEX/MATCH fallback avoids XLOOKUP on legacy Excel",
        )
    if isinstance(spec, ConditionalSumFormulaSpec):
        criteria_cell = _cell(spec.criteria_cell, "criteria_cell")
        criteria_range = _range(spec.criteria_range, "criteria_range")
        sum_range = _range(spec.sum_range, "sum_range")
        sheet = _sheet(spec.source_sheet)
        return (
            f"=SUMIFS({sheet}!{sum_range},{sheet}!{criteria_range},{criteria_cell})",
            "SUMIFS",
            "SUMIFS is supported by all target versions",
        )
    if isinstance(spec, DateAddFormulaSpec):
        date_cell = _cell(spec.date_cell, "date_cell")
        if not isinstance(spec.months, int) or isinstance(spec.months, bool):
            raise FormulaPlanError("months must be an integer")
        return (
            f"=EDATE({date_cell},{spec.months})",
            "EDATE",
            "EDATE is supported by all target versions",
        )
    raise FormulaPlanError("unsupported formula specification")


def _operation(spec: FormulaSpec) -> str:
    if isinstance(spec, LookupFormulaSpec):
        return "lookup"
    if isinstance(spec, ConditionalSumFormulaSpec):
        return "conditional_sum"
    if isinstance(spec, DateAddFormulaSpec):
        return "date_add"
    raise FormulaPlanError("unsupported formula specification")


def _cell(value: str, field: str) -> str:
    normalized = value.strip().upper()
    if not _A1.fullmatch(normalized):
        raise FormulaPlanError(f"{field} must be one A1 cell reference")
    return normalized


def _range(value: str, field: str) -> str:
    normalized = value.strip().upper()
    if not _RANGE.fullmatch(normalized):
        raise FormulaPlanError(f"{field} must be one bounded A1 range")
    return normalized


def _sheet(value: str) -> str:
    normalized = value.strip()
    if not normalized or any(character in normalized for character in "[]:/*?\\"):
        raise FormulaPlanError("source_sheet contains an invalid Excel sheet name")
    return "'" + normalized.replace("'", "''") + "'"
