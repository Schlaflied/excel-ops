from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_ops.delivery_verification import DeliveryContract, DeliveryVerificationError, StageCounts, verify_and_deliver
from excel_ops.formula_verification import ExpectedErrorMarker, FormulaRegion, FormulaVerifier, SummaryReconciliation
from excel_ops.template_writer import TemplateMapping, write_template


def _written(tmp_path: Path, mutate):
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["id", "amount", "calculated"])
    sheet.append(["row-1", 10, "=B2*2"])
    sheet.append(["row-2", 20, "=B3*2"])
    sheet.append(["row-3", 30, "=B4*2"])
    summary = workbook.create_sheet("Summary")
    summary["A1"] = "Total"
    summary["B1"] = 60
    mutate(workbook)
    workbook.save(template)
    mapping = TemplateMapping(
        sheet="Data",
        field_columns={"id": "A", "amount": "B"},
        header_row=1,
        data_start_row=2,
    )
    return write_template(template, [], mapping, output_path=tmp_path / "staged.xlsx")


def _contract(verifier: FormulaVerifier) -> DeliveryContract:
    return DeliveryContract(StageCounts(0, 0, 0, 0), (), verifiers=(verifier,))


def _codes(error: pytest.ExceptionInfo[DeliveryVerificationError]) -> set[str]:
    return {finding.code for finding in error.value.result.findings}


def test_formula_errors_broken_sheet_external_refs_and_cycles_block_delivery(tmp_path: Path):
    def mutate(workbook):
        data = workbook["Data"]
        data["C2"] = "='Missing Sheet'!A1"
        data["C3"] = "='[old.xlsx]Sheet1'!A1"
        data["D2"] = "=D3"
        data["D3"] = "=D2"
        data["E2"] = "#REF!"

    result = _written(tmp_path, mutate)
    with pytest.raises(DeliveryVerificationError) as error:
        verify_and_deliver(result, tmp_path / "delivery.xlsx", _contract(FormulaVerifier()))
    assert _codes(error) >= {"missing_formula_sheet", "external_reference", "circular_reference", "formula_error_literal"}


@pytest.mark.parametrize(
    ("replacement", "expected"),
    [(None, "formula_fill_gap"), (99, "formula_overwritten_by_constant"), ("=B4*3", "formula_pattern_mismatch")],
)
def test_declared_formula_region_detects_fill_breaks(tmp_path: Path, replacement, expected: str):
    result = _written(tmp_path, lambda workbook: setattr(workbook["Data"]["C4"], "value", replacement))
    verifier = FormulaVerifier(formula_regions=(FormulaRegion("Data", "C2:C4", "C2"),), report_recalculation_limit=False)
    with pytest.raises(DeliveryVerificationError) as error:
        verify_and_deliver(result, tmp_path / "delivery.xlsx", _contract(verifier))
    assert expected in _codes(error)


def test_expected_missing_marker_is_not_treated_as_formula_failure(tmp_path: Path):
    result = _written(tmp_path, lambda workbook: setattr(workbook["Data"]["D2"], "value", "#N/A"))
    verifier = FormulaVerifier(expected_error_markers=(ExpectedErrorMarker("Data", "D2:D2"),))
    verification = verify_and_deliver(result, tmp_path / "delivery.xlsx", _contract(verifier))
    assert verification.passed
    assert [item.code for item in verification.findings] == ["recalculation_not_verified"]
    assert verification.findings[0].severity == "warning"


def test_summary_is_reconciled_from_independent_detail_values(tmp_path: Path):
    result = _written(tmp_path, lambda workbook: setattr(workbook["Summary"]["B1"], "value", 59))
    verifier = FormulaVerifier(
        reconciliations=(SummaryReconciliation("Data", "B2:B4", "Summary", "B1"),),
        report_recalculation_limit=False,
    )
    with pytest.raises(DeliveryVerificationError) as error:
        verify_and_deliver(result, tmp_path / "delivery.xlsx", _contract(verifier))
    assert "summary_reconciliation_mismatch" in _codes(error)


def test_formula_reconciliation_reports_engine_limit_without_claiming_recalculation(tmp_path: Path):
    result = _written(tmp_path, lambda workbook: setattr(workbook["Summary"]["B1"], "value", "=SUM(Data!B2:B4)"))
    verifier = FormulaVerifier(
        reconciliations=(SummaryReconciliation("Data", "B2:B4", "Summary", "B1"),),
    )
    verification = verify_and_deliver(result, tmp_path / "delivery.xlsx", _contract(verifier))
    assert verification.passed
    assert {item.code for item in verification.findings} == {"reconciliation_not_computable", "recalculation_not_verified"}
    assert all(item.severity == "warning" for item in verification.findings)


def test_dynamic_array_formula_is_reported(tmp_path: Path):
    result = _written(tmp_path, lambda workbook: setattr(workbook["Data"]["D2"], "value", "=FILTER(A2:A4,B2:B4>10)"))
    with pytest.raises(DeliveryVerificationError) as error:
        verify_and_deliver(result, tmp_path / "delivery.xlsx", _contract(FormulaVerifier()))
    assert "unsupported_dynamic_array" in _codes(error)
