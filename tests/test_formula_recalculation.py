import shutil
import zipfile
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook, load_workbook

from excel_ops.formula_recalculation import (
    FormulaEngineUnavailable,
    FormulaRecalculationError,
    FormulaValueExpectation,
    verify_formula_recalculation,
)
from excel_ops.formula_verification import FormulaRegion


def _source(path: Path, formula: str = "=SUM(A2:B2)") -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"
    sheet["A2"] = 2
    sheet["B2"] = 3
    sheet["C2"] = formula
    workbook.save(path)
    workbook.close()


def _calculated(value):
    def recalculate(source: Path, destination: Path) -> str:
        workbook = load_workbook(source, data_only=False)
        workbook["Report"]["C2"] = value
        workbook.save(destination)
        workbook.close()
        return "test-engine"

    return recalculate


def _replace_cached_value(path: Path, *, cell_reference: str, value: str) -> None:
    replacement = path.with_suffix(".replacement.xlsx")
    worksheet_path = "xl/worksheets/sheet1.xml"
    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        replacement, "w"
    ) as destination:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == worksheet_path:
                root = ET.fromstring(data)
                cell = root.find(f".//{namespace}c[@r='{cell_reference}']")
                assert cell is not None
                cached = cell.find(f"{namespace}v")
                assert cached is not None
                cached.text = value
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            destination.writestr(item, data)
    replacement.replace(path)


def test_recalculation_verifies_an_independent_expected_value(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "recalculated.xlsx"
    _source(source)

    result = verify_formula_recalculation(
        source,
        output,
        expectations=(FormulaValueExpectation("Report", "C2", 5),),
        formula_regions=(FormulaRegion("Report", "C2:C2"),),
        recalculator=_formula_preserving_calculated(5),
    )

    assert result.status == "verified"
    assert result.passed is True
    assert result.engine == "test-engine"
    assert result.recalculated_path == output.resolve()
    assert output.exists()
    workbook = load_workbook(output, data_only=False)
    try:
        assert workbook["Report"]["C2"].value == "=SUM(A2:B2)"
    finally:
        workbook.close()


def _formula_preserving_calculated(value):
    def recalculate(source: Path, destination: Path) -> str:
        shutil.copy2(source, destination)
        _replace_cached_value(destination, cell_reference="C2", value=str(value))
        return "test-engine"

    return recalculate


def test_recalculation_rejects_an_engine_that_removes_the_formula(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "recalculated.xlsx"
    _source(source)

    result = verify_formula_recalculation(
        source,
        output,
        expectations=(FormulaValueExpectation("Report", "C2", 5),),
        formula_regions=(FormulaRegion("Report", "C2:C2"),),
        recalculator=_calculated(5),
    )

    assert result.status == "failed"
    assert "recalculated_formula_not_preserved" in {
        item.code for item in result.findings
    }
    assert not output.exists()


def test_recalculation_uses_numeric_tolerance(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "recalculated.xlsx"
    _source(source)

    result = verify_formula_recalculation(
        source,
        output,
        expectations=(
            FormulaValueExpectation("Report", "C2", Decimal("5.00"), Decimal("0.01")),
        ),
        recalculator=_calculated(5.005),
    )

    assert result.status == "verified"


def test_mismatch_fails_without_publishing_output(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "recalculated.xlsx"
    _source(source)

    result = verify_formula_recalculation(
        source,
        output,
        expectations=(FormulaValueExpectation("Report", "C2", 6),),
        recalculator=_calculated(5),
    )

    assert result.status == "failed"
    assert {item.code for item in result.findings} == {"recalculation_mismatch"}
    assert not output.exists()


def test_recalculated_formula_error_is_blocking(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "recalculated.xlsx"
    _source(source)

    result = verify_formula_recalculation(
        source,
        output,
        expectations=(FormulaValueExpectation("Report", "C2", 5),),
        recalculator=_calculated("#N/A"),
    )

    assert result.status == "failed"
    assert {item.code for item in result.findings} == {
        "recalculated_formula_error"
    }


def test_recalculation_scans_errors_outside_expected_cells(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "recalculated.xlsx"
    _source(source)

    def recalculate(source: Path, destination: Path) -> str:
        workbook = load_workbook(source)
        workbook["Report"]["C2"] = 5
        error_cell = workbook["Report"]["D8"]
        error_cell.value = "#GETTING_DATA"
        error_cell.data_type = "e"
        workbook.save(destination)
        workbook.close()
        return "test-engine"

    result = verify_formula_recalculation(
        source,
        output,
        expectations=(FormulaValueExpectation("Report", "C2", 5),),
        recalculator=recalculate,
    )

    assert result.status == "failed"
    finding = next(
        item for item in result.findings if item.code == "recalculated_formula_error"
    )
    assert (finding.sheet, finding.cell) == ("Report", "D8")


def test_static_formula_failure_blocks_engine_execution(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "recalculated.xlsx"
    _source(source, "=#REF!")
    called = False

    def should_not_run(source: Path, destination: Path) -> str:
        nonlocal called
        called = True
        return "test-engine"

    result = verify_formula_recalculation(
        source,
        output,
        expectations=(FormulaValueExpectation("Report", "C2", 5),),
        recalculator=should_not_run,
    )

    assert result.status == "failed"
    assert "broken_reference" in {item.code for item in result.findings}
    assert called is False


@pytest.mark.parametrize("cell", ["C2:C3", "not-a-cell"])
def test_invalid_expectation_cell_is_rejected_before_recalculation(tmp_path, cell):
    source = tmp_path / "source.xlsx"
    _source(source)
    called = False

    def should_not_run(source: Path, destination: Path) -> str:
        nonlocal called
        called = True
        return "test-engine"

    with pytest.raises(
        FormulaRecalculationError,
        match="expectation cell must be a single cell reference",
    ):
        verify_formula_recalculation(
            source,
            tmp_path / "recalculated.xlsx",
            expectations=(FormulaValueExpectation("Report", cell, 5),),
            recalculator=should_not_run,
        )

    assert called is False


def test_missing_expectations_are_unverified(tmp_path):
    source = tmp_path / "source.xlsx"
    _source(source)

    result = verify_formula_recalculation(
        source,
        tmp_path / "recalculated.xlsx",
        expectations=(),
        recalculator=_calculated(5),
    )

    assert result.status == "unverified"
    assert [item.code for item in result.findings] == [
        "recalculation_expectations_missing"
    ]


def test_missing_engine_is_reported_as_unverified(tmp_path):
    source = tmp_path / "source.xlsx"
    _source(source)

    def unavailable(source: Path, destination: Path) -> str:
        raise FormulaEngineUnavailable("No calculation engine is installed")

    result = verify_formula_recalculation(
        source,
        tmp_path / "recalculated.xlsx",
        expectations=(FormulaValueExpectation("Report", "C2", 5),),
        recalculator=unavailable,
    )

    assert result.status == "unverified"
    assert [item.code for item in result.findings] == [
        "recalculation_not_verified"
    ]


def test_circular_reference_is_detected_before_recalculation(tmp_path):
    source = tmp_path / "source.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"
    sheet["A1"] = "=B1"
    sheet["B1"] = "=A1"
    workbook.save(source)
    workbook.close()

    result = verify_formula_recalculation(
        source,
        tmp_path / "recalculated.xlsx",
        expectations=(FormulaValueExpectation("Report", "A1", 1),),
        recalculator=_calculated(1),
    )

    assert result.status == "failed"
    assert "circular_reference" in {item.code for item in result.findings}


def test_libreoffice_replaces_a_stale_cached_formula_value(tmp_path):
    if shutil.which("soffice") is None:
        pytest.skip("LibreOffice soffice is not installed")

    source = tmp_path / "source.xlsx"
    output = tmp_path / "recalculated.xlsx"
    _source(source)
    _replace_cached_value(source, cell_reference="C2", value="999")

    result = verify_formula_recalculation(
        source,
        output,
        expectations=(FormulaValueExpectation("Report", "C2", 5),),
        engine="libreoffice",
    )

    assert result.status == "verified"
    workbook = load_workbook(output, data_only=True)
    try:
        assert workbook["Report"]["C2"].value == 5
    finally:
        workbook.close()
