import csv

import pytest
from openpyxl import Workbook, load_workbook

from excel_ops.workbook_export import WorkbookExportError, export_csv, export_xlsx


def _workbook(path):
    workbook = Workbook()
    first = workbook.active
    first.title = "Payroll"
    first.append(["员工", "金额"])
    first.append(["林", 12.5])
    second = workbook.create_sheet("Summary")
    second.append(["总数"])
    second.append([1])
    workbook.save(path)


def test_xlsx_export_copies_a_real_workbook(tmp_path):
    source = tmp_path / "source.xlsx"
    destination = tmp_path / "out.xlsx"
    _workbook(source)
    assert export_xlsx(source, destination) == destination
    assert load_workbook(destination, read_only=True).sheetnames == ["Payroll", "Summary"]


def test_csv_single_sheet_is_utf8_and_preserves_values(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "payroll.csv"
    _workbook(source)
    outputs = export_csv(source, output, sheet="Payroll")
    assert outputs == (output,)
    with output.open(encoding="utf-8-sig", newline="") as stream:
        assert list(csv.reader(stream)) == [["员工", "金额"], ["林", "12.5"]]


def test_csv_can_export_each_sheet_to_a_directory(tmp_path):
    source = tmp_path / "source.xlsx"
    output_dir = tmp_path / "csv"
    _workbook(source)
    outputs = export_csv(source, output_dir, sheets=["Payroll", "Summary"])
    assert [path.name for path in outputs] == ["Payroll.csv", "Summary.csv"]


def test_csv_rejects_ambiguous_or_missing_sheets(tmp_path):
    source = tmp_path / "source.xlsx"
    _workbook(source)
    with pytest.raises(WorkbookExportError, match="choose sheet"):
        export_csv(source, tmp_path / "out.csv", sheet="Payroll", sheets=["Summary"])
    with pytest.raises(WorkbookExportError, match="unknown worksheet"):
        export_csv(source, tmp_path / "out.csv", sheet="Missing")


def test_csv_preserves_explicit_empty_selection_as_an_error(tmp_path):
    source = tmp_path / "source.xlsx"
    _workbook(source)
    with pytest.raises(WorkbookExportError, match="at least one sheet"):
        export_csv(source, tmp_path / "csv", sheets=[])


def test_csv_sanitizes_windows_sheet_names(tmp_path):
    source = tmp_path / "source.xlsx"
    workbook = Workbook()
    workbook.active.title = "CON"
    workbook.create_sheet("Summary")
    workbook.save(source)
    outputs = export_csv(source, tmp_path / "csv", sheets=["CON", "Summary"])
    assert [path.name for path in outputs] == ["_CON.csv", "Summary.csv"]


def test_csv_rejects_formula_without_cached_result(tmp_path):
    source = tmp_path / "source.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "=1+1"
    workbook.save(source)
    with pytest.raises(WorkbookExportError, match="no cached result"):
        export_csv(source, tmp_path / "out.csv", sheet="Sheet")


def test_xlsm_is_not_silently_renamed_to_xlsx(tmp_path):
    source = tmp_path / "source.xlsm"
    source.write_bytes(b"not a workbook")
    with pytest.raises(WorkbookExportError, match="XLSX source"):
        export_xlsx(source, tmp_path / "out.xlsx")
