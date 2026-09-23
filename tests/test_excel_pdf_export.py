from pathlib import Path
from subprocess import CompletedProcess

import os

import pytest
from openpyxl import Workbook

from excel_ops.excel_pdf_export import ExcelPdfExportError, export_pdf_with_excel


def _source(path: Path) -> None:
    """Create a two-sheet workbook for native Excel export tests."""

    workbook = Workbook()
    workbook.active.title = "Report"
    workbook.create_sheet("数据")
    workbook.save(path)
    workbook.close()


def test_native_excel_exports_selected_sheets_as_literal_arguments(tmp_path, monkeypatch):
    """Pass literal paths and Unicode sheet names without shell interpolation."""

    source = tmp_path / "source.xlsx"
    output = tmp_path / "report.pdf"
    _source(source)
    monkeypatch.setattr("excel_ops.excel_pdf_export._is_windows", lambda: True)

    def runner(command, **kwargs):
        """Inspect the request and emulate Excel writing a PDF."""

        assert command[4] == str(source.resolve())
        assert command[6:] == ["数据"]
        Path(command[5]).write_bytes(b"%PDF-1.7\nsynthetic")
        return CompletedProcess(command, 0, "", "")

    result = export_pdf_with_excel(
        source, output, sheets=["数据"], cscript="cscript.exe", runner=runner
    )
    assert result == output.resolve()
    assert output.read_bytes().startswith(b"%PDF-")


def test_native_excel_requires_windows(tmp_path, monkeypatch):
    """Fail before invoking PowerShell on a non-Windows host."""

    monkeypatch.setattr("excel_ops.excel_pdf_export._is_windows", lambda: False)
    with pytest.raises(ExcelPdfExportError, match="requires Windows"):
        export_pdf_with_excel(tmp_path / "source.xlsx", tmp_path / "out.pdf")


def test_native_excel_validates_sheet_scope(tmp_path, monkeypatch):
    """Reject empty, duplicate, and missing selections before Excel starts."""

    source = tmp_path / "source.xlsx"
    _source(source)
    monkeypatch.setattr("excel_ops.excel_pdf_export._is_windows", lambda: True)
    cases = [([], "at least one"), (["Report", "Report"], "duplicates"), (["Missing"], "unknown")]
    for sheets, message in cases:
        with pytest.raises(ExcelPdfExportError, match=message):
            export_pdf_with_excel(
                source, tmp_path / "out.pdf", sheets=sheets, cscript="cscript.exe"
            )


def test_native_excel_rejects_invalid_output(tmp_path, monkeypatch):
    """A successful COM process cannot publish non-PDF bytes."""

    source = tmp_path / "source.xlsx"
    _source(source)
    monkeypatch.setattr("excel_ops.excel_pdf_export._is_windows", lambda: True)

    def runner(command, **kwargs):
        """Emulate Excel producing an invalid artifact."""

        Path(command[5]).write_bytes(b"invalid")
        return CompletedProcess(command, 0, "", "")

    with pytest.raises(ExcelPdfExportError, match="not a PDF"):
        export_pdf_with_excel(
            source, tmp_path / "out.pdf", cscript="cscript.exe", runner=runner
        )


@pytest.mark.skipif(
    os.environ.get("EXCEL_OPS_TEST_NATIVE_EXCEL") != "1",
    reason="requires an interactive Windows host with Microsoft Excel installed",
)
def test_native_excel_real_export(tmp_path):
    """Exercise the installed Excel COM server when explicitly enabled."""

    source = tmp_path / "source.xlsx"
    output = tmp_path / "report.pdf"
    _source(source)
    export_pdf_with_excel(source, output, sheets=["数据"])
    assert output.read_bytes().startswith(b"%PDF-")
