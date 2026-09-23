from pathlib import Path
import io
import subprocess

import os

import pytest
from openpyxl import Workbook

from excel_ops.excel_pdf_export import ExcelPdfExportError, export_pdf_with_excel


def _source(path: Path) -> None:
    """Create a two-sheet workbook for native Excel export tests."""

    workbook = Workbook()
    workbook.active.title = "Report"
    workbook.create_sheet("数据")
    workbook.create_sheet("Helper").sheet_state = "hidden"
    workbook.save(path)
    workbook.close()


def test_native_excel_exports_selected_sheets_as_literal_arguments(tmp_path, monkeypatch):
    """Pass literal paths and Unicode sheet names without shell interpolation."""

    source = tmp_path / "source.xlsx"
    output = tmp_path / "report.pdf"
    _source(source)
    monkeypatch.setattr("excel_ops.excel_pdf_export._is_windows", lambda: True)

    class Process:
        """Inspect the request and emulate Excel writing a PDF."""

        def __init__(self, command, **kwargs):
            assert command[4] == str(source.resolve())
            assert command[6:] == ["Report", "数据"]
            self.command = command
            self.stdout = io.StringIO("123\n")
            self.returncode = None

        def communicate(self, timeout=None):
            Path(self.command[5]).write_bytes(b"%PDF-1.7\nsynthetic")
            self.returncode = 0
            return "", ""

    result = export_pdf_with_excel(
        source, output, cscript="cscript.exe", popen=Process
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
    cases = [
        ([], "at least one"),
        (["Report", "Report"], "duplicates"),
        (["Missing"], "unknown"),
        (["Helper"], "hidden worksheet"),
    ]
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

    class Process:
        """Emulate Excel producing an invalid artifact."""

        def __init__(self, command, **kwargs):
            self.command = command
            self.stdout = io.StringIO("123\n")
            self.returncode = None

        def communicate(self, timeout=None):
            Path(self.command[5]).write_bytes(b"invalid")
            self.returncode = 0
            return "", ""

    with pytest.raises(ExcelPdfExportError, match="not a PDF"):
        export_pdf_with_excel(
            source, tmp_path / "out.pdf", cscript="cscript.exe", popen=Process
        )


def test_native_excel_timeout_terminates_its_window_process(tmp_path, monkeypatch):
    """A timeout kills cscript and the Excel process identified by its window."""

    source = tmp_path / "source.xlsx"
    _source(source)
    monkeypatch.setattr("excel_ops.excel_pdf_export._is_windows", lambda: True)
    terminated = []
    monkeypatch.setattr(
        "excel_ops.excel_pdf_export._terminate_excel_window",
        lambda window: terminated.append(window) or True,
    )

    class Process:
        """Emit an Excel handle, then simulate a blocked export."""

        def __init__(self, command, **kwargs):
            self.command = command
            self.stdout = io.StringIO("456\n")
            self.returncode = None
            self.killed = False

        def communicate(self, timeout=None):
            if timeout is not None:
                raise subprocess.TimeoutExpired(self.command, timeout)
            self.returncode = -9
            return "", ""

        def kill(self):
            self.killed = True

    with pytest.raises(ExcelPdfExportError, match="timed out; Excel process terminated"):
        export_pdf_with_excel(
            source, tmp_path / "out.pdf", sheets=["Report"], cscript="cscript.exe", popen=Process
        )
    assert terminated == [456]


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
