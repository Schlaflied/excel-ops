from pathlib import Path
from subprocess import CompletedProcess

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference

from excel_ops.pdf_export import PdfExportError, export_pdf


def _source(path: Path) -> None:
    """Create a workbook with a selected report and an external formula dependency."""

    workbook = Workbook()
    workbook.active.title = "Report"
    workbook.active["A1"] = "=Archive!A1"
    archive = workbook.create_sheet("Archive")
    archive["A1"] = 42
    chart = BarChart()
    chart.add_data(Reference(archive, min_col=1, min_row=1))
    workbook.create_chartsheet("Chart").add_chart(chart)
    workbook.save(path)
    workbook.close()


def _renderer(command, **kwargs):
    """Verify renderer isolation and the prepared workbook before writing a fake PDF."""

    outdir = Path(command[command.index("--outdir") + 1])
    prepared = Path(command[-1])
    profile_argument = next(item for item in command if item.startswith("-env:UserInstallation="))
    profile_uri = profile_argument.split("=", 1)[1]
    assert profile_uri == (outdir / "profile").as_uri()
    workbook = load_workbook(prepared, read_only=True)
    try:
        assert workbook.sheetnames == ["Report", "Archive"]
        assert workbook["Report"].sheet_state == "visible"
        assert workbook["Report"]["A1"].value == "=Archive!A1"
        assert workbook["Archive"].sheet_state == "hidden"
        assert workbook.active.title == "Report"
    finally:
        workbook.close()
    (outdir / "prepared.pdf").write_bytes(b"%PDF-1.7\nsynthetic")
    return CompletedProcess(command, 0, "", "")


def test_pdf_export_uses_selected_sheets_and_verifies_signature(tmp_path):
    """Selected sheets stay visible while formula dependencies remain intact."""

    source = tmp_path / "source.xlsx"
    output = tmp_path / "report.pdf"
    _source(source)
    assert export_pdf(source, output, sheets=["Report"], soffice="soffice", runner=_renderer) == output
    assert output.read_bytes().startswith(b"%PDF-")


def test_pdf_export_requires_renderer_and_explicit_valid_scope(tmp_path, monkeypatch):
    """Unavailable renderers and invalid sheet scopes fail before conversion."""

    source = tmp_path / "source.xlsx"
    _source(source)
    monkeypatch.setattr("excel_ops.pdf_export.shutil.which", lambda _: None)
    with pytest.raises(PdfExportError, match="renderer unavailable"):
        export_pdf(source, tmp_path / "out.pdf", soffice=None)
    with pytest.raises(PdfExportError, match="at least one sheet"):
        export_pdf(source, tmp_path / "out.pdf", sheets=[], soffice="soffice")
    with pytest.raises(PdfExportError, match="unknown worksheet"):
        export_pdf(source, tmp_path / "out.pdf", sheets=["Missing"], soffice="soffice")
    with pytest.raises(PdfExportError, match="contains duplicates"):
        export_pdf(source, tmp_path / "out.pdf", sheets=["Report", "Report"], soffice="soffice")


def test_pdf_export_rejects_non_pdf_renderer_output(tmp_path):
    """A successful process cannot publish output without a PDF signature."""

    source = tmp_path / "source.xlsx"
    _source(source)

    def invalid_renderer(command, **kwargs):
        """Write non-PDF bytes to exercise signature validation."""

        outdir = Path(command[command.index("--outdir") + 1])
        (outdir / "prepared.pdf").write_bytes(b"not-a-pdf")
        return CompletedProcess(command, 0, "", "")

    with pytest.raises(PdfExportError, match="not a PDF"):
        export_pdf(source, tmp_path / "out.pdf", soffice="soffice", runner=invalid_renderer)


def test_pdf_export_reports_renderer_start_failure(tmp_path):
    """Renderer startup failures are normalized as PDF export errors."""

    source = tmp_path / "source.xlsx"
    _source(source)

    def missing_renderer(command, **kwargs):
        """Simulate a renderer executable disappearing before invocation."""

        raise FileNotFoundError("soffice")

    with pytest.raises(PdfExportError, match="renderer could not run"):
        export_pdf(source, tmp_path / "out.pdf", soffice="soffice", runner=missing_renderer)


def test_pdf_export_keeps_a_selected_chartsheet_visible(tmp_path):
    """Selected chart sheets remain renderable while worksheets stay as dependencies."""

    source = tmp_path / "source.xlsx"
    _source(source)

    def chart_renderer(command, **kwargs):
        """Verify chart selection and publish a fake PDF."""

        outdir = Path(command[command.index("--outdir") + 1])
        workbook = load_workbook(Path(command[-1]), read_only=True)
        try:
            assert workbook.sheetnames == ["Report", "Archive", "Chart"]
            assert workbook["Report"].sheet_state == "hidden"
            assert workbook["Archive"].sheet_state == "hidden"
            assert workbook["Chart"].sheet_state == "visible"
            assert workbook.active.title == "Chart"
        finally:
            workbook.close()
        (outdir / "prepared.pdf").write_bytes(b"%PDF-1.7\nsynthetic")
        return CompletedProcess(command, 0, "", "")

    export_pdf(source, tmp_path / "chart.pdf", sheets=["Chart"], soffice="soffice", runner=chart_renderer)
