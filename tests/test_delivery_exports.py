from pathlib import Path

from openpyxl import Workbook

from excel_ops.delivery_exports import export_delivery_artifacts


def _workbook(path: Path) -> None:
    """Create one visible and one hidden synthetic worksheet."""

    workbook = Workbook()
    workbook.active.title = "Visible"
    hidden = workbook.create_sheet("Hidden")
    hidden.sheet_state = "hidden"
    workbook.save(path)
    workbook.close()


def test_windows_pdf_all_uses_visible_sheets_and_stays_unverified(tmp_path, monkeypatch):
    """Native Excel receives no hidden sheet and existence is not layout proof."""

    source = tmp_path / "report.xlsx"
    _workbook(source)
    calls = []

    def fake_export(source_path, output_path, *, sheets):
        calls.append(tuple(sheets))
        Path(output_path).write_bytes(b"%PDF-synthetic")

    monkeypatch.setattr("excel_ops.delivery_exports._is_windows", lambda: True)
    monkeypatch.setattr("excel_ops.delivery_exports.export_pdf_with_excel", fake_export)

    artifact = export_delivery_artifacts(
        source,
        formats=("pdf",),
        selection={"pdf": None},
    )[0]

    assert calls == [("Visible",)]
    assert artifact.sheets == ("Visible",)
    assert artifact.verification == "unverified"
    assert "pdf_layout_not_verified" in artifact.warnings
