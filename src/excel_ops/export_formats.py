"""Validate the format-selection contract without performing an export."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass


SUPPORTED_EXPORT_FORMATS = ("xlsx", "csv", "pdf")


class ExportFormatError(ValueError):
    """Raised when a delivery format request is malformed or unsafe."""


@dataclass(frozen=True)
class ExportFormatSelection:
    formats: tuple[str, ...]
    csv_mode: str | None = None
    csv_sheet: str | None = None
    pdf_sheets: tuple[str, ...] | None = None

    def warnings(self) -> tuple[str, ...]:
        warnings: list[str] = []
        if "csv" in self.formats:
            warnings.append("csv_drops_workbook_features")
        if "pdf" in self.formats:
            warnings.append("pdf_is_not_editable")
        return tuple(warnings)

    def to_dict(self) -> dict[str, object]:
        return {
            "formats": list(self.formats),
            "csv": {"mode": self.csv_mode, "sheet": self.csv_sheet}
            if self.csv_mode
            else None,
            "pdf": {"sheets": list(self.pdf_sheets)}
            if self.pdf_sheets is not None
            else None,
            "warnings": list(self.warnings()),
        }


def parse_export_formats(payload: Mapping[str, object]) -> ExportFormatSelection:
    """Parse ``delivery.formats`` while preserving the legacy XLSX default."""

    delivery = payload.get("delivery", {})
    if delivery is None:
        delivery = {}
    if not isinstance(delivery, Mapping):
        raise ExportFormatError("delivery must be an object")
    raw_formats = delivery.get("formats", payload.get("formats", ("xlsx",)))
    if isinstance(raw_formats, str) or not isinstance(raw_formats, Sequence):
        raise ExportFormatError("delivery.formats must be an array")
    formats: list[str] = []
    for raw in raw_formats:
        if not isinstance(raw, str) or raw not in SUPPORTED_EXPORT_FORMATS:
            raise ExportFormatError(
                f"unsupported export format: {raw!r}; expected xlsx, csv, or pdf"
            )
        if raw not in formats:
            formats.append(raw)
    if not formats:
        raise ExportFormatError("delivery.formats must not be empty")

    raw_csv = delivery.get("csv", {})
    if raw_csv is None:
        raw_csv = {}
    if not isinstance(raw_csv, Mapping):
        raise ExportFormatError("delivery.csv must be an object")
    csv_mode = raw_csv.get("mode")
    csv_sheet = raw_csv.get("sheet")
    if csv_sheet is not None and not isinstance(csv_sheet, str):
        raise ExportFormatError("csv.sheet must be a string")
    if "csv" in formats:
        if csv_mode not in ("one-file-per-sheet", "single-sheet"):
            raise ExportFormatError(
                "CSV export requires csv.mode=single-sheet or one-file-per-sheet"
            )
        if csv_mode == "single-sheet" and (not isinstance(csv_sheet, str) or not csv_sheet.strip()):
            raise ExportFormatError("single-sheet CSV export requires csv.sheet")
    elif csv_mode is not None or csv_sheet is not None:
        raise ExportFormatError("CSV options require csv in delivery.formats")

    raw_pdf = delivery.get("pdf", {})
    if raw_pdf is None:
        raw_pdf = {}
    if not isinstance(raw_pdf, Mapping):
        raise ExportFormatError("delivery.pdf must be an object")
    pdf_sheets = raw_pdf.get("sheets")
    normalized_pdf: tuple[str, ...] | None = None
    if "pdf" in formats:
        if pdf_sheets == "all":
            normalized_pdf = None
        elif (
            isinstance(pdf_sheets, Sequence)
            and not isinstance(pdf_sheets, str)
            and len(pdf_sheets) > 0
        ):
            if not all(isinstance(item, str) and item.strip() for item in pdf_sheets):
                raise ExportFormatError("pdf.sheets must contain non-empty sheet names")
            normalized_pdf = tuple(pdf_sheets)
        else:
            raise ExportFormatError("PDF export requires pdf.sheets=all or a sheet array")
    elif pdf_sheets is not None:
        raise ExportFormatError("PDF options require pdf in delivery.formats")
    return ExportFormatSelection(
        tuple(formats),
        str(csv_mode) if csv_mode is not None else None,
        str(csv_sheet).strip() if isinstance(csv_sheet, str) else None,
        normalized_pdf,
    )
