"""Deterministic XLSX and CSV export helpers."""

from __future__ import annotations

import csv
import shutil
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook


class WorkbookExportError(ValueError):
    """Raised when an export request cannot be completed safely."""


def export_xlsx(source: str | Path, destination: str | Path) -> Path:
    """Copy an existing workbook to a new XLSX destination without changing it."""

    source_path = Path(source)
    destination_path = Path(destination)
    if source_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise WorkbookExportError("XLSX export requires an XLSX or XLSM source")
    if destination_path.suffix.lower() != ".xlsx":
        raise WorkbookExportError("XLSX export destination must end in .xlsx")
    if not source_path.is_file():
        raise WorkbookExportError(f"source workbook does not exist: {source_path}")
    destination_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_path, destination_path)
    return destination_path


def export_csv(
    source: str | Path,
    destination: str | Path,
    *,
    sheet: str | None = None,
    sheets: Iterable[str] | None = None,
) -> tuple[Path, ...]:
    """Export one workbook sheet to UTF-8 CSV, or one CSV file per sheet."""

    source_path = Path(source)
    destination_path = Path(destination)
    if source_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise WorkbookExportError("CSV export requires an XLSX or XLSM source")
    if destination_path.suffix.lower() != ".csv" and sheet is not None:
        raise WorkbookExportError("single-sheet CSV destination must end in .csv")
    if sheet is not None and sheets is not None:
        raise WorkbookExportError("choose sheet or sheets, not both")
    workbook = load_workbook(source_path, read_only=True, data_only=False)
    try:
        selected = [sheet] if sheet is not None else list(sheets or workbook.sheetnames)
        if not selected:
            raise WorkbookExportError("CSV export requires at least one sheet")
        missing = [name for name in selected if name not in workbook.sheetnames]
        if missing:
            raise WorkbookExportError(f"unknown worksheet(s): {', '.join(missing)}")
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        outputs: list[Path] = []
        for name in selected:
            output = destination_path if sheet is not None else destination_path / f"{name}.csv"
            if sheet is None and destination_path.suffix:
                raise WorkbookExportError("multi-sheet CSV destination must be a directory")
            if output.suffix.lower() != ".csv":
                raise WorkbookExportError("CSV output must end in .csv")
            output.parent.mkdir(parents=True, exist_ok=True)
            with output.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream, lineterminator="\n")
                for row in workbook[name].iter_rows(values_only=True):
                    writer.writerow(list(row))
            outputs.append(output)
        return tuple(outputs)
    finally:
        workbook.close()
