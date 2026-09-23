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
    if source_path.suffix.lower() != ".xlsx":
        raise WorkbookExportError("XLSX export requires an XLSX source")
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
    if source_path.suffix.lower() != ".xlsx":
        raise WorkbookExportError("CSV export requires an XLSX source")
    if destination_path.suffix.lower() != ".csv" and sheet is not None:
        raise WorkbookExportError("single-sheet CSV destination must end in .csv")
    if sheet is not None and sheets is not None:
        raise WorkbookExportError("choose sheet or sheets, not both")
    workbook = load_workbook(source_path, read_only=True, data_only=False)
    values_workbook = load_workbook(source_path, read_only=True, data_only=True)
    try:
        selected = [sheet] if sheet is not None else list(workbook.sheetnames if sheets is None else sheets)
        if not selected:
            raise WorkbookExportError("CSV export requires at least one sheet")
        missing = [name for name in selected if name not in workbook.sheetnames]
        if missing:
            raise WorkbookExportError(f"unknown worksheet(s): {', '.join(missing)}")
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        outputs: list[Path] = []
        used_names: set[str] = set()
        for name in selected:
            output = destination_path if sheet is not None else destination_path / _portable_sheet_filename(name, used_names)
            if sheet is None and destination_path.suffix:
                raise WorkbookExportError("multi-sheet CSV destination must be a directory")
            if output.suffix.lower() != ".csv":
                raise WorkbookExportError("CSV output must end in .csv")
            output.parent.mkdir(parents=True, exist_ok=True)
            with output.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream, lineterminator="\n")
                for row in _cached_rows(workbook[name], values_workbook[name]):
                    writer.writerow(list(row))
            outputs.append(output)
        return tuple(outputs)
    finally:
        workbook.close()
        values_workbook.close()


def _cached_rows(formula_sheet, values_sheet):
    """Yield cached formula results, rejecting formulas without cached values."""

    for formula_row, values_row in zip(
        formula_sheet.iter_rows(values_only=False), values_sheet.iter_rows(values_only=True)
    ):
        values = []
        for formula_cell, cached in zip(formula_row, values_row):
            if isinstance(formula_cell.value, str) and formula_cell.value.startswith("=") and cached is None:
                raise WorkbookExportError(
                    f"formula has no cached result: {formula_sheet.title}!{formula_cell.coordinate}"
                )
            values.append(cached)
        yield values


def _portable_sheet_filename(title: str, used: set[str]) -> str:
    """Return a deterministic CSV filename safe on Windows and POSIX."""

    invalid = '<>:"/\\|?*'
    stem = "".join("_" if char in invalid or ord(char) < 32 else char for char in title).strip(" .")
    if not stem:
        stem = "sheet"
    if stem.upper() in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        stem = f"_{stem}"
    candidate = f"{stem}.csv"
    index = 2
    while candidate.casefold() in {item.casefold() for item in used}:
        candidate = f"{stem}-{index}.csv"
        index += 1
    used.add(candidate)
    return candidate
