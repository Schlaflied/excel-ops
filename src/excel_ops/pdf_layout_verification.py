"""Conservative page-level verification for rendered workbook PDFs."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Sequence

from openpyxl import load_workbook
from pypdf import PdfReader


@dataclass(frozen=True)
class PdfLayoutFinding:
    """One machine-readable reason a PDF layout could not be verified."""

    code: str
    message: str
    sheet: str | None = None
    page: int | None = None

    def to_dict(self) -> dict[str, Any]:
        """Return JSON-safe evidence without workbook cell values."""

        return asdict(self)


@dataclass(frozen=True)
class PdfLayoutVerification:
    """Page-level verification result for one selected-sheet PDF."""

    passed: bool
    pages: int
    findings: tuple[PdfLayoutFinding, ...]


def verify_pdf_layout(
    workbook_path: str | Path,
    pdf_path: str | Path,
    *,
    sheets: Sequence[str],
) -> PdfLayoutVerification:
    """Check pagination, page bounds, readable text, headers, and edge content."""

    findings: list[PdfLayoutFinding] = []
    try:
        reader = PdfReader(str(pdf_path))
        pages = tuple(reader.pages)
    except Exception as error:
        return PdfLayoutVerification(
            False,
            0,
            (PdfLayoutFinding("pdf_unreadable", f"PDF could not be parsed: {error}"),),
        )

    if len(pages) < len(sheets):
        findings.append(
            PdfLayoutFinding(
                "page_count_too_small",
                "The PDF has fewer pages than selected worksheets.",
            )
        )

    page_text: list[str] = []
    for index, page in enumerate(pages, start=1):
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        if width <= 0 or height <= 0 or width > 2000 or height > 2000:
            findings.append(
                PdfLayoutFinding(
                    "invalid_page_bounds",
                    "The rendered page has invalid or unreasonable dimensions.",
                    page=index,
                )
            )
        text = page.extract_text() or ""
        page_text.append(_normalize(text))
        if not text.strip():
            findings.append(
                PdfLayoutFinding(
                    "unreadable_page",
                    "No readable text could be extracted from the rendered page.",
                    page=index,
                )
            )

    workbook = load_workbook(workbook_path, read_only=False, data_only=False)
    try:
        rendered_text = " ".join(page_text)
        for sheet_name in sheets:
            worksheet = workbook[sheet_name]
            if worksheet.max_row > 50 and not worksheet.print_title_rows:
                findings.append(
                    PdfLayoutFinding(
                        "repeated_header_not_configured",
                        "A long worksheet has no repeated print-title rows configured.",
                        sheet=sheet_name,
                    )
                )
            if not _has_safe_scaling(worksheet):
                findings.append(
                    PdfLayoutFinding(
                        "scaling_not_bounded",
                        "Worksheet print scaling is not bounded to a supported range.",
                        sheet=sheet_name,
                    )
                )
            for token in _boundary_tokens(worksheet):
                if token not in rendered_text:
                    findings.append(
                        PdfLayoutFinding(
                            "boundary_content_missing",
                            "Printable boundary content was not found in the rendered PDF.",
                            sheet=sheet_name,
                        )
                    )
                    break
            if len(sheets) == 1 and len(pages) > 1 and worksheet.print_title_rows:
                title_tokens = _title_tokens(worksheet)
                if title_tokens and any(
                    any(token not in text for token in title_tokens) for text in page_text
                ):
                    findings.append(
                        PdfLayoutFinding(
                            "repeated_header_missing",
                            "Configured print-title content was not found on every page.",
                            sheet=sheet_name,
                        )
                    )
    finally:
        workbook.close()

    return PdfLayoutVerification(not findings, len(pages), tuple(findings))


def _has_safe_scaling(worksheet) -> bool:
    """Accept fit-to-page or an explicit, bounded percentage scale."""

    setup = worksheet.page_setup
    if setup.fitToWidth is not None:
        return 1 <= int(setup.fitToWidth) <= 2
    if setup.scale is not None:
        return 10 <= int(setup.scale) <= 400
    return True


def _boundary_tokens(worksheet) -> tuple[str, ...]:
    """Return non-sensitive normalized strings from the printable data edges."""

    first: str | None = None
    last: str | None = None
    for row in worksheet.iter_rows():
        populated = [cell.value for cell in row if cell.value not in (None, "")]
        if populated:
            first = first or str(populated[0])
            last = str(populated[-1])
    return tuple(
        dict.fromkeys(
            token
            for value in (first, last)
            if value is not None
            for token in (_normalize(value),)
            if len(token) >= 3
        )
    )


def _normalize(value: str) -> str:
    """Normalize renderer whitespace for conservative text-presence checks."""

    return re.sub(r"\s+", " ", value).strip().casefold()


def _title_tokens(worksheet) -> tuple[str, ...]:
    """Return stable text tokens from configured repeated title rows."""

    title_range = worksheet.print_title_rows
    if not title_range:
        return ()
    bounds = title_range.replace("$", "").split(":", maxsplit=1)
    start, end = (int(bounds[0]), int(bounds[-1]))
    tokens = []
    for row in worksheet.iter_rows(min_row=start, max_row=end):
        for cell in row:
            token = _normalize(str(cell.value or ""))
            if len(token) >= 3:
                tokens.append(token)
    return tuple(dict.fromkeys(tokens))
