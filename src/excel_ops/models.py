from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any


@dataclass(frozen=True)
class ExtractedRecord:
    location: str
    event_date: str
    identifier: str
    category: str
    confidence: float
    source: str
    source_file: str = ""
    source_sheet: str = ""
    source_row: int | None = None
    source_region: str = ""
    amount: int | float | None = None
    currency: str = ""

    @classmethod
    def from_dict(cls, value: dict[str, Any], source: str) -> "ExtractedRecord":
        raw_confidence = value.get("confidence", 0)
        try:
            confidence = float(raw_confidence)
        except (TypeError, ValueError):
            confidence = 0.0
        source_file = str(value.get("source_file") or source).strip()
        source_sheet = str(value.get("source_sheet") or "").strip()
        source_region = str(value.get("source_region") or "").strip()
        raw_row = value.get("source_row")
        try:
            source_row = int(raw_row) if raw_row is not None else None
        except (TypeError, ValueError):
            source_row = None
        locator = source_file
        if source_sheet:
            locator += f"#{source_sheet}"
        if source_row is not None:
            locator += f":{source_row}"
        if source_region:
            locator += f"@{source_region}"
        return cls(
            location=str(value.get("location", "")).strip(),
            event_date=str(value.get("event_date", "")).strip(),
            identifier=str(value.get("identifier", "")).strip(),
            category=str(value.get("category", "")).strip(),
            confidence=max(0.0, min(1.0, confidence)),
            source=locator,
            source_file=source_file,
            source_sheet=source_sheet,
            source_row=source_row,
            source_region=source_region,
            amount=_amount(value.get("amount")),
            currency=str(value.get("currency") or "").strip(),
        )


def _amount(value: Any) -> int | float | None:
    if value in (None, "") or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip().upper().replace(",", "")
    for token in ("CA$", "US$", "CAD", "USD", "CNY", "EUR", "$", "¥", "￥", "€"):
        text = text.replace(token, "")
    try:
        parsed = Decimal(text.strip())
    except InvalidOperation:
        return None
    if not parsed.is_finite():
        return None
    return int(parsed) if parsed == parsed.to_integral_value() else float(parsed)


@dataclass(frozen=True)
class ReviewedRecord:
    record: ExtractedRecord
    status: str
    reasons: tuple[str, ...]
