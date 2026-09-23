from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from math import isfinite
from typing import Any


_CURRENCY_TOKENS = (
    ("CA$", "CAD"),
    ("US$", "USD"),
    ("CAD", "CAD"),
    ("USD", "USD"),
    ("CNY", "CNY"),
    ("EUR", "EUR"),
    ("€", "EUR"),
    ("¥", "¥"),
    ("￥", "￥"),
    ("$", "$"),
)
_AMBIGUOUS_CURRENCY_MATCHES = {
    "$": frozenset({"CAD", "USD", "$"}),
    "¥": frozenset({"CNY", "¥"}),
    "￥": frozenset({"CNY", "￥"}),
}


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
        amount, detected_currency = _amount(value.get("amount"))
        explicit_currency = _currency_evidence(value.get("currency"))
        if detected_currency and explicit_currency:
            compatible = _AMBIGUOUS_CURRENCY_MATCHES.get(detected_currency)
            if compatible is not None and explicit_currency in compatible:
                currency = explicit_currency
            elif detected_currency != explicit_currency:
                raise ValueError(
                    f"amount currency {detected_currency} conflicts with currency field "
                    f"{explicit_currency}"
                )
            else:
                currency = explicit_currency
        else:
            currency = explicit_currency or detected_currency or ""
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
            amount=amount,
            currency=currency,
        )


def _amount(value: Any) -> tuple[int | float | None, str | None]:
    if value in (None, "") or isinstance(value, bool):
        return None, None
    if isinstance(value, int):
        return value, None
    if isinstance(value, float):
        return (value, None) if isfinite(value) else (None, None)
    text = str(value).strip().upper().replace(",", "")
    detected_currency = None
    for token, normalized in _CURRENCY_TOKENS:
        if token in text:
            detected_currency = normalized
            text = text.replace(token, "")
            break
    try:
        parsed = Decimal(text.strip())
    except InvalidOperation:
        return None, detected_currency
    if not parsed.is_finite():
        return None, detected_currency
    numeric = int(parsed) if parsed == parsed.to_integral_value() else float(parsed)
    return numeric, detected_currency


def _currency_evidence(value: Any) -> str | None:
    text = str(value or "").strip().upper()
    if not text:
        return None
    for token, normalized in _CURRENCY_TOKENS:
        if text == token:
            return normalized
    return text


@dataclass(frozen=True)
class ReviewedRecord:
    record: ExtractedRecord
    status: str
    reasons: tuple[str, ...]
