from __future__ import annotations

import re
from collections import Counter
from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable, Mapping


_IDENTIFIER_NAMES = re.compile(
    r"(^|[_\s-])(id|identifier|code|postal|postcode|zip|phone|telephone|mobile|employee|customer|client)([_\s-]|$)",
    re.IGNORECASE,
)
_DATE_NAMES = re.compile(r"(^|[_\s-])(date|datetime|timestamp|日期|时间)([_\s-]|$)", re.IGNORECASE)
_PERCENT_NAMES = re.compile(r"(^|[_\s-])(percent|percentage|rate|pct|百分比|比例)([_\s-]|$)", re.IGNORECASE)
_CURRENCY_NAMES = re.compile(r"(^|[_\s-])(amount|price|cost|salary|pay|revenue|total|金额|价格|工资)([_\s-]|$)", re.IGNORECASE)
_PHONE = re.compile(r"^\+?[\d() .-]{7,}$")
_ISO_DATE = re.compile(r"^\d{4}-\d{1,2}-\d{1,2}$")
_CHINESE_DATE = re.compile(r"^(\d{4})\s*[年/-]\s*(\d{1,2})\s*[月/-]\s*(\d{1,2})\s*日?$")
_SLASH_DATE = re.compile(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$")
_CURRENCY_SYMBOLS = {"$": "currency", "¥": "currency", "￥": "currency", "€": "currency", "£": "currency"}
_EXCEL_EPOCH = datetime(1899, 12, 30)


@dataclass(frozen=True)
class ValueInference:
    inferred_type: str
    confidence: float
    locale: str | None = None
    unit: str | None = None
    normalized_value: Any = None
    ambiguous: bool = False
    reason: str | None = None


@dataclass(frozen=True)
class FieldInference:
    field: str
    inferred_type: str
    confidence: float
    locale: str | None
    unit: str | None
    sample_count: int
    low_confidence_count: int
    ambiguous_count: int
    leading_zero_preserved: bool

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _is_identifier_field(field: str) -> bool:
    return bool(_IDENTIFIER_NAMES.search(field.replace(".", "_")))


def _parse_number(text: str, locale: str | None) -> tuple[Decimal | None, str | None, bool]:
    compact = text.strip().replace("\u00a0", " ").replace(" ", "")
    for symbol in _CURRENCY_SYMBOLS:
        compact = compact.replace(symbol, "")
    compact = compact.rstrip("%").strip()
    if not compact:
        return None, locale, False

    comma = compact.rfind(",")
    dot = compact.rfind(".")
    inferred_locale = locale
    ambiguous = False
    if comma >= 0 and dot >= 0:
        decimal_mark = "," if comma > dot else "."
    elif comma >= 0:
        tail = len(compact) - comma - 1
        if locale in {"de-DE", "fr-FR"}:
            decimal_mark = ","
        elif locale in {"en-US", "en-CA", "en-GB", "zh-CN"}:
            decimal_mark = "."
        elif tail == 3:
            decimal_mark = "."
            ambiguous = True
        else:
            decimal_mark = ","
            inferred_locale = "de-DE"
    else:
        decimal_mark = "."

    normalized = compact
    if decimal_mark == ",":
        normalized = normalized.replace(".", "").replace(",", ".")
    else:
        normalized = normalized.replace(",", "")
    try:
        return Decimal(normalized), inferred_locale, ambiguous
    except InvalidOperation:
        return None, inferred_locale, False


def infer_value(field: str, value: Any, locale: str | None = None) -> ValueInference:
    """Infer a value without coercing identifiers or resolving ambiguous dates."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return ValueInference("empty", 1.0, locale=locale)

    text = str(value).strip()
    if _is_identifier_field(field):
        reason = "identifier_field"
        if text.startswith("0") and text.isdigit() and len(text) > 1:
            reason = "leading_zero_identifier"
        elif _PHONE.fullmatch(text) and ("phone" in field.lower() or "电话" in field):
            reason = "phone_identifier"
        return ValueInference("text_identifier", 0.99, locale=locale, normalized_value=text, reason=reason)

    date_field = bool(_DATE_NAMES.search(field))
    if isinstance(value, (date, datetime)):
        normalized = value.isoformat()
        return ValueInference("datetime" if isinstance(value, datetime) else "date", 1.0, locale=locale, normalized_value=normalized)

    chinese = _CHINESE_DATE.fullmatch(text)
    if chinese:
        try:
            parsed = date(*(int(part) for part in chinese.groups()))
        except ValueError:
            return ValueInference("date", 0.2, locale="zh-CN", ambiguous=True, reason="invalid_date")
        return ValueInference("date", 0.99, locale="zh-CN", normalized_value=parsed.isoformat())

    if _ISO_DATE.fullmatch(text):
        try:
            return ValueInference("date", 1.0, locale=locale, normalized_value=date.fromisoformat(text).isoformat())
        except ValueError:
            return ValueInference("date", 0.2, locale=locale, ambiguous=True, reason="invalid_date")

    slash = _SLASH_DATE.fullmatch(text)
    if slash:
        first, second, year = (int(part) for part in slash.groups())
        if first <= 12 and second <= 12 and locale not in {"en-US", "en-CA", "en-GB", "zh-CN"}:
            return ValueInference("date", 0.4, ambiguous=True, reason="ambiguous_day_month")
        month_first = locale in {"en-US", "en-CA"} or (first <= 12 and second > 12)
        month, day = (first, second) if month_first else (second, first)
        try:
            parsed = date(year, month, day)
        except ValueError:
            return ValueInference("date", 0.2, locale=locale, ambiguous=True, reason="invalid_date")
        return ValueInference("date", 0.95 if locale else 0.9, locale=locale, normalized_value=parsed.isoformat())

    number, inferred_locale, number_ambiguous = _parse_number(text, locale)
    if number is not None:
        if date_field and Decimal(1) <= number <= Decimal(2958465) and number == number.to_integral_value():
            parsed = _EXCEL_EPOCH + timedelta(days=int(number))
            return ValueInference("excel_serial_date", 0.94, locale=locale, normalized_value=parsed.date().isoformat())
        if text.startswith("0") and text.isdigit() and len(text) > 1:
            return ValueInference("text_identifier", 0.7, locale=locale, normalized_value=text, reason="leading_zero_candidate")
        if text.endswith("%") or _PERCENT_NAMES.search(field):
            normalized = number / 100 if text.endswith("%") else number
            return ValueInference("percentage", 0.96, locale=inferred_locale, unit="percent", normalized_value=str(normalized), ambiguous=number_ambiguous)
        currency = next((symbol for symbol in _CURRENCY_SYMBOLS if symbol in text), None)
        if currency or _CURRENCY_NAMES.search(field):
            return ValueInference("decimal", 0.94, locale=inferred_locale, unit="currency", normalized_value=str(number), ambiguous=number_ambiguous)
        inferred_type = "integer" if number == number.to_integral_value() else "decimal"
        return ValueInference(inferred_type, 0.88, locale=inferred_locale, normalized_value=str(number), ambiguous=number_ambiguous, reason="text_number")

    return ValueInference("text", 0.9, locale=locale, normalized_value=text)


def infer_fields(
    records: Iterable[Mapping[str, Any]], locale: str | None = None, low_confidence_threshold: float = 0.8
) -> list[FieldInference]:
    rows = list(records)
    fields = list(dict.fromkeys(key for row in rows for key in row))
    report: list[FieldInference] = []
    for field in fields:
        values = [row.get(field) for row in rows if row.get(field) not in (None, "")]
        inferred = [infer_value(field, value, locale) for value in values]
        counts = Counter(item.inferred_type for item in inferred)
        dominant, dominant_count = counts.most_common(1)[0] if counts else ("empty", 0)
        confidence = sum(item.confidence for item in inferred) / len(inferred) if inferred else 1.0
        if inferred and dominant_count != len(inferred):
            confidence *= dominant_count / len(inferred)
        units = Counter(item.unit for item in inferred if item.unit)
        locales = Counter(item.locale for item in inferred if item.locale)
        report.append(
            FieldInference(
                field=field,
                inferred_type=dominant,
                confidence=round(confidence, 3),
                locale=locales.most_common(1)[0][0] if locales else locale,
                unit=units.most_common(1)[0][0] if units else None,
                sample_count=len(inferred),
                low_confidence_count=sum(item.confidence < low_confidence_threshold for item in inferred),
                ambiguous_count=sum(item.ambiguous for item in inferred),
                leading_zero_preserved=any(item.reason in {"leading_zero_identifier", "leading_zero_candidate"} for item in inferred),
            )
        )
    return report
