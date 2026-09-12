"""Deterministic report-period resolution for agent workflows."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError


class PeriodResolutionError(ValueError):
    """Raised when a report period cannot be resolved safely."""


class AmbiguousPeriodError(PeriodResolutionError):
    """Raised when wording has more than one reasonable interpretation."""


@dataclass(frozen=True)
class PeriodResult:
    """A resolved interval plus the rules needed to reproduce it."""

    period_start: date
    period_end: date
    display_text: str
    as_of_date: date
    timezone: str
    timezone_applied: bool
    rule: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return {
            "period_start": self.period_start.isoformat(),
            "period_end": self.period_end.isoformat(),
            "display_text": self.display_text,
            "as_of_date": self.as_of_date.isoformat(),
            "timezone": self.timezone,
            "timezone_applied": self.timezone_applied,
            "rule": dict(self.rule),
        }


_EXPRESSIONS = {
    "本周": "current_week",
    "this week": "current_week",
    "上周": "previous_complete_week",
    "上一个完整周": "previous_complete_week",
    "last complete week": "previous_complete_week",
    "previous complete week": "previous_complete_week",
    "本月": "current_month",
    "this month": "current_month",
    "上个月": "previous_month",
    "last month": "previous_month",
    "本季度": "current_quarter",
    "this quarter": "current_quarter",
    "上季度": "previous_quarter",
    "last quarter": "previous_quarter",
    "本财年": "current_fiscal_year",
    "this fiscal year": "current_fiscal_year",
    "上财年": "previous_fiscal_year",
    "last fiscal year": "previous_fiscal_year",
    "本财季": "current_fiscal_quarter",
    "this fiscal quarter": "current_fiscal_quarter",
    "上财季": "previous_fiscal_quarter",
    "last fiscal quarter": "previous_fiscal_quarter",
    "自定义": "custom",
    "custom": "custom",
}

_AMBIGUOUS = {"最近一周", "近一周", "last week", "recent week", "past week"}


def _coerce_date(value: date | datetime | str | None, timezone: str) -> tuple[date, bool]:
    try:
        zone = ZoneInfo(timezone)
    except (ZoneInfoNotFoundError, ValueError) as exc:
        raise PeriodResolutionError(f"Unknown timezone: {timezone}") from exc

    if value is None:
        return datetime.now(zone).date(), True
    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.replace(tzinfo=zone).date(), False
        converted = value.astimezone(zone)
        return converted.date(), converted.date() != value.date()
    if isinstance(value, date):
        return value, False
    if isinstance(value, str):
        text = value.strip()
        try:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError as exc:
            try:
                return date.fromisoformat(text), False
            except ValueError:
                raise PeriodResolutionError(f"Invalid as_of date or datetime: {value}") from exc
        if parsed.tzinfo is None:
            return parsed.date(), False
        converted = parsed.astimezone(zone)
        return converted.date(), converted.date() != parsed.date()
    raise PeriodResolutionError("as_of must be a date, datetime, ISO string, or None")


def _month_start(day: date) -> date:
    return day.replace(day=1)


def _add_months(day: date, months: int) -> date:
    index = day.year * 12 + day.month - 1 + months
    return date(index // 12, index % 12 + 1, 1)


def _month_end(day: date) -> date:
    return _add_months(_month_start(day), 1) - timedelta(days=1)


def _fiscal_year_start(day: date, start_month: int) -> date:
    year = day.year if day.month >= start_month else day.year - 1
    return date(year, start_month, 1)


def _parse_custom(value: date | str | None, field: str) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value.strip())
        except ValueError as exc:
            raise PeriodResolutionError(f"{field} must be an ISO date") from exc
    raise PeriodResolutionError(f"{field} is required for a custom period")


def resolve_period(
    expression: str,
    *,
    as_of: date | datetime | str | None = None,
    timezone: str = "UTC",
    week_start: int = 0,
    fiscal_year_start_month: int = 1,
    custom_start: date | str | None = None,
    custom_end: date | str | None = None,
) -> PeriodResult:
    """Resolve a named business period without guessing ambiguous wording.

    ``week_start`` uses Python weekday numbers (Monday=0, Sunday=6). The
    default is therefore explicitly Monday--Sunday. The returned ``rule``
    stores inputs and policy, never the calculated dates.
    """

    if not isinstance(expression, str) or not expression.strip():
        raise PeriodResolutionError("A report-period expression is required")
    normalized = " ".join(expression.strip().lower().split())
    if normalized in _AMBIGUOUS:
        raise AmbiguousPeriodError(
            f"'{expression}' is ambiguous; choose a rolling seven-day window or 上一个完整周"
        )
    kind = _EXPRESSIONS.get(normalized)
    if kind is None:
        raise PeriodResolutionError(f"Unsupported report period: {expression}")
    if not 0 <= week_start <= 6:
        raise PeriodResolutionError("week_start must be between 0 (Monday) and 6 (Sunday)")
    if not 1 <= fiscal_year_start_month <= 12:
        raise PeriodResolutionError("fiscal_year_start_month must be between 1 and 12")

    anchor, timezone_applied = _coerce_date(as_of, timezone)
    rule: dict[str, Any] = {
        "expression": expression.strip(),
        "period_kind": kind,
        "week_start": week_start,
        "fiscal_year_start_month": fiscal_year_start_month,
        "timezone": timezone,
    }

    if kind == "custom":
        start = _parse_custom(custom_start, "custom_start")
        end = _parse_custom(custom_end, "custom_end")
        if start > end:
            raise PeriodResolutionError("custom_start must not be after custom_end")
        rule["custom_start"] = start.isoformat()
        rule["custom_end"] = end.isoformat()
    elif kind in {"current_week", "previous_complete_week"}:
        start = anchor - timedelta(days=(anchor.weekday() - week_start) % 7)
        if kind == "previous_complete_week":
            start -= timedelta(days=7)
        end = start + timedelta(days=6)
    elif kind == "current_month":
        start, end = _month_start(anchor), _month_end(anchor)
    elif kind == "previous_month":
        start = _add_months(_month_start(anchor), -1)
        end = _month_end(start)
    elif kind in {"current_quarter", "previous_quarter"}:
        start = date(anchor.year, ((anchor.month - 1) // 3) * 3 + 1, 1)
        if kind == "previous_quarter":
            start = _add_months(start, -3)
        end = _add_months(start, 3) - timedelta(days=1)
    else:
        fiscal_start = _fiscal_year_start(anchor, fiscal_year_start_month)
        if kind in {"current_fiscal_year", "previous_fiscal_year"}:
            start = fiscal_start
            if kind == "previous_fiscal_year":
                start = _add_months(start, -12)
            end = _add_months(start, 12) - timedelta(days=1)
        else:
            months_since_start = (anchor.year - fiscal_start.year) * 12 + anchor.month - fiscal_start.month
            start = _add_months(fiscal_start, (months_since_start // 3) * 3)
            if kind == "previous_fiscal_quarter":
                start = _add_months(start, -3)
            end = _add_months(start, 3) - timedelta(days=1)

    return PeriodResult(
        period_start=start,
        period_end=end,
        display_text=f"{start.isoformat()} 至 {end.isoformat()}",
        as_of_date=anchor,
        timezone=timezone,
        timezone_applied=timezone_applied,
        rule=rule,
    )
