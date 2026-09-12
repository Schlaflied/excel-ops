from datetime import datetime, timezone

import pytest

from excel_ops.periods import AmbiguousPeriodError, PeriodResolutionError, resolve_period


def test_previous_complete_monday_sunday_week_is_deterministic():
    result = resolve_period("上一个完整周", as_of="2026-09-09", timezone="America/Toronto")
    assert result.to_dict() == {
        "period_start": "2026-08-31",
        "period_end": "2026-09-06",
        "display_text": "2026-08-31 至 2026-09-06",
        "as_of_date": "2026-09-09",
        "timezone": "America/Toronto",
        "timezone_applied": False,
        "rule": {
            "expression": "上一个完整周",
            "period_kind": "previous_complete_week",
            "week_start": 0,
            "fiscal_year_start_month": 1,
            "timezone": "America/Toronto",
        },
    }


@pytest.mark.parametrize(
    ("expression", "as_of", "start", "end"),
    [
        ("上个月", "2026-01-10", "2025-12-01", "2025-12-31"),
        ("本月", "2024-02-29", "2024-02-01", "2024-02-29"),
        ("上季度", "2026-01-01", "2025-10-01", "2025-12-31"),
    ],
)
def test_calendar_boundaries(expression, as_of, start, end):
    result = resolve_period(expression, as_of=as_of)
    assert result.period_start.isoformat() == start
    assert result.period_end.isoformat() == end


def test_timezone_boundary_is_recorded_when_it_changes_the_date():
    result = resolve_period(
        "本月",
        as_of=datetime(2026, 3, 1, 1, tzinfo=timezone.utc),
        timezone="America/Toronto",
    )
    assert result.as_of_date.isoformat() == "2026-02-28"
    assert result.timezone_applied is True
    assert result.period_end.isoformat() == "2026-02-28"


def test_recent_week_is_not_silently_treated_as_complete_week():
    with pytest.raises(AmbiguousPeriodError, match="rolling seven-day"):
        resolve_period("最近一周", as_of="2026-09-09")


def test_custom_period_requires_explicit_valid_bounds():
    result = resolve_period(
        "自定义", as_of="2026-09-09", custom_start="2025-12-29", custom_end="2026-01-04"
    )
    assert result.period_start.isoformat() == "2025-12-29"
    assert result.period_end.isoformat() == "2026-01-04"
    with pytest.raises(PeriodResolutionError, match="must not be after"):
        resolve_period("自定义", custom_start="2026-02-01", custom_end="2026-01-01")


def test_fiscal_periods_use_configured_start_month():
    quarter = resolve_period("本财季", as_of="2026-02-15", fiscal_year_start_month=4)
    assert quarter.period_start.isoformat() == "2026-01-01"
    assert quarter.period_end.isoformat() == "2026-03-31"

    year = resolve_period("本财年", as_of="2026-02-15", fiscal_year_start_month=4)
    assert year.period_start.isoformat() == "2025-04-01"
    assert year.period_end.isoformat() == "2026-03-31"


def test_recipe_keeps_rules_instead_of_resolved_dates():
    rule = resolve_period("上一个完整周", as_of="2026-09-09").rule
    assert "period_start" not in rule
    assert "period_end" not in rule
    assert rule["period_kind"] == "previous_complete_week"


def test_sunday_week_start_is_explicitly_supported():
    result = resolve_period("上一个完整周", as_of="2026-09-09", week_start=6)
    assert result.period_start.isoformat() == "2026-08-30"
    assert result.period_end.isoformat() == "2026-09-05"


@pytest.mark.parametrize("timezone_name", ["Not/A_Zone", ""])
def test_invalid_timezone_is_rejected(timezone_name):
    with pytest.raises(PeriodResolutionError, match="Unknown timezone"):
        resolve_period("本月", as_of="2026-09-09", timezone=timezone_name)
