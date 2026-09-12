from dataclasses import dataclass
from datetime import date
from pathlib import Path

import pytest

from excel_ops.naming import OutputNamingError, resolve_output_path, verify_output_name


@dataclass(frozen=True)
class FixedPeriod:
    period_start: date
    period_end: date
    as_of_date: date


@pytest.fixture
def period() -> FixedPeriod:
    return FixedPeriod(date(2026, 9, 1), date(2026, 9, 7), date(2026, 9, 12))


def test_prompt_slashes_are_safe_filename_separators(tmp_path: Path, period: FixedPeriod):
    result = resolve_output_path(
        "payroll-YYYY/MM/DD", output_dir=tmp_path, extension="xlsx", period=period
    )

    assert result.path == tmp_path / "payroll-2026-09-12.xlsx"
    assert result.original_pattern == "payroll-YYYY/MM/DD"


def test_directory_mode_requires_explicit_opt_in(tmp_path: Path, period: FixedPeriod):
    result = resolve_output_path(
        "payroll/YYYY/MM/DD/payroll",
        output_dir=tmp_path,
        extension=".xlsx",
        period=period,
        directory_mode=True,
    )

    assert result.path == tmp_path / "payroll" / "2026" / "09" / "12" / "payroll.xlsx"


def test_uses_resolved_business_period_not_system_date(tmp_path: Path, period: FixedPeriod):
    result = resolve_output_path(
        "report-{period_start}-to-{period_end}-{YYYY}{MM}{DD}",
        output_dir=tmp_path,
        extension="pdf",
        period=period,
    )

    assert result.path.name == "report-2026-09-01-to-2026-09-07-20260912.pdf"


def test_business_tokens_and_illegal_characters_are_sanitised(tmp_path: Path, period: FixedPeriod):
    result = resolve_output_path(
        "{client}:{department}?{workflow}",
        output_dir=tmp_path,
        extension="csv",
        period=period,
        values={"client": "ACME", "department": "Finance", "workflow": "Payroll*Final"},
    )

    assert result.path.name == "ACME-Finance-Payroll-Final.csv"


def test_identical_collision_is_noop_and_different_content_gets_revision(
    tmp_path: Path, period: FixedPeriod
):
    existing = tmp_path / "payroll-2026-09-12.xlsx"
    existing.write_bytes(b"same")

    noop = resolve_output_path(
        "payroll-YYYY-MM-DD",
        output_dir=tmp_path,
        extension="xlsx",
        period=period,
        content=b"same",
    )
    revision = resolve_output_path(
        "payroll-YYYY-MM-DD",
        output_dir=tmp_path,
        extension="xlsx",
        period=period,
        content=b"different",
    )

    assert noop.path == existing
    assert noop.collision == "identical_noop"
    assert revision.path == tmp_path / "payroll-2026-09-12-r2.xlsx"
    assert revision.collision == "revision"


def test_revision_skips_existing_suffixes(tmp_path: Path, period: FixedPeriod):
    (tmp_path / "payroll-2026-09-12.xlsx").write_bytes(b"one")
    (tmp_path / "payroll-2026-09-12-r2.xlsx").write_bytes(b"two")

    result = resolve_output_path(
        "payroll-YYYY-MM-DD",
        output_dir=tmp_path,
        extension="xlsx",
        period=period,
    )

    assert result.path.name == "payroll-2026-09-12-r3.xlsx"


def test_unknown_braced_token_stops_resolution(tmp_path: Path, period: FixedPeriod):
    with pytest.raises(OutputNamingError, match="unknown or unavailable"):
        resolve_output_path(
            "report-{unknown}", output_dir=tmp_path, extension="xlsx", period=period
        )


def test_verify_checks_disk_and_period_alignment(tmp_path: Path, period: FixedPeriod):
    result = resolve_output_path(
        "payroll-YYYY-MM-DD", output_dir=tmp_path, extension="xlsx", period=period
    )
    result.path.write_bytes(b"workbook")

    verify_output_name(
        result,
        period=period,
        manifest={"period_start": "2026-09-01", "period_end": "2026-09-07"},
    )

    with pytest.raises(OutputNamingError, match="manifest period_end"):
        verify_output_name(
            result,
            period=period,
            manifest={"period_start": "2026-09-01", "period_end": "2026-09-08"},
        )

    with pytest.raises(OutputNamingError, match="was not written"):
        verify_output_name(
            resolve_output_path(
                "other-YYYY-MM-DD", output_dir=tmp_path, extension="xlsx", period=period
            ),
            period=period,
        )
