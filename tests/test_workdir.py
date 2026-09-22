import json
import os
from datetime import date, datetime, timezone
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_ops import workdir
from excel_ops.cli import main
from excel_ops.workdir import (
    EXCLUDE,
    INCLUDE,
    INPUT,
    PRIOR_DELIVERY,
    REVIEW,
    REVIEW_RETURN,
    TEMPLATE,
    UNKNOWN,
    ClassificationOverride,
    PeriodWindow,
    ScanScope,
    StabilitySample,
    UnauthorizedPathError,
    WorkdirScanError,
    format_dry_run,
    load_workdir_recipe,
    override_from_entry,
    save_workdir_recipe,
    scan_workdir,
    version_key,
)


NOW = datetime(2026, 9, 21, 12, 0, tzinfo=timezone.utc)
PERIOD = PeriodWindow.of(date(2026, 9, 1), date(2026, 9, 30))
IN_PERIOD = datetime(2026, 9, 10, tzinfo=timezone.utc).timestamp()
BEFORE_PERIOD = datetime(2026, 8, 5, tzinfo=timezone.utc).timestamp()


def _file(directory: Path, name: str, content: str = "a,b\n1,2\n", *, when: float = IN_PERIOD) -> Path:
    path = directory / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    os.utime(path, (when, when))
    return path


def _workbook(directory: Path, name: str, *, sheet: str = "Data", when: float = IN_PERIOD, **properties) -> Path:
    path = directory / name
    path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    book.active.title = sheet
    book.active.append(["Location", "Date"])
    for key, value in properties.items():
        setattr(book.properties, key, value)
    book.save(path)
    book.close()
    os.utime(path, (when, when))
    return path


def _scan(root, **kwargs):
    options = {"period_window": PERIOD, "now": NOW, "stability_window_seconds": 5.0}
    options.update(kwargs)
    return scan_workdir(root, **options)


def test_synthetic_directory_separates_current_period_input_from_history(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "september-hours.csv")
    _file(work, "august-hours.csv", "a,b\n9,9\n", when=BEFORE_PERIOD)
    _file(work, "north-template.xlsx", "template-bytes")
    _file(work, "2026-08-delivery.xlsx", "delivered-bytes", when=BEFORE_PERIOD)
    _file(work, "review-pack-return.xlsx", "review-bytes")
    _file(work, "notes.txt", "free text")

    result = _scan(work)

    assert [item.relative_path for item in result.included] == ["september-hours.csv"]
    assert result.entry("september-hours.csv").classification == INPUT
    assert result.entry("september-hours.csv").period_state == "current"

    historical = result.entry("august-hours.csv")
    assert historical.classification == PRIOR_DELIVERY
    assert historical.disposition == EXCLUDE
    assert historical.reason == "historical_outside_current_period"
    assert historical.period_state == "historical"

    assert result.entry("north-template.xlsx").classification == TEMPLATE
    assert result.entry("2026-08-delivery.xlsx").classification == PRIOR_DELIVERY
    assert result.entry("review-pack-return.xlsx").classification == REVIEW_RETURN
    unsupported = result.entry("notes.txt")
    assert (unsupported.classification, unsupported.disposition) == (UNKNOWN, EXCLUDE)
    assert unsupported.reason == "unsupported_extension"
    assert result.counts()["include"] == 1


def test_workbook_metadata_classifies_when_the_filename_says_nothing(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _workbook(work, "q3-figures.xlsx", sheet="Master Template")
    _workbook(work, "returned-file.xlsx", sheet="Sheet1", category="Review Pack")

    result = _scan(work)

    metadata_entry = result.entry("q3-figures.xlsx")
    assert metadata_entry.classification == TEMPLATE
    assert metadata_entry.signal == "workbook_metadata"
    assert result.entry("returned-file.xlsx").classification == REVIEW_RETURN


def test_identical_content_under_different_names_is_processed_once(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "hours.csv", "loc,date\nNorth,2026-09-02\n")
    _file(work, "hours-emailed-copy.csv", "loc,date\nNorth,2026-09-02\n")
    _file(work, "other-hours.csv", "loc,date\nSouth,2026-09-03\n")

    result = _scan(work)

    assert len(result.duplicate_groups) == 1
    group = result.duplicate_groups[0]
    assert group.paths == ("hours-emailed-copy.csv", "hours.csv")
    assert group.representative == "hours.csv"
    included = {item.relative_path for item in result.included}
    assert included == {"hours.csv", "other-hours.csv"}
    duplicate = result.entry("hours-emailed-copy.csv")
    assert (duplicate.disposition, duplicate.reason) == (EXCLUDE, "duplicate_content")
    assert duplicate.duplicate_of == "hours.csv"
    # The duplicate is content-identical, not a version question.
    assert result.version_groups == ()


def test_version_candidates_are_grouped_and_none_is_auto_selected(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "final.xlsx", "one")
    _file(work, "final (1).xlsx", "two")
    _file(work, "final-final.xlsx", "three")

    result = _scan(work)

    assert len(result.version_groups) == 1
    group = result.version_groups[0]
    assert group.version_key == "final"
    assert group.paths == ("final (1).xlsx", "final-final.xlsx", "final.xlsx")
    assert group.selected is None
    assert group.reason == "version_candidates_require_confirmation"
    assert result.included == ()
    for name in group.paths:
        entry = result.entry(name)
        assert entry.disposition == REVIEW
        assert entry.reason == "version_candidates_require_confirmation"
    assert "no version selected" in format_dry_run(result)


def test_version_key_collapses_common_decorations():
    assert version_key("final.xlsx") == version_key("final (1).xlsx") == version_key("final-final.xlsx")
    assert version_key("payroll-v2.xlsx") == version_key("payroll - Copy.xlsx") == "payroll"
    assert version_key("hours.csv") != version_key("wages.csv")


def test_a_file_still_being_written_is_never_read_or_included(tmp_path, monkeypatch):
    work = tmp_path / "work"
    work.mkdir()
    settled = _file(work, "settled-hours.csv")
    mid_write = _file(work, "mid-write-hours.csv", "a,b\n", when=NOW.timestamp() - 1)

    opened: list[str] = []
    real_read = workdir._read_bytes
    monkeypatch.setattr(
        workdir, "_read_bytes", lambda path: (opened.append(str(path)), real_read(path))[1]
    )

    result = _scan(work)

    unstable = result.entry("mid-write-hours.csv")
    assert unstable.stable is False
    assert unstable.disposition == REVIEW
    assert unstable.reason == "not_stable_yet"
    assert unstable.content_hash is None
    assert str(mid_write) not in opened
    assert str(settled) in opened
    assert [item.relative_path for item in result.included] == ["settled-hours.csv"]


def test_a_file_that_changed_since_the_previous_sample_is_not_trusted(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "hours.csv")

    first = _scan(work)
    assert first.entry("hours.csv").disposition == INCLUDE

    samples = first.stability_samples()
    stale = {"hours.csv": StabilitySample(samples["hours.csv"].size + 10, samples["hours.csv"].modified_at)}
    second = _scan(work, previous_samples=stale)

    changed = second.entry("hours.csv")
    assert changed.stable is False
    assert changed.reason == "not_stable_yet"
    # Re-running with the matching sample trusts the now settled file again.
    third = _scan(work, previous_samples=samples)
    assert third.entry("hours.csv").disposition == INCLUDE


def test_cloud_sync_artifacts_are_flagged_instead_of_ingested(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "~$hours.xlsx", "lock")
    _file(work, "budget.xlsx.crdownload", "partial")
    _file(work, "hours.xlsx.part", "partial")
    _file(work, "plan (conflicted copy from LAPTOP-7).xlsx", "conflict")
    _file(work, "工资表 (冲突副本).xlsx", "conflict")

    result = _scan(work)

    assert result.included == ()
    reasons = {item.relative_path: item.reason for item in result.files}
    assert reasons["~$hours.xlsx"] == "sync_artifact_lock_file"
    assert reasons["budget.xlsx.crdownload"] == "sync_artifact_incomplete_download"
    assert reasons["hours.xlsx.part"] == "sync_artifact_incomplete_download"
    # A conflict copy needs a human decision; this module never resolves one.
    conflicts = [
        item for item in result.files if item.reason == "sync_artifact_conflict_copy"
    ]
    assert {item.relative_path for item in conflicts} == {
        "plan (conflicted copy from LAPTOP-7).xlsx",
        "工资表 (冲突副本).xlsx",
    }
    assert all(item.disposition == REVIEW for item in conflicts)
    assert all(item.content_hash is None for item in conflicts)


def test_files_outside_the_authorized_scope_are_never_accessed(tmp_path, monkeypatch):
    authorized = tmp_path / "authorized"
    authorized.mkdir()
    _file(authorized, "hours.csv")
    forbidden = tmp_path / "forbidden"
    forbidden.mkdir()
    _file(forbidden, "payroll-secrets.csv", "ssn\n123\n")
    _file(forbidden, "nested/more-secrets.csv", "ssn\n456\n")

    touched: list[str] = []
    real_read = workdir._read_bytes
    monkeypatch.setattr(
        workdir, "_read_bytes", lambda path: (touched.append(str(path)), real_read(path))[1]
    )

    result = _scan(authorized)

    assert [item.relative_path for item in result.files] == ["hours.csv"]
    assert all(str(forbidden) not in path for path in touched)
    assert all(str(forbidden) not in path for path in result.accessed_paths)
    assert all(str(forbidden) not in item.path for item in result.files)

    scope = ScanScope.of(authorized)
    assert scope.authorized_root(forbidden / "payroll-secrets.csv") is None
    with pytest.raises(UnauthorizedPathError):
        scope.require(forbidden / "payroll-secrets.csv")
    with pytest.raises(UnauthorizedPathError):
        scope.require(authorized / ".." / "forbidden" / "payroll-secrets.csv")
    # Widening the allowlist is an explicit act, never inferred from the tree.
    widened = ScanScope.of([authorized, forbidden])
    assert widened.authorized_root(forbidden / "payroll-secrets.csv") == forbidden


def test_non_recursive_scope_reports_the_subdirectory_instead_of_reading_it(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "hours.csv")
    _file(work, "archive/old-hours.csv", "a,b\n7,7\n", when=BEFORE_PERIOD)

    result = _scan(work, recursive=False)

    assert [item.relative_path for item in result.files] == ["hours.csv"]
    assert any(item["reason"] == "subdirectory_not_scanned" for item in result.skipped_paths)

    recursive = _scan(work)
    assert recursive.entry("archive/old-hours.csv") is not None


def test_windows_style_paths_and_chinese_filenames_round_trip(tmp_path):
    work = tmp_path / "工作目录"
    work.mkdir()
    _file(work, "九月工时.csv", "地点,日期\n北区,2026-09-02\n")
    _file(work, "八月工时.csv", "地点,日期\n北区,2026-08-02\n", when=BEFORE_PERIOD)
    _file(work, "模板-北区.xlsx", "template")
    _file(work, "子目录\\九月补充.csv".replace("\\", "/"), "地点,日期\n南区,2026-09-05\n")

    windows_style = str(work).replace("/", "\\")
    result = _scan(windows_style)

    assert result.entry("九月工时.csv").disposition == INCLUDE
    assert result.entry("八月工时.csv").disposition == EXCLUDE
    assert result.entry("模板-北区.xlsx").classification == TEMPLATE
    nested = result.entry("子目录/九月补充.csv")
    assert nested is not None and nested.disposition == INCLUDE
    # Relative paths stay POSIX-separated so a Recipe is portable.
    assert all("\\" not in item.relative_path for item in result.files)
    payload = json.loads(json.dumps(result.to_dict(), ensure_ascii=False))
    assert "九月工时.csv" in {item["relative_path"] for item in payload["files"]}


def test_override_is_applied_and_saved_as_a_reusable_recipe(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "北区月报.xlsx", "looks-like-nothing-in-particular")
    recipe = tmp_path / "recipes" / "workdir.json"

    first = _scan(work)
    automatic = first.entry("北区月报.xlsx")
    assert automatic.classification == INPUT

    override = override_from_entry(
        automatic, TEMPLATE, disposition=EXCLUDE, note="this is our regional template"
    )
    saved = save_workdir_recipe([override], recipe)
    assert saved.is_file()

    reloaded = load_workdir_recipe(recipe)
    assert reloaded["北区月报.xlsx"].classification == TEMPLATE
    assert reloaded["北区月报.xlsx"].note == "this is our regional template"

    second = _scan(work, recipe_path=recipe)
    overridden = second.entry("北区月报.xlsx")
    assert overridden.classification == TEMPLATE
    assert overridden.disposition == EXCLUDE
    assert overridden.overridden_from == INPUT
    assert overridden.signal == "override"
    assert "this is our regional template" in overridden.notes


def test_a_glob_override_cannot_promote_an_unstable_or_artifact_file(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "mid-write.csv", "a,b\n", when=NOW.timestamp() - 1)
    _file(work, "plan (conflicted copy from LAPTOP-7).xlsx", "conflict")

    result = _scan(
        work,
        overrides=[
            ClassificationOverride(path="*.csv", classification=INPUT, disposition=INCLUDE),
            ClassificationOverride(path="*.xlsx", classification=INPUT, disposition=INCLUDE),
        ],
    )

    unstable = result.entry("mid-write.csv")
    assert unstable.disposition == REVIEW
    assert unstable.reason == "override_refused_unstable_file"
    artifact = result.entry("plan (conflicted copy from LAPTOP-7).xlsx")
    assert artifact.disposition == REVIEW
    assert artifact.reason == "override_refused_sync_artifact"
    assert result.included == ()


def test_recipe_format_and_override_values_are_validated(tmp_path):
    bad = tmp_path / "other.json"
    bad.write_text(json.dumps({"format": "excel-ops-recipe-v1", "decisions": []}), encoding="utf-8")
    with pytest.raises(WorkdirScanError):
        load_workdir_recipe(bad)
    with pytest.raises(WorkdirScanError):
        ClassificationOverride(path="hours.csv", classification="not-a-class")
    with pytest.raises(WorkdirScanError):
        ClassificationOverride(path="hours.csv", classification=INPUT, disposition="maybe")
    with pytest.raises(WorkdirScanError):
        ScanScope.of([])
    with pytest.raises(WorkdirScanError):
        ScanScope.of(tmp_path / "missing")
    with pytest.raises(WorkdirScanError):
        PeriodWindow.of(date(2026, 9, 30), date(2026, 9, 1))


def test_without_a_declared_period_data_files_go_to_review_not_input(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "hours.csv")

    result = scan_workdir(work, now=NOW, stability_window_seconds=5.0)

    entry = result.entry("hours.csv")
    assert (entry.classification, entry.disposition) == (UNKNOWN, REVIEW)
    assert entry.reason == "no_period_window_declared"
    assert result.included == ()


def test_scan_is_read_only_and_leaves_the_directory_untouched(tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "hours.csv")
    _file(work, "august-hours.csv", "a,b\n9,9\n", when=BEFORE_PERIOD)
    before = {
        path.name: (path.stat().st_size, path.stat().st_mtime) for path in sorted(work.iterdir())
    }

    _scan(work)

    after = {
        path.name: (path.stat().st_size, path.stat().st_mtime) for path in sorted(work.iterdir())
    }
    assert before == after


def test_cli_scan_workdir_reports_inclusions_exclusions_and_reasons(tmp_path, capsys):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "september-hours.csv")
    _file(work, "august-hours.csv", "a,b\n9,9\n", when=BEFORE_PERIOD)
    _file(work, "final.xlsx", "one")
    _file(work, "final (1).xlsx", "two")
    report = tmp_path / "scan.json"

    main(
        [
            "scan-workdir",
            str(work),
            "--period-start",
            "2026-09-01",
            "--period-end",
            "2026-09-30",
            "--stability-window",
            "0",
            "--result",
            str(report),
        ]
    )

    printed = capsys.readouterr().out
    assert "september-hours.csv -> input (current_period_input" in printed
    assert "historical_outside_current_period" in printed
    assert "no version selected" in printed

    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["format"] == "excel-ops-workdir-scan-v1"
    assert payload["counts"]["include"] == 1
    assert payload["version_groups"][0]["selected"] is None


def test_cli_override_can_be_saved_as_a_recipe_and_reused(tmp_path, capsys):
    work = tmp_path / "work"
    work.mkdir()
    _file(work, "monthly.xlsx", "bytes")
    recipe = tmp_path / "workdir-recipe.json"

    main(
        [
            "scan-workdir",
            str(work),
            "--period-start",
            "2026-09-01",
            "--period-end",
            "2026-09-30",
            "--stability-window",
            "0",
            "--override",
            "monthly.xlsx=template:exclude",
            "--save-recipe",
            str(recipe),
        ]
    )
    capsys.readouterr()
    assert load_workdir_recipe(recipe)["monthly.xlsx"].classification == TEMPLATE

    main(
        [
            "scan-workdir",
            str(work),
            "--period-start",
            "2026-09-01",
            "--period-end",
            "2026-09-30",
            "--stability-window",
            "0",
            "--recipe",
            str(recipe),
            "--json",
        ]
    )
    payload = json.loads(capsys.readouterr().out)
    entry = payload["files"][0]
    assert entry["classification"] == TEMPLATE
    assert entry["disposition"] == EXCLUDE
    assert entry["overridden_from"] == INPUT
