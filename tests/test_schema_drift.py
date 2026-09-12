import json
from pathlib import Path

from openpyxl import Workbook

from excel_ops.cli import main
from excel_ops.schema_drift import compare_workbooks, save_confirmed_mapping, write_drift_report


def _workbook(path: Path, headers: list[str], rows: list[list[object]]) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Employees"
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()
    return path


def test_three_workbooks_separate_rename_from_grain_change(tmp_path):
    baseline = _workbook(tmp_path / "baseline.xlsx", ["Employee ID", "Name"], [[1, "Ana"], [2, "Bo"]])
    renamed = _workbook(tmp_path / "renamed.xlsx", ["Staff Number", "Name"], [[1, "Ana"], [2, "Bo"]])
    repeated = _workbook(tmp_path / "repeated.xlsx", ["Employee ID", "Name"], [[1, "Ana"], [1, "Ana"], [2, "Bo"]])

    report = compare_workbooks([baseline, renamed, repeated], key_field="Employee ID")

    first = report["comparisons"][0]
    assert {change["kind"] for change in first["changes"]} == {"field_added", "field_removed"}
    assert first["mapping_suggestions"][0]["status"] == "review"
    assert first["impact"]["append"] == "blocked"
    second = report["comparisons"][1]
    assert "grain_changed" in {change["kind"] for change in second["changes"]}


def test_confirmed_mapping_is_reused_but_unknown_field_reopens_review(tmp_path):
    baseline = _workbook(tmp_path / "baseline.xlsx", ["Employee ID", "Name"], [[1, "Ana"]])
    renamed = _workbook(tmp_path / "renamed.xlsx", ["Staff Number", "Name"], [[1, "Ana"]])
    unknown = _workbook(tmp_path / "unknown.xlsx", ["Staff Number", "Name", "Cost Centre"], [[1, "Ana", "A"]])
    mapping = tmp_path / "mapping.json"
    mapping.write_text(json.dumps({"confirmed_mappings": {"Staff Number": "Employee ID"}}), encoding="utf-8")

    report = compare_workbooks([baseline, renamed, unknown], mapping_path=mapping)

    assert report["comparisons"][0]["changes"][0]["kind"] == "field_renamed"
    assert report["comparisons"][0]["impact"]["requires_review"] is False
    assert any(change["kind"] == "field_added" for change in report["comparisons"][1]["changes"])
    assert report["comparisons"][1]["impact"]["requires_review"] is True


def test_human_confirmation_is_saved_and_enum_change_is_reported(tmp_path):
    mapping = tmp_path / "mapping.json"
    save_confirmed_mapping(mapping, "Staff Number", "Employee ID")
    assert json.loads(mapping.read_text(encoding="utf-8"))["confirmed_mappings"]["Staff Number"] == "Employee ID"

    baseline = _workbook(tmp_path / "baseline.xlsx", ["ID", "Status"], [[1, "Active"], [2, "Inactive"]])
    changed = _workbook(tmp_path / "changed.xlsx", ["ID", "Status"], [[1, "Active"], [2, "Leave"]])
    report = compare_workbooks([baseline, changed])
    assert "field_enum_changed" in {item["kind"] for item in report["comparisons"][0]["changes"]}


def test_report_is_machine_readable(tmp_path):
    first = _workbook(tmp_path / "first.xlsx", ["ID"], [[1]])
    second = _workbook(tmp_path / "second.xlsx", ["ID"], [[1]])
    output = tmp_path / "drift.json"

    write_drift_report([first, second], output)

    assert json.loads(output.read_text(encoding="utf-8"))["report_version"] == 1


def test_cli_keeps_legacy_pipeline_and_adds_schema_drift(tmp_path):
    delivery = tmp_path / "delivery.xlsx"
    main(["examples/extracted-records.json", str(delivery)])
    assert delivery.exists()

    first = _workbook(tmp_path / "first.xlsx", ["ID"], [[1]])
    second = _workbook(tmp_path / "second.xlsx", ["ID"], [[1]])
    report = tmp_path / "report.json"
    main(["schema-drift", str(first), str(second), "--output", str(report)])
    assert report.exists()
