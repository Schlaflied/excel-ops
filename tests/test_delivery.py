"""End-to-end delivery tests (issue #46).

Every fixture in this module is fabricated. No customer, employee, payroll, or
real enterprise template data appears here.
"""

from __future__ import annotations

import json
import os
from dataclasses import replace
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from excel_ops.ambiguity import decide, group_ambiguities, load_project_recipe, save_project_recipe
from excel_ops.delivery import (
    ACCEPTED,
    REJECTED,
    REVIEW,
    SKIPPED_EXISTING,
    WRITTEN,
    DeliveryPlanError,
    DeliveryTarget,
    _ambiguities_from_matches,
    load_delivery_targets,
    plan_delivery,
    run_delivery,
)
from excel_ops.delivery_manifest import load_delivery_manifest
from excel_ops.delivery_verification import PeriodExpectation
from excel_ops.formula_verification import FormulaRegion, FormulaVerifier
from excel_ops.idempotency import (
    CHANGED,
    FAILED,
    NO_OP,
    RETRY,
    SUCCEEDED,
    IdempotencyOptions,
    file_content_digest,
    load_run_record,
    state_path,
)
from excel_ops.matching import Destination, MatchCandidate, MatchResult
from excel_ops.models import ExtractedRecord
from excel_ops.template_writer import TemplateMapping


NORTH = "北区仓库"
MAPLE = "Maple Street Depot"
PERIOD_BANNER = "2026-09-07 至 2026-09-13"

TABLE_FIELDS = ("record_id", "location", "event_date", "identifier", "category", "source")


# --------------------------------------------------------------------------- #
# Synthetic inputs: two differently shaped layouts plus an extraction JSON
# --------------------------------------------------------------------------- #


def _layout_one(path: Path) -> None:
    """Site/Inspection Date/Asset ID headers on row 3 of a named worksheet."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "巡检记录"
    sheet["A1"] = "synthetic weekly export (fabricated)"
    for column, header in enumerate(
        ("Site", "Inspection Date", "Asset ID", "Inspection Type", "Confidence"), start=1
    ):
        sheet.cell(3, column).value = header
    sheet.append([])
    rows = (
        (NORTH, date(2026, 9, 8), "AS-001", "routine", 0.97),
        ("Maple St. Depot", date(2026, 9, 9), "AS-002", "routine", 0.96),
    )
    for offset, row in enumerate(rows):
        for column, value in enumerate(row, start=1):
            sheet.cell(4 + offset, column).value = value
    workbook.save(path)


def _layout_two(path: Path) -> None:
    """A CSV with a preamble and a different header order and vocabulary."""

    path.write_text(
        "Weekly synthetic export\n"
        "\n"
        "Category,Code,Address,Date,Score\n"
        f"deep clean,AS-003,{MAPLE},2026-09-10,0.95\n"
        f"routine,,{NORTH},2026-09-11,0.99\n",
        encoding="utf-8",
    )


def _extraction_json(path: Path) -> None:
    """A provider-neutral image-extraction fixture with a duplicate region."""

    duplicate = {
        "location": MAPLE,
        "event_date": "2026-09-13",
        "identifier": "AS-006",
        "category": "routine",
        "confidence": 0.9,
    }
    payload = {
        "source": "巡检截图.png",
        "records": [
            {
                "location": NORTH,
                "event_date": "2026-09-12",
                "identifier": "AS-005",
                "category": "routine",
                "confidence": 0.99,
                "source_region": "page-1/box-3",
            },
            duplicate,
            dict(duplicate),
        ],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _inputs(directory: Path) -> list[Path]:
    first = directory / "北区巡检-源数据.xlsx"
    second = directory / "weekly-export.csv"
    third = directory / "巡检截图-extraction.json"
    _layout_one(first)
    _layout_two(second)
    _extraction_json(third)
    return [first, second, third]


# --------------------------------------------------------------------------- #
# Synthetic templates: two different enterprise-style shapes
# --------------------------------------------------------------------------- #


def _north_template(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "北区"
    sheet["A1"] = "合成巡检交付模板"
    sheet.merge_cells("A1:F1")
    sheet["A2"] = "报告周期"
    sheet["B2"] = PERIOD_BANNER
    sheet["G2"] = "未授权区域 - 不得修改"
    for column, header in enumerate(
        ("record_id", "location", "event_date", "identifier", "category", "source"), start=1
    ):
        sheet.cell(4, column).value = header
        sheet.cell(4, column).font = Font(bold=True)
    sheet["A5"].fill = PatternFill("solid", fgColor="FFF2CC")
    sheet["H4"] = "=1+1"
    workbook.save(path)


def _client_template(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Upload"
    for column, header in enumerate(
        ("source", "record_id", "category", "identifier", "event_date", "location"), start=1
    ):
        sheet.cell(1, column).value = header
    sheet["G1"] = "UNAUTHORISED REGION"
    sheet["H1"] = "=3*4"
    workbook.save(path)


def _north_target(template: Path) -> DeliveryTarget:
    return DeliveryTarget(
        destination=Destination(NORTH, (NORTH, "North Warehouse")),
        template_path=template,
        mapping=TemplateMapping(
            sheet="北区",
            header_row=4,
            data_start_row=5,
            field_columns={
                "record_id": "A",
                "location": "B",
                "event_date": "C",
                "identifier": "D",
                "category": "E",
                "source": "F",
            },
        ),
        required_fields=TABLE_FIELDS,
        period_expectations=(PeriodExpectation("北区", "B2", PERIOD_BANNER),),
    )


def _client_target(template: Path) -> DeliveryTarget:
    return DeliveryTarget(
        destination=Destination(MAPLE, (MAPLE,)),
        template_path=template,
        mapping=TemplateMapping(
            sheet="Upload",
            header_row=1,
            data_start_row=2,
            field_columns={
                "source": "A",
                "record_id": "B",
                "category": "C",
                "identifier": "D",
                "event_date": "E",
                "location": "F",
            },
        ),
        required_fields=TABLE_FIELDS,
    )


def _scenario(tmp_path: Path) -> dict[str, object]:
    sources = tmp_path / "sources"
    templates = tmp_path / "templates"
    sources.mkdir()
    templates.mkdir()
    inputs = _inputs(sources)
    north = templates / "巡检模板.xlsx"
    client = templates / "client-upload.xlsx"
    _north_template(north)
    _client_template(client)
    return {
        "inputs": inputs,
        "north": north,
        "client": client,
        "targets": [_north_target(north), _client_target(client)],
        "staging": tmp_path / "staging",
        "delivery": tmp_path / "delivery",
        "before": {path: path.read_bytes() for path in (*inputs, north, client)},
    }


def _run(scenario: dict[str, object], **kwargs):
    return run_delivery(
        scenario["inputs"],
        scenario["targets"],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
        **kwargs,
    )


def _assert_originals_untouched(scenario: dict[str, object]) -> None:
    for path, content in scenario["before"].items():
        assert path.read_bytes() == content, f"original file changed: {path.name}"


# --------------------------------------------------------------------------- #
# Success path
# --------------------------------------------------------------------------- #


def test_plan_is_checkable_before_any_file_is_touched(tmp_path: Path):
    scenario = _scenario(tmp_path)

    plan_run = plan_delivery(
        scenario["inputs"],
        scenario["targets"],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )

    assert plan_run.dry_run is True
    assert plan_run.delivered is False
    assert plan_run.plan.writable is True
    assert plan_run.plan.counts == {
        "input": 7,
        WRITTEN: 0,
        ACCEPTED: 4,
        REVIEW: 2,
        REJECTED: 1,
        SKIPPED_EXISTING: 0,
    }
    planned = {item.destination_key: item for item in plan_run.plan.targets}
    assert planned[NORTH].expected_written == 2
    assert planned[MAPLE].expected_written == 2
    assert planned[MAPLE].expected_review == 1
    assert set(planned[NORTH].field_mapping) == set(TABLE_FIELDS)
    assert plan_run.plan.unresolved_ambiguities  # the fuzzy record is a blocker
    assert not scenario["staging"].exists()
    assert not scenario["delivery"].exists()
    _assert_originals_untouched(scenario)
    json.dumps(plan_run.to_dict())


def test_two_layouts_and_image_json_deliver_into_two_templates(tmp_path: Path):
    scenario = _scenario(tmp_path)

    result = _run(scenario)

    assert result.delivered is True
    assert result.failures == ()
    assert result.counts == {
        "input": 7,
        WRITTEN: 4,
        ACCEPTED: 0,
        REVIEW: 2,
        REJECTED: 1,
        SKIPPED_EXISTING: 0,
    }
    # Every input lands in exactly one terminal state and the counts reconcile.
    assert sum(result.counts[state] for state in (WRITTEN, ACCEPTED, REVIEW, REJECTED, SKIPPED_EXISTING)) == 7
    assert len(result.records) == 7
    # A duplicated extraction region intentionally shares one stable record ID.
    assert len({item.record_id for item in result.records}) == 6

    targets = {item.destination_key: item for item in result.targets}
    assert targets[NORTH].delivered is True
    assert targets[MAPLE].delivered is True
    assert Path(targets[NORTH].delivery_path) != Path(targets[MAPLE].delivery_path)
    assert Path(targets[NORTH].delivery_path).name == "巡检模板-北区仓库.xlsx"

    # Normalized to one record contract regardless of input shape.
    written = result.records_by_status(WRITTEN)
    assert {item.values["identifier"] for item in written} == {"AS-001", "AS-003", "AS-005", "AS-006"}
    assert {item.source.source_file for item in written} == {
        "北区巡检-源数据.xlsx",
        "weekly-export.csv",
        "巡检截图.png",
    }

    _assert_originals_untouched(scenario)
    json.dumps(result.to_dict())


def test_success_path_matches_the_real_persisted_file(tmp_path: Path):
    scenario = _scenario(tmp_path)

    result = _run(scenario)
    north = next(item for item in result.targets if item.destination_key == NORTH)
    delivered = Path(north.delivery_path)

    workbook = load_workbook(delivered, data_only=False)
    try:
        assert workbook.sheetnames == ["北区"]
        sheet = workbook["北区"]
        # Declared period and unauthorized regions survived the write.
        assert sheet["B2"].value == PERIOD_BANNER
        assert sheet["G2"].value == "未授权区域 - 不得修改"
        assert sheet["H4"].value == "=1+1"
        assert sheet["A4"].value == "record_id"
        assert sheet["A4"].font.bold is True

        expected = sorted(item for item in north.written_record_ids)
        actual_ids = sorted(str(sheet.cell(row, 1).value) for row in (5, 6))
        assert actual_ids == expected
        # Required fields, values, and dates match the report, read from the file.
        for row in (5, 6):
            assert sheet.cell(row, 2).value == NORTH
            assert sheet.cell(row, 3).value in {"2026-09-08", "2026-09-12"}
            assert sheet.cell(row, 4).value in {"AS-001", "AS-005"}
            assert sheet.cell(row, 5).value == "routine"
            assert sheet.cell(row, 6).value
        assert sheet.cell(7, 1).value is None
    finally:
        workbook.close()

    report = json.loads(Path(north.verification_report_path).read_text(encoding="utf-8"))
    assert report["passed"] is True
    assert report["delivery_path"] == str(delivered)
    assert report["counts"] == {"input": 2, "accepted": 2, "review": 0, "written": 2}


def test_run_delivery_produces_a_correct_manifest_for_a_real_synthetic_delivery(tmp_path: Path):
    """Issue #20, integration: the manifest ``run_delivery`` wires up end to end.

    Uses the same three-input, two-target, two-layout-plus-extraction-JSON
    scenario as the rest of this module (real xlsx/csv/json inputs, a
    duplicate-region merge, a rejected record) and checks the manifest
    ``run_delivery`` actually attaches -- not a hand-built one -- against the
    real persisted files.
    """

    scenario = _scenario(tmp_path)

    result = _run(scenario)

    assert result.delivered is True
    manifests = {item.destination_key: item for item in result.manifests}
    assert set(manifests) == {NORTH, MAPLE}

    for key, manifest in manifests.items():
        target = next(item for item in result.targets if item.destination_key == key)
        delivered = Path(target.delivery_path)
        assert manifest.output == str(delivered)
        assert manifest.output_name == delivered.name
        assert manifest.output_hash == file_content_digest(delivered)
        assert manifest.verification.status == "passed"
        assert manifest.verification.passed is True
        assert manifest.reconciled is True
        assert manifest.discrepancies == ()

        # Sources: only the inputs that actually contributed to this output,
        # each identified by its real content hash.
        contributing = {item.file: item for item in manifest.sources}
        for path in contributing:
            source_path = next(p for p in scenario["inputs"] if p.name == path)
            assert contributing[path].hash == file_content_digest(source_path)
        assert sum(item.written for item in manifest.sources) == manifest.written

    north = manifests[NORTH]
    maple = manifests[MAPLE]
    assert [tab.sheet for tab in north.tabs] == ["北区"]
    assert [tab.sheet for tab in maple.tabs] == ["Upload"]
    assert north.tabs[0].written == north.tabs[0].rows == 2
    assert maple.tabs[0].written == maple.tabs[0].rows == 2
    # The duplicate extraction region collapsed to one record, and each
    # written record landed on exactly one of the two outputs.
    assert north.written + maple.written == result.counts[WRITTEN] == 4

    # Template/recipe versions are content hashes of what was really used.
    assert north.template_version == file_content_digest(scenario["north"])
    assert maple.template_version == file_content_digest(scenario["client"])
    assert north.recipe_version == maple.recipe_version  # no recipe supplied

    # Period expectations declared on the North target surface in its manifest.
    assert north.period_display_text is None  # no PeriodResult was passed in

    # Both forms were written next to the real delivered files, and reload.
    for manifest in manifests.values():
        assert manifest.manifest_path().parent == Path(manifest.output).parent
        reloaded = load_delivery_manifest(manifest.manifest_path())
        assert reloaded["output_hash"] == manifest.output_hash
        assert reloaded["reconciled"] is True
        assert manifest.readable_path().is_file()

    _assert_originals_untouched(scenario)
    json.dumps(result.to_dict())


def test_every_written_cell_is_traceable_to_source_and_record_id(tmp_path: Path):
    scenario = _scenario(tmp_path)

    result = _run(scenario)

    written = result.records_by_status(WRITTEN)
    assert written
    for record in written:
        assert len(record.cells) == len(TABLE_FIELDS)
        for cell in record.cells:
            assert cell.record_id == record.record_id
            assert cell.source.source_file
            assert cell.cell[0].isalpha() and cell.cell[1:].isdigit()
            # A row from a sheet keeps its row number; an image region keeps its box.
            assert cell.source.source_row is not None or cell.source.source_region or cell.source.locator

    image_record = next(item for item in written if item.values["identifier"] == "AS-005")
    assert image_record.source.source_region == "page-1/box-3"
    sheet_record = next(item for item in written if item.values["identifier"] == "AS-001")
    assert sheet_record.source.source_sheet == "巡检记录"
    assert sheet_record.source.source_row == 4


# --------------------------------------------------------------------------- #
# Accepted / review / rejected routing
# --------------------------------------------------------------------------- #


def test_ambiguous_duplicate_and_invalid_records_never_reach_the_file(tmp_path: Path):
    scenario = _scenario(tmp_path)

    result = _run(scenario)

    review = result.records_by_status(REVIEW)
    rejected = result.records_by_status(REJECTED)
    assert len(review) == 2
    assert len(rejected) == 1
    assert rejected[0].reasons == ("missing:identifier",)
    assert rejected[0].destination_key is None

    fuzzy = next(item for item in review if item.values["location"] == "Maple St. Depot")
    assert fuzzy.match_status == "review"
    assert fuzzy.candidates == (MAPLE,)
    assert "unresolved_ambiguity" in fuzzy.reasons
    assert fuzzy.destination_key is None
    duplicate = next(item for item in review if item.match_status == "duplicate")
    assert "duplicate_record_id" in duplicate.reasons

    written_ids = [value for target in result.targets for value in target.written_record_ids]
    # The duplicate shares its ID with the accepted original, which is written once.
    assert len(written_ids) == len(set(written_ids)) == 4
    assert duplicate.record_id in written_ids
    withheld_ids = {fuzzy.record_id, rejected[0].record_id}
    assert withheld_ids.isdisjoint(written_ids)

    for target in result.targets:
        workbook = load_workbook(target.delivery_path, data_only=True)
        try:
            sheet = workbook[workbook.sheetnames[0]]
            values = [str(cell.value) for row in sheet.iter_rows() for cell in row]
            assert withheld_ids.isdisjoint(values)
            assert "Maple St. Depot" not in values
            assert values.count(duplicate.record_id) <= 1
        finally:
            workbook.close()


def test_conflicting_destinations_are_never_written(tmp_path: Path):
    scenario = _scenario(tmp_path)
    overlapping = DeliveryTarget(
        destination=Destination("shared-alias-target", (NORTH,)),
        template_path=scenario["client"],
        mapping=_client_target(scenario["client"]).mapping,
        required_fields=TABLE_FIELDS,
    )

    result = run_delivery(
        scenario["inputs"],
        [*scenario["targets"], overlapping],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )

    conflicts = [item for item in result.records if item.match_status == "conflict"]
    assert conflicts
    for item in conflicts:
        assert item.status == REVIEW
        assert item.destination_key is None
    assert not any("shared-alias-target" in (target.destination_key or "") and target.delivered
                   for target in result.targets if target.written_record_ids)


def test_project_recipe_resolves_the_ambiguity_and_is_reusable(tmp_path: Path):
    scenario = _scenario(tmp_path)
    recipe = tmp_path / "recipes" / "project-recipe.json"

    first = plan_delivery(
        scenario["inputs"],
        scenario["targets"],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )
    pending = [item for item in first.confirmation_batch.items if item.status == "pending"]
    assert len(pending) == 1
    decision = decide(pending[0].ambiguity, MAPLE, scope="project", source="test-operator")
    save_project_recipe([decision], recipe)
    assert set(load_project_recipe(recipe)) == {decision.key}

    result = _run(scenario, recipe_path=recipe)

    assert result.delivered is True
    assert result.counts[WRITTEN] == 5
    assert result.counts[REVIEW] == 1
    resolved = next(item for item in result.ambiguities if item.status == "resolved")
    assert resolved.selected == MAPLE
    assert resolved.decision_source == "project:test-operator"
    confirmed = next(
        item for item in result.records_by_status(WRITTEN)
        if item.values["location"] == "Maple St. Depot"
    )
    assert confirmed.rule == "human_confirmation"
    _assert_originals_untouched(scenario)


# --------------------------------------------------------------------------- #
# Rerun behaviour (basic case only; full fingerprinting stays with #19)
# --------------------------------------------------------------------------- #


def test_rerunning_the_same_confirmed_input_appends_nothing(tmp_path: Path):
    scenario = _scenario(tmp_path)

    first = _run(scenario)
    assert first.delivered is True
    delivered = {item.destination_key: Path(item.delivery_path) for item in first.targets}
    before = {key: path.read_bytes() for key, path in delivered.items()}

    second = _run(scenario)

    assert second.counts[SKIPPED_EXISTING] == 4
    assert second.counts[WRITTEN] == 0
    assert all(item.status == "no_new_records" for item in second.targets)
    assert second.delivered is True
    for key, path in delivered.items():
        assert path.read_bytes() == before[key]
        workbook = load_workbook(path, data_only=True)
        try:
            sheet = workbook[workbook.sheetnames[0]]
            ids = [cell.value for row in sheet.iter_rows(min_col=1, max_col=6) for cell in row]
            assert len([value for value in ids if str(value).startswith("rec_")]) == 2
        finally:
            workbook.close()
    _assert_originals_untouched(scenario)


# --------------------------------------------------------------------------- #
# Formula checks
# --------------------------------------------------------------------------- #


def _formula_template(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for column, header in enumerate(TABLE_FIELDS, start=1):
        sheet.cell(1, column).value = header
    sheet["G1"] = "id_length"
    sheet["G2"] = "=LEN(A2)"
    sheet["G3"] = "=LEN(A3)"
    workbook.save(path)


def test_declared_formula_regions_are_verified_and_recalculation_is_not_claimed(tmp_path: Path):
    source = tmp_path / "formula-input.json"
    source.write_text(
        json.dumps(
            {
                "source": "synthetic-formula.png",
                "records": [
                    {
                        "location": NORTH,
                        "event_date": "2026-09-12",
                        "identifier": "AS-101",
                        "category": "routine",
                        "confidence": 0.99,
                        "source_region": "box-1",
                    },
                    {
                        "location": NORTH,
                        "event_date": "2026-09-13",
                        "identifier": "AS-102",
                        "category": "routine",
                        "confidence": 0.99,
                        "source_region": "box-2",
                    },
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    template = tmp_path / "formula-template.xlsx"
    _formula_template(template)
    before = template.read_bytes()
    target = DeliveryTarget(
        destination=Destination(NORTH, (NORTH,)),
        template_path=template,
        mapping=TemplateMapping(
            sheet="Data",
            header_row=1,
            data_start_row=2,
            field_columns={name: index for index, name in enumerate(TABLE_FIELDS, start=1)},
        ),
        required_fields=TABLE_FIELDS,
        formula_verifier=FormulaVerifier(formula_regions=(FormulaRegion("Data", "G2:G3"),)),
    )

    result = run_delivery(
        [source],
        [target],
        staging_dir=tmp_path / "staging",
        delivery_dir=tmp_path / "delivery",
    )

    assert result.delivered is True
    findings = {item.code for item in result.targets[0].findings}
    assert findings == {"recalculation_not_verified"}
    assert all(item.severity == "warning" for item in result.targets[0].findings)
    assert template.read_bytes() == before
    workbook = load_workbook(result.targets[0].delivery_path)
    try:
        assert workbook["Data"]["G3"].value == "=LEN(A3)"
    finally:
        workbook.close()


# --------------------------------------------------------------------------- #
# Failure paths: nothing that looks delivered may be emitted
# --------------------------------------------------------------------------- #


def test_unknown_layout_stops_the_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    unknown = tmp_path / "sources" / "unknown-layout.csv"
    unknown.write_text("alpha,beta,gamma\n1,2,3\n", encoding="utf-8")

    result = run_delivery(
        [*scenario["inputs"], unknown],
        scenario["targets"],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )

    assert result.delivered is False
    assert [item.code for item in result.failures] == ["unknown_layout"]
    assert not scenario["delivery"].exists()
    assert not scenario["staging"].exists()
    _assert_originals_untouched(scenario)


def test_mapping_an_unknown_field_blocks_the_plan_before_writing(tmp_path: Path):
    scenario = _scenario(tmp_path)
    broken = DeliveryTarget(
        destination=Destination(NORTH, (NORTH,)),
        template_path=scenario["north"],
        mapping=TemplateMapping(
            sheet="北区",
            header_row=4,
            data_start_row=5,
            field_columns={"record_id": "A", "employee_salary": "B"},
        ),
        required_fields=("record_id", "employee_salary"),
    )

    result = run_delivery(
        scenario["inputs"],
        [broken],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )

    assert result.delivered is False
    assert result.plan.writable is False
    assert any("unmapped_contract_field:employee_salary" in item for item in result.plan.blocking_items)
    assert [item.code for item in result.failures] == ["plan_blocked"]
    assert not scenario["delivery"].exists()
    _assert_originals_untouched(scenario)


def test_mapping_a_missing_sheet_fails_without_delivering(tmp_path: Path):
    scenario = _scenario(tmp_path)
    broken = DeliveryTarget(
        destination=Destination(NORTH, (NORTH,)),
        template_path=scenario["north"],
        mapping=TemplateMapping(
            sheet="NoSuchSheet",
            header_row=4,
            data_start_row=5,
            field_columns={"record_id": "A", "location": "B"},
        ),
        required_fields=("record_id", "location"),
    )

    result = run_delivery(
        scenario["inputs"],
        [broken],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )

    assert result.delivered is False
    assert [item.code for item in result.failures] == ["template_write_failed"]
    assert result.delivery_paths == ()
    assert not any(scenario["delivery"].glob("*.xlsx")) if scenario["delivery"].exists() else True
    _assert_originals_untouched(scenario)


def test_a_protected_cell_blocks_delivery_instead_of_writing_part_of_a_row(tmp_path: Path):
    scenario = _scenario(tmp_path)
    protected = scenario["north"]
    workbook = load_workbook(protected)
    try:
        workbook["北区"]["C5"] = "=TODAY()"
        workbook.save(protected)
    finally:
        workbook.close()
    scenario["before"][protected] = protected.read_bytes()

    result = _run(scenario)

    north = next(item for item in result.targets if item.destination_key == NORTH)
    assert result.delivered is False
    assert north.delivered is False
    assert north.status == "incomplete_write"
    assert north.delivery_path is None
    assert any(item["reason"] == "protected_formula" for item in north.skipped_writes)
    assert [item.code for item in result.failures] == ["incomplete_write"]
    assert not (scenario["delivery"] / "巡检模板-北区仓库.xlsx").exists()
    _assert_originals_untouched(scenario)


def test_plan_predicts_protected_cells_and_excludes_them_from_expected_written(tmp_path: Path):
    scenario = _scenario(tmp_path)
    protected = scenario["north"]
    workbook = load_workbook(protected)
    try:
        workbook["北区"]["C5"] = "=TODAY()"
        workbook.save(protected)
    finally:
        workbook.close()
    scenario["before"][protected] = protected.read_bytes()

    plan_run = plan_delivery(
        scenario["inputs"],
        scenario["targets"],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )

    planned = {item.destination_key: item for item in plan_run.plan.targets}
    north = planned[NORTH]
    # Without the fix, expected_written stayed at 2 and protected_cells stayed
    # empty because `_predicted_protected_cells` was never called.
    assert north.protected_cells == ("C5",)
    assert north.expected_skipped_protected == 1
    assert north.expected_written == 1
    _assert_originals_untouched(scenario)


def test_planned_target_blocking_items_are_populated_when_blockers_exist(tmp_path: Path):
    scenario = _scenario(tmp_path)
    broken = DeliveryTarget(
        destination=Destination(NORTH, (NORTH,)),
        template_path=scenario["north"],
        mapping=TemplateMapping(
            sheet="北区",
            header_row=4,
            data_start_row=5,
            field_columns={"record_id": "A", "employee_salary": "B"},
        ),
        required_fields=("record_id", "employee_salary"),
    )

    plan_run = plan_delivery(
        scenario["inputs"],
        [broken],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )

    north = next(item for item in plan_run.plan.targets if item.destination_key == NORTH)
    # Without the fix, `tuple(blockers)` was positionally assigned to
    # `expected_skipped_protected`, leaving `blocking_items` empty here even
    # though real blockers exist for this target.
    assert north.blocking_items != ()
    assert any("unmapped_contract_field:employee_salary" in item for item in north.blocking_items)
    assert isinstance(north.expected_skipped_protected, int)
    _assert_originals_untouched(scenario)


def test_a_corrupted_staging_file_is_never_delivered(tmp_path: Path):
    scenario = _scenario(tmp_path)

    def corrupt(staged: Path) -> None:
        staged.write_bytes(b"this is not a workbook")

    result = _run(scenario, post_stage_hook=corrupt)

    assert result.delivered is False
    codes = {item.code for item in result.failures}
    assert codes == {"verification_failed"}
    for target in result.targets:
        assert target.delivered is False
        assert target.delivery_path is None
        assert "unreadable_workbook" in {item.code for item in target.findings}
    assert not scenario["delivery"].exists() or not any(scenario["delivery"].glob("*.xlsx"))
    _assert_originals_untouched(scenario)


def test_failed_verification_withholds_the_delivery(tmp_path: Path):
    scenario = _scenario(tmp_path)
    wrong_period = DeliveryTarget(
        destination=Destination(NORTH, (NORTH, "North Warehouse")),
        template_path=scenario["north"],
        mapping=_north_target(scenario["north"]).mapping,
        required_fields=TABLE_FIELDS,
        period_expectations=(PeriodExpectation("北区", "B2", "2026-10-05 至 2026-10-11"),),
    )

    result = run_delivery(
        scenario["inputs"],
        [wrong_period],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )

    north = result.targets[0]
    assert result.delivered is False
    assert north.status == "verification_failed"
    assert north.delivery_path is None
    assert "period_mismatch" in {item.code for item in north.findings}
    report = json.loads(Path(north.verification_report_path).read_text(encoding="utf-8"))
    assert report["passed"] is False
    assert report["delivery_path"] is None
    assert not scenario["delivery"].exists() or not any(scenario["delivery"].glob("*.xlsx"))
    _assert_originals_untouched(scenario)


def test_duplicate_target_keys_are_rejected(tmp_path: Path):
    scenario = _scenario(tmp_path)

    with pytest.raises(DeliveryPlanError):
        run_delivery(
            scenario["inputs"],
            [scenario["targets"][0], scenario["targets"][0]],
            staging_dir=scenario["staging"],
            delivery_dir=scenario["delivery"],
        )


# --------------------------------------------------------------------------- #
# Non-ASCII names and native (Windows) path separators
# --------------------------------------------------------------------------- #


def test_non_ascii_names_and_native_path_separators(tmp_path: Path):
    root = tmp_path / "交付 输出" / "巡检 2026"
    root.mkdir(parents=True)
    scenario = _scenario(tmp_path)
    staging = str(root / "暂存")
    delivery = str(root / "交付件")
    if os.name == "nt":
        assert "\\" in staging and ":" in staging

    result = run_delivery(
        [str(path) for path in scenario["inputs"]],
        scenario["targets"],
        staging_dir=staging,
        delivery_dir=delivery,
    )

    assert result.delivered is True
    delivered = Path(next(item for item in result.targets if item.destination_key == NORTH).delivery_path)
    assert delivered.parent == Path(delivery).resolve()
    assert delivered.name == "巡检模板-北区仓库.xlsx"
    assert delivered.is_file()
    workbook = load_workbook(delivered, data_only=True)
    try:
        assert workbook["北区"]["B5"].value == NORTH
    finally:
        workbook.close()
    _assert_originals_untouched(scenario)


# --------------------------------------------------------------------------- #
# Declarative configuration used by the CLI and the Agent Skill
# --------------------------------------------------------------------------- #


def test_declarative_configuration_round_trips_through_the_cli(tmp_path: Path):
    from excel_ops.cli import main

    scenario = _scenario(tmp_path)
    template = Path(scenario["targets"][0].template_path)
    workbook = load_workbook(template)
    try:
        for worksheet in workbook.worksheets:
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        cell.value = None
        workbook.save(template)
    finally:
        workbook.close()
    scenario["before"][template] = template.read_bytes()
    config = tmp_path / "delivery-plan.json"
    config.write_text(
        json.dumps(
            {
                "inputs": [
                    "sources/北区巡检-源数据.xlsx",
                    "sources/weekly-export.csv",
                    "sources/巡检截图-extraction.json",
                ],
                "staging_dir": "staging",
                "delivery_dir": "delivery",
                "delivery": {
                    "formats": ["xlsx", "csv"],
                    "csv": {"mode": "one-file-per-sheet"},
                },
                "targets": [
                    {
                        "key": NORTH,
                        "aliases": [NORTH, "North Warehouse"],
                        "template": "templates/巡检模板.xlsx",
                        "sheet": "北区",
                        "header_row": 4,
                        "data_start_row": 5,
                        "field_columns": {
                            "record_id": "A",
                            "location": "B",
                            "event_date": "C",
                            "identifier": "D",
                            "category": "E",
                            "source": "F",
                        },
                        "required_fields": list(TABLE_FIELDS),
                        "period_expectations": [
                            {"sheet": "北区", "cell": "B2", "expected": PERIOD_BANNER}
                        ],
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    targets, options = load_delivery_targets(
        json.loads(config.read_text(encoding="utf-8")), base_dir=config.parent
    )
    assert [item.key for item in targets] == [NORTH]
    assert options["delivery_dir"] == tmp_path / "delivery"

    result_path = tmp_path / "run.json"
    main(["deliver", str(config), "--dry-run", "--result", str(result_path)])
    dry = json.loads(result_path.read_text(encoding="utf-8"))
    assert dry["dry_run"] is True
    assert dry["plan"]["writable"] is True
    assert not (tmp_path / "delivery").exists()

    main(["deliver", str(config), "--result", str(result_path)])
    report = json.loads(result_path.read_text(encoding="utf-8"))
    assert report["delivered"] is True
    assert report["counts"][WRITTEN] == 2
    assert Path(report["delivery_paths"][0]).is_file()
    assert {item["actual_format"] for item in report["exports"]} == {"xlsx", "csv"}
    manifest = json.loads(Path(report["manifest_paths"][0]).read_text(encoding="utf-8"))
    assert {item["actual_format"] for item in manifest["exports"]} == {"xlsx", "csv"}
    readable = Path(report["manifest_paths"][0]).with_suffix(".txt").read_text(encoding="utf-8")
    assert "digest=" in readable
    assert "verification=passed" in readable
    assert "summary=verified output" in readable
    _assert_originals_untouched(scenario)


# --------------------------------------------------------------------------- #
# Finding 7: ambiguity keys must be stable across records sharing a shape
# --------------------------------------------------------------------------- #


def _fuzzy_record(location: str, identifier: str) -> ExtractedRecord:
    return ExtractedRecord(
        location=location,
        event_date="2026-09-10",
        identifier=identifier,
        category="routine",
        confidence=0.9,
        source="synthetic",
    )


def _review_match(location: str, identifier: str) -> MatchResult:
    """A 'review' match whose only difference from another is the source record.

    Both records are fuzzily proposed for the same single candidate, so they
    describe the *same kind* of open question ("is this a match for North?"),
    even though each comes from a different source location. A stable
    ambiguity key must merge them into one confirmation item instead of
    minting a fresh key per record.
    """

    record = _fuzzy_record(location, identifier)
    return MatchResult(
        record_id=f"rec_{identifier}",
        record=record,
        status="review",
        destination_key=None,
        rule="fuzzy",
        confidence=0.8,
        candidates=(MatchCandidate(NORTH, NORTH, "fuzzy", 0.8),),
    )


def test_ambiguity_key_merges_records_sharing_the_same_shape():
    matches = [
        _review_match("北区仓库附近", "AS-101"),
        _review_match("北区仓庫 (typo)", "AS-102"),
        _review_match("Bei Qu Warehouse", "AS-103"),
    ]

    items, records_by_key = _ambiguities_from_matches(matches)

    # _ambiguities_from_matches emits one Ambiguity per match (affected_count=1
    # each); the merging happens downstream in group_ambiguities. What must be
    # stable here is the *key* -- it must be identical across all three
    # records' matches even though each record has its own location/identifier.
    assert len(items) == 3
    for item in items:
        assert item.affected_count == 1  # per-match affected_count before grouping
        assert item.candidates == (NORTH,)
    keys = {item.key for item in items}
    assert len(keys) == 1, f"expected one stable key across records, got {keys}"

    # group_ambiguities (as build_confirmation_batch calls it) must then merge
    # all three per-match ambiguities into a single confirmation item whose
    # affected_count reflects all three records, instead of leaving three
    # separate items that each demand their own confirmation.
    merged = group_ambiguities(items)
    assert len(merged) == 1
    assert merged[0].affected_count == 3
    assert set(records_by_key) == keys


# --------------------------------------------------------------------------- #
# Whole-run idempotency wiring (issue #19)
#
# These exercise the run-level short-circuit ABOVE the per-record deduplication
# covered by test_rerunning_the_same_confirmed_input_appends_nothing: that one
# still re-ingests, re-matches and re-plans before finding nothing to append,
# while an idempotent rerun never gets that far.
# --------------------------------------------------------------------------- #


TASK_KEY = "周报-北区"


def _idempotent() -> IdempotencyOptions:
    return IdempotencyOptions(task_key=TASK_KEY, template_profile_version="profile-v1")


def test_an_idempotent_rerun_short_circuits_to_a_no_op(tmp_path: Path, monkeypatch):
    scenario = _scenario(tmp_path)
    options = _idempotent()

    first = _run(scenario, idempotency=options)
    assert first.delivered is True
    assert first.no_op is False
    assert first.run_decision.decision == CHANGED
    assert first.run_decision.reason == "no_previous_run"
    delivered = {item.destination_key: Path(item.delivery_path) for item in first.targets}
    before = {key: path.read_bytes() for key, path in delivered.items()}

    # Nothing may be re-ingested, re-matched, re-written or re-verified: the
    # only correct second run never reaches the pipeline at all.
    def _never(*args, **kwargs):
        raise AssertionError("an idempotent no-op must not redo the delivery work")

    monkeypatch.setattr("excel_ops.delivery._ingest", _never)
    second = _run(scenario, idempotency=options)

    assert second.no_op is True
    assert second.delivered is False
    assert second.failures == ()
    assert second.records == ()
    assert second.counts[WRITTEN] == 0
    assert {item.status for item in second.targets} == {NO_OP}
    assert second.run_decision.decision == NO_OP
    # The recorded fingerprint is the one recomputed *after* the delivery, so it
    # already accounts for the output the first run produced; otherwise the
    # output content hash could never match on a rerun.
    baseline = load_run_record(state_path(scenario["delivery"]), TASK_KEY)
    assert baseline.status == SUCCEEDED
    assert second.fingerprint == baseline.fingerprint
    assert second.fingerprint != first.run_decision.fingerprint.digest
    assert "unchanged_since_successful_run" in second.run_decision.explain()
    for key, path in delivered.items():
        assert path.read_bytes() == before[key], "a no-op must not rewrite the delivery"
    json.dumps(second.to_dict(), default=str)
    _assert_originals_untouched(scenario)


def test_a_missing_recorded_export_bypasses_no_op_and_is_regenerated(tmp_path: Path):
    scenario = _scenario(tmp_path)
    options = _idempotent()

    def attach_csv(run, manifests):
        recovered = []
        for manifest in manifests:
            sibling = Path(manifest.output).with_suffix(".csv")
            sibling.write_text("record_id\nsynthetic\n", encoding="utf-8")
            evidence = {
                "actual_format": "csv",
                "output": str(sibling),
                "sheets": [scenario["targets"][0].mapping.sheet],
                "warnings": ["csv_drops_workbook_features"],
                "digest": file_content_digest(sibling),
                "verification": "passed",
                "summary": "verified output",
            }
            recovered.append(replace(manifest, exports=(evidence,)))
        return tuple(recovered)

    first = _run(scenario, idempotency=options, artifact_hook=attach_csv)
    assert first.delivered is True
    missing = Path(first.manifests[0].exports[0]["output"])
    missing.unlink()

    retried = _run(scenario, idempotency=options, artifact_hook=attach_csv)

    assert retried.no_op is False
    assert retried.run_decision.decision == RETRY
    assert retried.run_decision.reason == "recorded_export_missing_or_changed"
    assert retried.manifest_paths
    assert missing.is_file()
    assert retried.manifests[0].exports
    assert _run(scenario, idempotency=options, artifact_hook=attach_csv).no_op is True


def test_a_changed_export_selection_bypasses_no_op(tmp_path: Path):
    scenario = _scenario(tmp_path)
    options = _idempotent()
    xlsx = {"formats": ["xlsx"], "selection": {}}
    csv = {
        "formats": ["xlsx", "csv"],
        "selection": {"csv": {"mode": "one-file-per-sheet"}},
    }

    first = _run(scenario, idempotency=options, artifact_fingerprint=xlsx)
    assert first.delivered is True
    assert _run(scenario, idempotency=options, artifact_fingerprint=xlsx).no_op is True

    changed = _run(scenario, idempotency=options, artifact_fingerprint=csv)

    assert changed.no_op is False
    assert changed.run_decision.decision == CHANGED
    assert "mapping" in changed.run_decision.changed_components


def test_cli_rejects_run_state_without_manifest():
    from excel_ops.cli import main

    with pytest.raises(SystemExit) as error:
        main(
            [
                "deliver",
                "unused.json",
                "--run-state",
                "state.json",
                "--no-manifest",
            ]
        )

    assert error.value.code == 2


def test_a_changed_input_does_not_short_circuit(tmp_path: Path):
    scenario = _scenario(tmp_path)
    options = _idempotent()
    assert _run(scenario, idempotency=options).delivered is True

    extra = Path(scenario["inputs"][1])
    extra.write_text(
        extra.read_text(encoding="utf-8") + f"routine,AS-007,{NORTH},2026-09-12,0.98\n",
        encoding="utf-8",
    )
    second = _run(scenario, idempotency=options)

    assert second.no_op is False
    assert second.run_decision.decision == CHANGED
    assert "inputs" in second.run_decision.changed_components
    assert second.counts[WRITTEN] == 1, "only the new record is appended"


def test_a_renamed_but_identical_input_still_short_circuits(tmp_path: Path):
    scenario = _scenario(tmp_path)
    options = _idempotent()
    assert _run(scenario, idempotency=options).delivered is True

    original = Path(scenario["inputs"][1])
    renamed = original.with_name("第38周-导出（副本）.csv")
    renamed.write_bytes(original.read_bytes())
    original.unlink()
    scenario["inputs"] = [scenario["inputs"][0], renamed, scenario["inputs"][2]]
    scenario["before"].pop(original)

    second = _run(scenario, idempotency=options)

    assert second.no_op is True


def test_a_changed_template_does_not_short_circuit(tmp_path: Path):
    scenario = _scenario(tmp_path)
    options = _idempotent()
    assert _run(scenario, idempotency=options).delivered is True

    template = Path(scenario["north"])
    workbook = load_workbook(template)
    try:
        workbook["北区"]["A1"] = "合成巡检交付模板 v2"
        workbook.save(template)
    finally:
        workbook.close()
    scenario["before"][template] = template.read_bytes()

    second = _run(scenario, idempotency=options)

    assert second.no_op is False
    assert "templates" in second.run_decision.changed_components


def test_a_changed_review_decision_does_not_short_circuit(tmp_path: Path):
    scenario = _scenario(tmp_path)
    options = _idempotent()
    assert _run(scenario, idempotency=options).delivered is True
    assert _run(scenario, idempotency=options).no_op is True

    pending = plan_delivery(
        scenario["inputs"],
        scenario["targets"],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    ).confirmation_batch
    question = next(item.ambiguity for item in pending.items)
    confirmed = decide(question, MAPLE, scope="this-run")

    third = _run(scenario, idempotency=options, decisions=[confirmed])

    assert third.no_op is False, "a new human decision must never be a no-op"
    assert "confirmations" in third.run_decision.changed_components


def test_a_this_run_decision_changes_the_recorded_recipe_version(tmp_path: Path):
    """A ``this-run``-scoped decision can change what gets delivered, so it must
    never be excluded from ``recipe_version`` -- otherwise two runs delivered
    under genuinely different effective rules could share the same version."""

    from excel_ops.ambiguity import RecipeDecision

    def _this_run_decision(selected: str) -> RecipeDecision:
        return RecipeDecision(
            field="synthetic-field",
            question="which value applies to this run only?",
            selected=selected,
            scope="this-run",
            source="test",
            decided_at="2026-09-14T00:00:00+00:00",
            candidates=("alpha", "beta"),
        )

    dir_a = tmp_path / "a"
    dir_a.mkdir()
    dir_b = tmp_path / "b"
    dir_b.mkdir()
    scenario_a = _scenario(dir_a)
    scenario_b = _scenario(dir_b)

    run_a = _run(scenario_a, decisions=[_this_run_decision("alpha")])
    run_b = _run(scenario_b, decisions=[_this_run_decision("beta")])

    assert run_a.delivered is True
    assert run_b.delivered is True
    assert run_a.manifests and run_b.manifests

    manifests_a = {item.destination_key: item for item in run_a.manifests}
    manifests_b = {item.destination_key: item for item in run_b.manifests}
    assert set(manifests_a) == set(manifests_b)
    for key in manifests_a:
        assert manifests_a[key].recipe_version != manifests_b[key].recipe_version


def test_a_blocked_run_is_never_recorded_as_a_successful_no_op_baseline(tmp_path: Path):
    scenario = _scenario(tmp_path)
    options = _idempotent()
    broken = DeliveryTarget(
        destination=Destination(NORTH, (NORTH,)),
        template_path=scenario["north"],
        mapping=TemplateMapping(
            sheet="北区",
            header_row=4,
            data_start_row=5,
            field_columns={"record_id": "A", "employee_salary": "B"},
        ),
    )

    def _blocked():
        return run_delivery(
            scenario["inputs"],
            [broken],
            staging_dir=scenario["staging"],
            delivery_dir=scenario["delivery"],
            idempotency=options,
        )

    first = _blocked()
    assert first.delivered is False
    assert [item.code for item in first.failures] == ["plan_blocked"]

    record = load_run_record(state_path(scenario["delivery"]), TASK_KEY)
    assert record is not None
    assert record.status == FAILED
    assert record.successful is False

    second = _blocked()

    assert second.no_op is False, "a blocked run must never become a no-op baseline"
    assert second.run_decision.decision == RETRY
    assert [item.code for item in second.failures] == ["plan_blocked"]
    _assert_originals_untouched(scenario)


def test_a_failed_verification_is_recorded_as_failed_and_the_next_run_retries(tmp_path: Path):
    scenario = _scenario(tmp_path)
    options = _idempotent()

    def corrupt(staged: Path) -> None:
        staged.write_bytes(b"not a workbook")

    failed = _run(scenario, idempotency=options, post_stage_hook=corrupt)
    assert failed.delivered is False
    assert any(item.code == "verification_failed" for item in failed.failures)
    record = load_run_record(state_path(scenario["delivery"]), TASK_KEY)
    assert record.status == FAILED

    retried = _run(scenario, idempotency=options)

    assert retried.no_op is False
    assert retried.run_decision.decision == RETRY
    assert retried.delivered is True
    assert _run(scenario, idempotency=options).no_op is True
    _assert_originals_untouched(scenario)


def test_a_manifest_write_failure_is_recorded_as_failed_and_the_next_run_retries(
    tmp_path: Path, monkeypatch
):
    from excel_ops.delivery_manifest import write_delivery_manifests as _real_write_delivery_manifests

    scenario = _scenario(tmp_path)
    options = _idempotent()

    def _boom(*args, **kwargs):
        raise OSError("disk full")

    monkeypatch.setattr("excel_ops.delivery.write_delivery_manifests", _boom)
    failed = _run(scenario, idempotency=options)

    # (a) the manifest-write failure is reported as a real failure, never a
    # silent success: the workbook was delivered, but the manifest was not.
    assert failed.delivered is False
    assert failed.manifests == ()
    assert any(item.code == "manifest_write_failed" for item in failed.failures)

    # (b) the idempotency record must not be marked SUCCEEDED for this
    # fingerprint -- otherwise a later identical run would see a matching
    # baseline and skip regenerating the manifest forever.
    record = load_run_record(state_path(scenario["delivery"]), TASK_KEY)
    assert record is not None
    assert record.status == FAILED
    assert record.successful is False

    # (c) with the write failure lifted, a subsequent identical run must
    # retry -- never short-circuit to a no-op -- and this time it actually
    # produces the manifest.
    monkeypatch.setattr(
        "excel_ops.delivery.write_delivery_manifests", _real_write_delivery_manifests
    )
    retried = _run(scenario, idempotency=options)

    assert retried.no_op is False
    assert retried.run_decision.decision == RETRY
    assert retried.delivered is True
    # The write failure is gone: nothing about this retry is still broken.
    assert retried.failures == ()
    assert _run(scenario, idempotency=options).no_op is True

    # With genuinely new work to do, a retry regenerates the manifest that the
    # first attempt never got to persist.
    extra = Path(scenario["inputs"][1])
    extra.write_text(
        extra.read_text(encoding="utf-8") + f"routine,AS-900,{NORTH},2026-09-12,0.98\n",
        encoding="utf-8",
    )
    healed = _run(scenario, idempotency=options)
    assert healed.no_op is False
    assert healed.delivered is True
    assert healed.manifests


def test_an_artifact_export_failure_is_recorded_as_failed_and_the_next_run_retries(
    tmp_path: Path,
):
    scenario = _scenario(tmp_path)
    options = _idempotent()

    def fail_export(run, manifests):
        raise ValueError("renderer unavailable")

    failed = _run(scenario, idempotency=options, artifact_hook=fail_export)

    assert failed.delivered is False
    assert failed.manifests == ()
    assert [item.code for item in failed.failures] == ["format_export_failed"]
    record = load_run_record(state_path(scenario["delivery"]), TASK_KEY)
    assert record is not None
    assert record.status == FAILED
    assert record.successful is False

    base_manifest_paths = tuple(
        Path(item.delivery_path).with_suffix(".xlsx.manifest.json")
        for item in failed.targets
        if item.delivery_path
    )
    assert base_manifest_paths
    assert all(path.is_file() for path in base_manifest_paths)

    evidence = {
        "actual_format": "csv",
        "output": "recovered.csv",
        "sheets": ["北区"],
        "warnings": [],
        "digest": "sha256:recovered",
        "verification": "passed",
        "summary": "verified output",
    }

    def attach_recovered_export(run, manifests):
        return tuple(replace(manifest, exports=(evidence,)) for manifest in manifests)

    retried = _run(
        scenario,
        idempotency=options,
        artifact_hook=attach_recovered_export,
    )

    assert retried.no_op is False
    assert retried.run_decision.decision == RETRY
    assert retried.delivered is True
    assert retried.manifest_paths
    assert retried.manifests[0].exports == (evidence,)
    persisted = json.loads(Path(retried.manifest_paths[0]).read_text(encoding="utf-8"))
    assert persisted["exports"] == [evidence]
    assert _run(scenario, idempotency=options).no_op is True


def test_a_dry_run_reports_the_verdict_without_short_circuiting(tmp_path: Path):
    scenario = _scenario(tmp_path)
    options = _idempotent()
    assert _run(scenario, idempotency=options).delivered is True

    dry = plan_delivery(
        scenario["inputs"],
        scenario["targets"],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
        idempotency=options,
    )

    assert dry.dry_run is True
    assert dry.no_op is False, "a dry run still returns the full plan"
    assert dry.run_decision.decision == NO_OP
    assert dry.plan.targets, "the plan is still computed and checkable"
    assert _run(scenario, idempotency=options).no_op is True


def test_a_run_state_file_never_contains_a_raw_path_or_destination_name(tmp_path: Path):
    scenario = _scenario(tmp_path)
    assert _run(scenario, idempotency=_idempotent()).delivered is True

    text = state_path(scenario["delivery"]).read_text(encoding="utf-8")

    assert NORTH not in text and MAPLE not in text
    assert "巡检模板" not in text
    assert str(tmp_path) not in text


def test_cli_run_state_turns_an_unchanged_rerun_into_a_successful_no_op(tmp_path: Path):
    from excel_ops.cli import main

    scenario = _scenario(tmp_path)
    config = tmp_path / "delivery-plan.json"
    config.write_text(
        json.dumps(
            {
                "inputs": ["sources/weekly-export.csv"],
                "staging_dir": "staging",
                "delivery_dir": "delivery",
                "targets": [
                    {
                        "key": MAPLE,
                        "aliases": [MAPLE],
                        "template": "templates/client-upload.xlsx",
                        "sheet": "Upload",
                        "header_row": 1,
                        "data_start_row": 2,
                        "field_columns": {
                            "source": "A",
                            "record_id": "B",
                            "category": "C",
                            "identifier": "D",
                            "event_date": "E",
                            "location": "F",
                        },
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    result_path = tmp_path / "run.json"
    command = [
        "deliver",
        str(config),
        "--run-state",
        "--task-key",
        TASK_KEY,
        "--result",
        str(result_path),
    ]

    main(command)
    first = json.loads(result_path.read_text(encoding="utf-8"))
    assert first["delivered"] is True
    assert first["no_op"] is False

    # No SystemExit: an unchanged rerun is a success, not a failed delivery.
    main(command)
    second = json.loads(result_path.read_text(encoding="utf-8"))

    assert second["no_op"] is True
    assert second["run_decision"]["decision"] == NO_OP
    assert second["failures"] == []
    _assert_originals_untouched(scenario)
