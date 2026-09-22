"""Per-delivery source-and-verification Manifest tests (issue #20).

Every fixture in this module is fabricated. No customer, employee, payroll, or
real enterprise template data appears here. The "sensitive-looking" fixture is
invented test data whose only purpose is to prove it never reaches a Manifest.
"""

from __future__ import annotations

import json
import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from excel_ops.delivery import DeliveryTarget, plan_delivery, run_delivery
from excel_ops.delivery_manifest import (
    MANIFEST_FORMAT,
    build_delivery_manifests,
    format_manifest,
    format_manifests,
    load_delivery_manifest,
    write_delivery_manifest,
)
from excel_ops.idempotency import file_content_digest, recipe_version_digest
from excel_ops.matching import Destination
from excel_ops.periods import resolve_period
from excel_ops.template_writer import TemplateMapping


FIELDS = ("record_id", "location", "event_date", "identifier", "category", "source")
DEPOT = "Maple Street Depot"
ANNEX = "River Annex"


# --------------------------------------------------------------------------- #
# Fabricated fixtures
# --------------------------------------------------------------------------- #


def _source_csv(path: Path, rows: tuple[tuple[str, str, str, str], ...]) -> None:
    lines = ["Location,Date,Asset ID,Category"]
    lines.extend(",".join(row) for row in rows)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _template(path: Path, sheet: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet
    for column, header in enumerate(FIELDS, start=1):
        worksheet.cell(1, column).value = header
    workbook.save(path)


def _target(template: Path, key: str, aliases: tuple[str, ...], sheet: str) -> DeliveryTarget:
    return DeliveryTarget(
        destination=Destination(key, aliases),
        template_path=template,
        mapping=TemplateMapping(
            sheet=sheet,
            header_row=1,
            data_start_row=2,
            field_columns={name: index + 1 for index, name in enumerate(FIELDS)},
        ),
        required_fields=FIELDS,
    )


def _scenario(tmp_path: Path) -> dict[str, object]:
    """Two inputs, two targets, and one record held back for review."""

    sources = tmp_path / "sources"
    templates = tmp_path / "templates"
    sources.mkdir()
    templates.mkdir()

    first = sources / "depot-week.csv"
    _source_csv(
        first,
        (
            (DEPOT, "2026-09-08", "AS-101", "routine"),
            (DEPOT, "2026-09-09", "AS-102", "routine"),
        ),
    )
    second = sources / "annex-week.csv"
    _source_csv(
        second,
        (
            (ANNEX, "2026-09-10", "AS-201", "routine"),
            # No location at all: rejected by the data contract, never matched.
            ("", "2026-09-11", "AS-202", "routine"),
        ),
    )

    depot_template = templates / "depot-template.xlsx"
    annex_template = templates / "annex-template.xlsx"
    _template(depot_template, "Depot")
    _template(annex_template, "Annex")

    return {
        "inputs": [first, second],
        "sources": [first, second],
        "templates": {DEPOT: depot_template, ANNEX: annex_template},
        "targets": [
            _target(depot_template, DEPOT, (DEPOT,), "Depot"),
            _target(annex_template, ANNEX, (ANNEX,), "Annex"),
        ],
        "staging": tmp_path / "staging",
        "delivery": tmp_path / "delivery",
    }


def _run(scenario: dict[str, object], **kwargs):
    return run_delivery(
        scenario["inputs"],
        scenario["targets"],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
        **kwargs,
    )


def _real_row_count(path: Path, sheet: str, *, data_start_row: int = 2, column: int = 1) -> int:
    """Count the rows really in the persisted workbook, by reopening it."""

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook[sheet]
        return sum(
            1
            for row in worksheet.iter_rows(
                min_row=data_start_row, min_col=column, max_col=column, values_only=True
            )
            if str(row[0] or "").strip()
        )
    finally:
        workbook.close()


def _by_key(run) -> dict[str, object]:
    return {item.destination_key: item for item in run.manifests}


# --------------------------------------------------------------------------- #
# Acceptance criterion: every source records file, count and hash
# --------------------------------------------------------------------------- #


def test_every_source_records_its_file_count_and_content_hash(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)

    assert run.delivered is True
    assert run.manifests, "a delivered run must produce a manifest"
    for manifest in run.manifests:
        listed = {item.file: item for item in manifest.sources}
        assert set(listed) == {path.name for path in scenario["sources"]}
        for path in scenario["sources"]:
            source = listed[path.name]
            assert source.hash == file_content_digest(path)
            assert source.hash.startswith("sha256:")
            assert source.records > 0
            # The count is real: it equals the records this file contributed.
            assert source.records == sum(source.statuses.values())
        # Records ingested per source add up to the run's input count.
        assert sum(item.records for item in manifest.sources) == manifest.run_counts["input"]


def test_a_source_is_identified_by_content_so_an_edit_changes_its_hash(tmp_path: Path):
    scenario = _scenario(tmp_path)
    before = _run(scenario).manifests[0]

    _source_csv(
        scenario["sources"][0],
        (
            (DEPOT, "2026-09-08", "AS-101", "routine"),
            (DEPOT, "2026-09-09", "AS-199", "deep clean"),
        ),
    )
    after = _run(_scenario_reuse(scenario)).manifests[0]

    first = scenario["sources"][0].name
    assert _hash_of(before, first) != _hash_of(after, first)


def _scenario_reuse(scenario: dict[str, object]) -> dict[str, object]:
    """Re-run into fresh staging/delivery directories with the same inputs."""

    delivery = Path(str(scenario["delivery"]))
    return {**scenario, "staging": delivery.parent / "staging-2", "delivery": delivery.parent / "delivery-2"}


def _hash_of(manifest, file_name: str) -> str:
    return next(item.hash for item in manifest.sources if item.file == file_name)


# --------------------------------------------------------------------------- #
# Acceptance criterion: accepted / review / rejected / written reconcile
# --------------------------------------------------------------------------- #


def test_counts_reconcile_with_the_real_row_count_in_the_delivered_workbook(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)

    manifests = _by_key(run)
    assert set(manifests) == {DEPOT, ANNEX}
    for key, manifest in manifests.items():
        delivered = Path(manifest.output)
        sheet = "Depot" if key == DEPOT else "Annex"
        real_rows = _real_row_count(delivered, sheet)
        # The manifest's own numbers, the run's counters, and the bytes on disk
        # all agree -- and the row count came from reopening the real file.
        assert manifest.written == real_rows
        assert manifest.rows == real_rows
        assert sum(tab.rows for tab in manifest.tabs) == real_rows
        assert sum(item.written for item in manifest.sources) == real_rows
        assert manifest.reconciled is True
        assert manifest.discrepancies == ()

    # The run-level states reconcile with the input count exactly once each.
    run_counts = run.manifests[0].run_counts
    assert run_counts["input"] == (
        run_counts["written"]
        + run_counts["accepted"]
        + run_counts["review"]
        + run_counts["rejected"]
        + run_counts["skipped_existing"]
    )
    # Every written record landed in exactly one delivered output.
    assert sum(item.written for item in run.manifests) == run_counts["written"]
    assert run_counts["rejected"] == 1  # the row with no location


def test_run_level_accepted_review_rejected_are_exposed_as_the_issue_declares(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)
    payload = run.manifests[0].to_dict()

    assert payload["accepted"] == run.counts["accepted"]
    assert payload["review"] == run.counts["review"]
    assert payload["rejected"] == run.counts["rejected"]
    assert payload["written"] == run.manifests[0].written
    assert payload["run_counts"] == dict(run.counts)


def test_a_row_count_disagreement_is_reported_instead_of_being_smoothed_over(tmp_path: Path):
    """The manifest must never claim a number the real file contradicts."""

    scenario = _scenario(tmp_path)
    run = _run(scenario)
    manifest = _by_key(run)[DEPOT]
    delivered = Path(manifest.output)

    # Delete a delivered row behind the manifest's back, then rebuild it.
    workbook = load_workbook(delivered)
    worksheet = workbook["Depot"]
    for column in range(1, len(FIELDS) + 1):
        worksheet.cell(2, column).value = None
    workbook.save(delivered)

    rebuilt = build_delivery_manifests(run, scenario["targets"])
    tampered = {item.destination_key: item for item in rebuilt}[DEPOT]

    assert tampered.rows < tampered.written
    assert tampered.reconciled is False
    assert f"row_count_mismatch:Depot" in tampered.discrepancies
    assert "written_count_mismatch" in tampered.discrepancies


# --------------------------------------------------------------------------- #
# Acceptance criterion: multi-tab deliveries record per-tab source and writes
# --------------------------------------------------------------------------- #


def test_multi_tab_delivery_records_per_tab_sources_and_write_counts(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)

    manifests = _by_key(run)
    assert [tab.sheet for tab in manifests[DEPOT].tabs] == ["Depot"]
    assert [tab.sheet for tab in manifests[ANNEX].tabs] == ["Annex"]

    depot_tab = manifests[DEPOT].tabs[0]
    assert depot_tab.written == 2
    assert depot_tab.rows == 2
    assert depot_tab.reconciled is True
    # The tab names the input files that fed it, and only those.
    contributing = {Path(name).name for name in depot_tab.sources}
    assert contributing == {"depot-week.csv"}
    assert sum(depot_tab.sources.values()) == depot_tab.written

    annex_tab = manifests[ANNEX].tabs[0]
    assert {Path(name).name for name in annex_tab.sources} == {"annex-week.csv"}
    assert annex_tab.written == annex_tab.rows == 1

    # Each delivered output is described by its own manifest and its own tabs.
    assert manifests[DEPOT].output != manifests[ANNEX].output
    assert manifests[DEPOT].tabs[0].sheet != manifests[ANNEX].tabs[0].sheet


def test_a_record_written_across_several_sheets_is_counted_in_each_tab(tmp_path: Path):
    """The per-tab breakdown follows the cells really written, not the mapping."""

    scenario = _scenario(tmp_path)
    run = _run(scenario)
    manifest = _by_key(run)[DEPOT]

    written = [
        item
        for item in run.records
        if item.status == "written" and item.destination_key == DEPOT
    ]
    assert written
    sheets = {cell.sheet for item in written for cell in item.cells}
    assert sheets == {tab.sheet for tab in manifest.tabs}


# --------------------------------------------------------------------------- #
# Acceptance criterion: template and recipe versions are recorded
# --------------------------------------------------------------------------- #


def test_template_and_recipe_versions_are_content_hashes_of_what_was_used(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)

    for key, manifest in _by_key(run).items():
        template = scenario["templates"][key]
        assert manifest.template == str(template)
        assert manifest.template_version == file_content_digest(template)
        # No Recipe was supplied, so the recipe version is the stable digest of
        # "no decisions" -- reused from idempotency, not reimplemented here.
        assert manifest.recipe_version == recipe_version_digest(())
        assert manifest.recipe_path is None


def test_editing_the_template_changes_the_recorded_template_version(tmp_path: Path):
    scenario = _scenario(tmp_path)
    before = _by_key(_run(scenario))[DEPOT].template_version

    template = scenario["templates"][DEPOT]
    workbook = load_workbook(template)
    workbook["Depot"]["H1"] = "extra declared column"
    workbook.save(template)
    after = _by_key(_run(_scenario_reuse(scenario)))[DEPOT].template_version

    assert before != after


def test_a_recipe_decision_changes_the_recorded_recipe_version(tmp_path: Path):
    from excel_ops.ambiguity import RecipeDecision

    decision = RecipeDecision(
        field="destination",
        question="Which declared destination owns this value?",
        selected=DEPOT,
        scope="project",
        source="user",
        decided_at="2026-09-14T00:00:00+00:00",
        candidates=(DEPOT, ANNEX),
    )

    assert recipe_version_digest(()) != recipe_version_digest((decision,))
    # Provenance is deliberately excluded: re-recording the same answer later is
    # not a rule change, so the version must not move.
    restated = RecipeDecision(
        field=decision.field,
        question=decision.question,
        selected=decision.selected,
        scope=decision.scope,
        source="another-reviewer",
        decided_at="2026-10-01T00:00:00+00:00",
        candidates=decision.candidates,
    )
    assert recipe_version_digest((decision,)) == recipe_version_digest((restated,))


# --------------------------------------------------------------------------- #
# Acceptance criterion: output filename, hash and verification status
# --------------------------------------------------------------------------- #


def test_output_name_hash_and_verification_status_are_recorded(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)

    for manifest in run.manifests:
        delivered = Path(manifest.output)
        assert delivered.is_file()
        assert manifest.output_name == delivered.name
        # Hashed from the persisted bytes, not from in-memory state.
        assert manifest.output_hash == file_content_digest(delivered)
        assert manifest.verification.status == "passed"
        assert manifest.verification.passed is True
        assert manifest.verification.findings == 0
        assert manifest.verification.codes == ()
        assert manifest.verification.report_path
        assert Path(manifest.verification.report_path).is_file()


def test_the_verification_block_mirrors_the_real_verification_result(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)

    for outcome in run.targets:
        manifest = next(
            item for item in run.manifests if item.destination_key == outcome.destination_key
        )
        report = json.loads(Path(str(outcome.verification_report_path)).read_text(encoding="utf-8"))
        # The manifest reports #4's own verdict for this delivery, not a new one.
        assert manifest.verification.passed is report["passed"]
        assert manifest.verification.findings == len(report["findings"])


def test_a_run_that_delivered_nothing_produces_no_manifest(tmp_path: Path):
    scenario = _scenario(tmp_path)

    planned = plan_delivery(
        scenario["inputs"],
        scenario["targets"],
        staging_dir=scenario["staging"],
        delivery_dir=scenario["delivery"],
    )

    assert planned.delivered is False
    assert planned.manifests == ()
    assert planned.manifest_paths == ()


def test_a_target_that_failed_verification_gets_no_manifest(tmp_path: Path):
    scenario = _scenario(tmp_path)

    def corrupt(staged: Path) -> None:
        staged.write_bytes(b"not a workbook")

    run = _run(scenario, post_stage_hook=corrupt)

    assert run.delivered is False
    assert run.manifests == ()
    assert not list(Path(str(scenario["delivery"])).glob("*.manifest.json"))


# --------------------------------------------------------------------------- #
# Acceptance criterion: sensitive cell content is never copied in
# --------------------------------------------------------------------------- #


SECRETS = (
    "SIN-046-454-286",
    "employee.private@example.invalid",
    "+1-555-0100",
    "Acct 4111111111111111",
)


def test_no_cell_value_from_the_delivery_appears_anywhere_in_the_manifest(tmp_path: Path):
    """Counts, hashes and metadata only -- never a raw cell value."""

    sources = tmp_path / "sources"
    templates = tmp_path / "templates"
    sources.mkdir()
    templates.mkdir()
    sensitive = sources / "sensitive-week.csv"
    _source_csv(
        sensitive,
        (
            (DEPOT, "2026-09-08", SECRETS[0], SECRETS[1]),
            (DEPOT, "2026-09-09", SECRETS[2], SECRETS[3]),
        ),
    )
    template = templates / "depot-template.xlsx"
    _template(template, "Depot")
    target = _target(template, DEPOT, (DEPOT,), "Depot")

    run = run_delivery(
        [sensitive],
        [target],
        staging_dir=tmp_path / "staging",
        delivery_dir=tmp_path / "delivery",
    )

    assert run.delivered is True
    manifest = run.manifests[0]
    assert manifest.written == 2
    # The values really are in the workbook -- so their absence below is a fact
    # about the manifest, not about the fixture.
    workbook = load_workbook(Path(manifest.output), read_only=True, data_only=True)
    try:
        cells = {
            str(value)
            for row in workbook["Depot"].iter_rows(min_row=2, values_only=True)
            for value in row
            if value is not None
        }
    finally:
        workbook.close()
    for secret in SECRETS:
        assert secret in cells, f"{secret!r} never reached the delivered workbook"

    serialized = manifest.to_json()
    readable = format_manifest(manifest)
    on_disk = manifest.manifest_path().read_text(encoding="utf-8")
    readable_on_disk = manifest.readable_path().read_text(encoding="utf-8")
    for secret in SECRETS:
        for rendering in (serialized, readable, on_disk, readable_on_disk):
            assert secret not in rendering, f"{secret!r} leaked into a manifest rendering"
    # A record ID is derived from source values, so it is metadata the manifest
    # also declines to copy.
    for record in run.records:
        assert record.record_id not in serialized


def test_a_verification_finding_message_is_never_copied_into_the_manifest(tmp_path: Path):
    """A finding message can quote the failing cell; only codes cross over."""

    from excel_ops.delivery_manifest import _verification
    from excel_ops.delivery import TargetOutcome
    from excel_ops.delivery_verification import VerificationFinding

    finding = VerificationFinding(
        "written_value_mismatch",
        f"Expected {SECRETS[0]!r}, got {SECRETS[1]!r}.",
        "Re-run write-back and verify the writer's real output path.",
        sheet="Depot",
        row=2,
        field="identifier",
        cell="C2",
    )
    outcome = TargetOutcome(DEPOT, "depot-template.xlsx", False, findings=(finding,))

    verification = _verification(outcome)
    serialized = json.dumps(verification.to_dict(), ensure_ascii=False)

    assert verification.status == "failed"
    assert verification.codes == ("written_value_mismatch",)
    assert verification.severities == {"error": 1}
    for secret in SECRETS:
        assert secret not in serialized


# --------------------------------------------------------------------------- #
# Acceptance criterion: consistent with the workbook's persisted state
# --------------------------------------------------------------------------- #


def test_the_output_hash_tracks_the_persisted_file_not_the_run(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)
    manifest = _by_key(run)[DEPOT]
    recorded = manifest.output_hash

    # Change the delivered file after the fact: the recorded hash no longer
    # matches, which is exactly how a tampered delivery is detected.
    delivered = Path(manifest.output)
    workbook = load_workbook(delivered)
    workbook["Depot"]["A99"] = "appended after delivery"
    workbook.save(delivered)

    assert file_content_digest(delivered) != recorded


def test_period_start_and_end_come_from_a_resolved_period_result(tmp_path: Path):
    scenario = _scenario(tmp_path)
    period = resolve_period("上周", as_of=date(2026, 9, 16))

    run = _run(scenario, period=period)

    for manifest in run.manifests:
        assert manifest.period_start == period.period_start.isoformat()
        assert manifest.period_end == period.period_end.isoformat()
        assert manifest.period_display_text == period.display_text
        assert manifest.to_dict()["period_start"] == "2026-09-07"
        assert manifest.to_dict()["period_end"] == "2026-09-13"


def test_no_declared_period_is_recorded_as_null_not_as_a_guess(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)

    for manifest in run.manifests:
        assert manifest.period_start is None
        assert manifest.period_end is None


# --------------------------------------------------------------------------- #
# Machine-readable and human-readable forms
# --------------------------------------------------------------------------- #


def test_manifest_is_written_as_json_and_as_a_readable_sibling(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)

    for manifest in run.manifests:
        machine = manifest.manifest_path()
        human = manifest.readable_path()
        assert machine.name == f"{Path(manifest.output).name}.manifest.json"
        assert human.name == f"{Path(manifest.output).name}.manifest.txt"
        assert machine.parent == Path(manifest.output).parent
        payload = load_delivery_manifest(machine)
        assert payload["format"] == MANIFEST_FORMAT
        assert payload == json.loads(manifest.to_json())
        text = human.read_text(encoding="utf-8")
        assert manifest.output_hash in text
        assert "delivery manifest:" in text
    assert list(run.manifest_paths) == [str(item.manifest_path()) for item in run.manifests]


def test_write_manifest_false_keeps_the_manifest_in_the_result_only(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario, write_manifest=False)

    assert run.manifests
    for manifest in run.manifests:
        assert not manifest.manifest_path().exists()
        assert not manifest.readable_path().exists()


def test_the_whole_run_result_stays_json_serializable(tmp_path: Path):
    scenario = _scenario(tmp_path)

    run = _run(scenario)
    payload = json.loads(json.dumps(run.to_dict(), ensure_ascii=False, default=str))

    assert len(payload["manifests"]) == 2
    assert payload["manifest_paths"]
    assert payload["manifests"][0]["format"] == MANIFEST_FORMAT
    assert format_manifests(run.manifests).count("delivery manifest:") == 2
    assert "nothing was delivered" in format_manifests(())


def test_loading_a_foreign_format_is_refused(tmp_path: Path):
    stray = tmp_path / "other.manifest.json"
    stray.write_text(json.dumps({"format": "something-else"}), encoding="utf-8")

    with pytest.raises(ValueError, match="unsupported delivery-manifest format"):
        load_delivery_manifest(stray)


def test_an_explicit_manifest_path_writes_both_forms_there(tmp_path: Path):
    scenario = _scenario(tmp_path)
    run = _run(scenario, write_manifest=False)
    elsewhere = tmp_path / "evidence" / "depot.manifest.json"

    written = write_delivery_manifest(_by_key(run)[DEPOT], path=elsewhere)

    assert written == elsewhere
    assert elsewhere.is_file()
    assert elsewhere.with_suffix(".txt").is_file()


# --------------------------------------------------------------------------- #
# Non-ASCII and platform-specific paths
# --------------------------------------------------------------------------- #


def test_chinese_and_non_ascii_paths_are_recorded_verbatim(tmp_path: Path):
    """Runs unconditionally on every platform, using a real ``tmp_path``."""

    sources = tmp_path / "来源目录"
    templates = tmp_path / "模板目录"
    sources.mkdir()
    templates.mkdir()
    csv_path = sources / "北区巡检-源数据.csv"
    _source_csv(csv_path, ((DEPOT, "2026-09-08", "AS-301", "routine"),))
    template = templates / "巡检模板.xlsx"
    _template(template, "北区")
    target = _target(template, DEPOT, (DEPOT,), "北区")

    run = run_delivery(
        [csv_path],
        [target],
        staging_dir=tmp_path / "暂存",
        delivery_dir=tmp_path / "交付",
    )

    assert run.delivered is True
    manifest = run.manifests[0]
    assert manifest.sources[0].file == "北区巡检-源数据.csv"
    assert manifest.tabs[0].sheet == "北区"
    assert manifest.manifest_path().is_file()
    reloaded = load_delivery_manifest(manifest.manifest_path())
    assert reloaded["sources"][0]["file"] == "北区巡检-源数据.csv"
    assert reloaded["tabs"]["北区"]["rows"] == 1


@pytest.mark.skipif(
    sys.platform != "win32",
    reason="backslash is only a path separator on Windows; a POSIX path never resolves one",
)
def test_a_windows_style_path_string_resolves_to_the_same_manifest(tmp_path: Path):
    """Windows-only: the manifest path round-trips through a backslash string.

    This never fabricates a Windows path by mangling a POSIX one -- on Windows
    ``tmp_path`` already is a ``WindowsPath``, so ``str(...)`` is a real native
    path and ``Path(...)`` resolves it back to the same file.
    """

    scenario = _scenario(tmp_path)
    run = _run(scenario)
    manifest = _by_key(run)[DEPOT]

    native = str(manifest.manifest_path())
    assert "\\" in native
    assert Path(native).is_file()
    assert load_delivery_manifest(native)["output_name"] == manifest.output_name
