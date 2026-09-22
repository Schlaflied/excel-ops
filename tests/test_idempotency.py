"""Whole-run idempotency tests (issue #19).

Every fixture here is fabricated. No customer, employee, payroll, or real
enterprise data — and no real credential — appears in this module.

These tests cover the *run-level* short-circuit only. The independent
per-record deduplication inside a delivered workbook is covered by
``tests/test_delivery.py`` and is deliberately untouched.
"""

from __future__ import annotations

import json
import sys
from datetime import date
from pathlib import Path

import pytest

from excel_ops.ambiguity import Ambiguity, decide
from excel_ops.idempotency import (
    CHANGED,
    COMPONENT_CONFIRMATIONS,
    COMPONENT_CONNECTOR,
    COMPONENT_INPUTS,
    COMPONENT_OUTPUT,
    COMPONENT_PERIOD,
    COMPONENT_RECIPE,
    COMPONENT_REMOTE_REVISION,
    COMPONENT_TEMPLATES,
    FAILED,
    NO_OP,
    RETRY,
    RUN_RECORD_FORMAT,
    STARTED,
    SUCCEEDED,
    UNREADABLE,
    ConnectorTarget,
    IdempotencyError,
    IdempotencyOptions,
    RunRecord,
    check_connector_conflict,
    compute_fingerprint,
    decide_run,
    evaluate_run,
    file_content_digest,
    format_decision,
    load_run_record,
    load_run_records,
    record_run,
    scrub_detail,
    state_path,
)
from excel_ops.periods import resolve_period


CONTENT = "Site,Inspection Date,Asset ID\n北区仓库,2026-09-08,AS-001\n"
TEMPLATE_CONTENT = "record_id,location,event_date\n"


# --------------------------------------------------------------------------- #
# Fabricated helpers
# --------------------------------------------------------------------------- #


def _write(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def _scenario(tmp_path: Path) -> dict[str, object]:
    """One fabricated periodic task: one input, one template, one local output."""

    source = _write(tmp_path / "sources" / "weekly-export.csv", CONTENT)
    template = _write(tmp_path / "templates" / "巡检模板.csv", TEMPLATE_CONTENT)
    output = tmp_path / "delivery" / "巡检模板-北区仓库.csv"
    return {
        "inputs": [source],
        "templates": [template],
        "connector": ConnectorTarget.local([output]),
        "output": output,
        "state": state_path(tmp_path / "delivery"),
    }


def _fingerprint(scenario: dict[str, object], **overrides):
    payload: dict[str, object] = {
        "inputs": scenario["inputs"],
        "templates": scenario["templates"],
        "connector": scenario["connector"],
    }
    payload.update(overrides)
    return compute_fingerprint(**payload)  # type: ignore[arg-type]


def _succeed(scenario: dict[str, object], fingerprint) -> RunRecord:
    return record_run(
        fingerprint,
        scenario["state"],
        task_key="task",
        status=SUCCEEDED,
        delivered=True,
        detail={"counts": {"written": 2}},
    )


class _StubCloudTarget:
    """A fake connector standing in for the Phase-2 cloud implementation.

    It implements exactly the :class:`RevisionSource` extension point a real
    cloud connector would implement, and counts how often the conflict check
    asked it for a revision.
    """

    def __init__(self, revision: str | None) -> None:
        self.revision = revision
        self.calls = 0

    def current_revision(self) -> str | None:
        self.calls += 1
        return self.revision


# --------------------------------------------------------------------------- #
# Acceptance criterion 1: same input + rules rerun returns no-op
# --------------------------------------------------------------------------- #


def test_same_input_and_rules_rerun_returns_no_op(tmp_path: Path):
    scenario = _scenario(tmp_path)
    first = _fingerprint(scenario)
    _succeed(scenario, first)

    second = _fingerprint(scenario)
    decision = evaluate_run(
        second, load_run_record(scenario["state"], "task"), connector=scenario["connector"]
    )

    assert second.digest == first.digest
    assert decision.decision == NO_OP
    assert decision.no_op is True
    assert decision.should_run is False
    assert decision.reason == "unchanged_since_successful_run"
    assert decision.changed_components == ()
    assert "no_op" in format_decision(decision)
    json.dumps(decision.to_dict())


def test_a_first_ever_run_is_changed_not_a_no_op(tmp_path: Path):
    scenario = _scenario(tmp_path)

    decision = decide_run(_fingerprint(scenario), None)

    assert decision.decision == CHANGED
    assert decision.reason == "no_previous_run"


def test_input_listing_order_does_not_change_the_fingerprint(tmp_path: Path):
    first = _write(tmp_path / "a.csv", CONTENT)
    second = _write(tmp_path / "b.csv", CONTENT + "北区仓库,2026-09-09,AS-002\n")

    forward = compute_fingerprint(inputs=[first, second])
    backward = compute_fingerprint(inputs=[second, first])

    assert forward.digest == backward.digest


# --------------------------------------------------------------------------- #
# Acceptance criterion 2: a renamed but identical file is not reprocessed
# --------------------------------------------------------------------------- #


def test_a_renamed_but_content_identical_file_does_not_trigger_reprocessing(tmp_path: Path):
    scenario = _scenario(tmp_path)
    _succeed(scenario, _fingerprint(scenario))

    renamed = _write(tmp_path / "sources" / "第38周-导出（副本）.csv", CONTENT)
    Path(scenario["inputs"][0]).unlink()

    decision = evaluate_run(
        _fingerprint(scenario, inputs=[renamed]),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == NO_OP, "a renamed but byte-identical input must not re-run"


def test_editing_a_file_without_renaming_it_does_trigger_a_new_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    _succeed(scenario, _fingerprint(scenario))

    _write(Path(scenario["inputs"][0]), CONTENT + "北区仓库,2026-09-09,AS-002\n")
    decision = evaluate_run(
        _fingerprint(scenario),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == CHANGED
    assert decision.changed_components == (COMPONENT_INPUTS,)


def test_an_unreadable_input_is_never_silently_hashed_as_content(tmp_path: Path):
    assert file_content_digest(tmp_path / "does-not-exist.csv") == UNREADABLE
    assert file_content_digest(_write(tmp_path / "x.csv", CONTENT)).startswith("sha256:")


# --------------------------------------------------------------------------- #
# Acceptance criterion 3: template, recipe, or review decision changes re-run
# --------------------------------------------------------------------------- #


def test_a_changed_template_triggers_a_new_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    _succeed(scenario, _fingerprint(scenario))

    _write(Path(scenario["templates"][0]), TEMPLATE_CONTENT + "identifier\n")
    decision = evaluate_run(
        _fingerprint(scenario),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == CHANGED
    assert decision.changed_components == (COMPONENT_TEMPLATES,)


def test_a_changed_template_profile_version_triggers_a_new_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    _succeed(scenario, _fingerprint(scenario, template_profile_version="profile-v1"))

    decision = evaluate_run(
        _fingerprint(scenario, template_profile_version="profile-v2"),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == CHANGED
    assert decision.changed_components == (COMPONENT_TEMPLATES,)


def _ambiguity() -> Ambiguity:
    return Ambiguity(
        "destination",
        "Which declared destination owns source values fuzzily matching [北区仓库]?",
        ("北区仓库附近",),
        ("北区仓库", "Maple Street Depot"),
        1,
        "北区仓库",
        "fuzzy candidate requires confirmation",
        0.8,
    )


def test_a_changed_recipe_decision_triggers_a_new_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    ambiguity = _ambiguity()
    north = decide(ambiguity, "北区仓库", scope="project", decided_at="2026-09-14T00:00:00+00:00")
    maple = decide(
        ambiguity, "Maple Street Depot", scope="project", decided_at="2026-09-14T00:00:00+00:00"
    )
    _succeed(scenario, _fingerprint(scenario, recipe_decisions=[north]))

    decision = evaluate_run(
        _fingerprint(scenario, recipe_decisions=[maple]),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == CHANGED
    assert decision.changed_components == (COMPONENT_RECIPE,)


def test_re_saving_the_same_recipe_answer_is_not_a_change(tmp_path: Path):
    """Provenance is not the rule: a new ``decided_at`` is not a new decision."""

    scenario = _scenario(tmp_path)
    ambiguity = _ambiguity()
    first = decide(ambiguity, "北区仓库", scope="project", decided_at="2026-09-14T00:00:00+00:00")
    again = decide(
        ambiguity,
        "北区仓库",
        scope="project",
        source="cli",
        decided_at="2026-09-21T09:30:00+00:00",
    )
    _succeed(scenario, _fingerprint(scenario, recipe_decisions=[first]))

    decision = evaluate_run(
        _fingerprint(scenario, recipe_decisions=[again]),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == NO_OP


def test_a_changed_human_review_confirmation_triggers_a_new_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    ambiguity = _ambiguity()
    applied = decide(ambiguity, "北区仓库", scope="this-run", decided_at="2026-09-14T00:00:00+00:00")
    _succeed(scenario, _fingerprint(scenario, confirmations=[applied]))

    reversed_call = decide(
        ambiguity, "Maple Street Depot", scope="this-run", decided_at="2026-09-14T00:00:00+00:00"
    )
    decision = evaluate_run(
        _fingerprint(scenario, confirmations=[reversed_call]),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == CHANGED
    assert decision.changed_components == (COMPONENT_CONFIRMATIONS,)


def test_withdrawing_every_confirmation_also_triggers_a_new_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    applied = decide(
        _ambiguity(), "北区仓库", scope="this-run", decided_at="2026-09-14T00:00:00+00:00"
    )
    _succeed(scenario, _fingerprint(scenario, confirmations=[applied]))

    decision = evaluate_run(
        _fingerprint(scenario),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == CHANGED
    assert COMPONENT_CONFIRMATIONS in decision.changed_components


def test_a_changed_report_period_triggers_a_new_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    september = resolve_period(
        "自定义", custom_start=date(2026, 9, 7), custom_end=date(2026, 9, 13)
    )
    october = resolve_period(
        "自定义", custom_start=date(2026, 10, 5), custom_end=date(2026, 10, 11)
    )
    _succeed(scenario, _fingerprint(scenario, period=september))

    decision = evaluate_run(
        _fingerprint(scenario, period=october),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == CHANGED
    assert decision.changed_components == (COMPONENT_PERIOD,)


def test_the_same_period_resolved_on_a_later_day_is_not_a_change(tmp_path: Path):
    """``as_of_date`` moves daily; the resolved window is what identifies a run."""

    scenario = _scenario(tmp_path)
    monday = resolve_period(
        "自定义",
        custom_start=date(2026, 9, 7),
        custom_end=date(2026, 9, 13),
        as_of=date(2026, 9, 14),
    )
    friday = resolve_period(
        "自定义",
        custom_start=date(2026, 9, 7),
        custom_end=date(2026, 9, 13),
        as_of=date(2026, 9, 18),
    )
    assert monday.as_of_date != friday.as_of_date
    _succeed(scenario, _fingerprint(scenario, period=monday))

    decision = evaluate_run(
        _fingerprint(scenario, period=friday),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == NO_OP


def test_a_changed_output_target_triggers_a_new_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    _succeed(scenario, _fingerprint(scenario))

    elsewhere = ConnectorTarget.local([tmp_path / "delivery" / "another-target.csv"])
    decision = evaluate_run(
        _fingerprint(scenario, connector=elsewhere),
        load_run_record(scenario["state"], "task"),
        connector=elsewhere,
    )

    assert decision.decision == CHANGED
    assert COMPONENT_CONNECTOR in decision.changed_components


def test_editing_the_delivered_output_triggers_a_new_run(tmp_path: Path):
    """The output content hash is in scope, so a hand-edited delivery re-runs."""

    scenario = _scenario(tmp_path)
    _write(Path(scenario["output"]), "record_id,location\nrec_1,北区仓库\n")
    _succeed(scenario, _fingerprint(scenario))

    _write(Path(scenario["output"]), "record_id,location\nrec_1,Maple Street Depot\n")
    decision = evaluate_run(
        _fingerprint(scenario),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == CHANGED
    assert decision.changed_components == (COMPONENT_OUTPUT,)


# --------------------------------------------------------------------------- #
# Acceptance criterion 4: a failed run is never a successful completion
# --------------------------------------------------------------------------- #


def test_a_failed_run_is_never_treated_as_a_completed_no_op_baseline(tmp_path: Path):
    scenario = _scenario(tmp_path)
    fingerprint = _fingerprint(scenario)
    record_run(
        fingerprint,
        scenario["state"],
        task_key="task",
        status=FAILED,
        delivered=False,
        detail={"failure_codes": ["verification_failed"]},
    )

    previous = load_run_record(scenario["state"], "task")
    decision = evaluate_run(
        _fingerprint(scenario), previous, connector=scenario["connector"]
    )

    assert previous.successful is False
    assert decision.decision == RETRY
    assert decision.reason == f"previous_run_{FAILED}"
    assert decision.no_op is False
    assert decision.should_run is True


def test_an_interrupted_run_is_retried_rather_than_reported_complete(tmp_path: Path):
    scenario = _scenario(tmp_path)
    fingerprint = _fingerprint(scenario)
    record_run(fingerprint, scenario["state"], task_key="task", status=STARTED)

    decision = evaluate_run(
        _fingerprint(scenario),
        load_run_record(scenario["state"], "task"),
        connector=scenario["connector"],
    )

    assert decision.decision == RETRY
    assert decision.reason == f"previous_run_{STARTED}"


def test_a_retry_that_then_succeeds_becomes_the_no_op_baseline(tmp_path: Path):
    scenario = _scenario(tmp_path)
    fingerprint = _fingerprint(scenario)
    failed = record_run(fingerprint, scenario["state"], task_key="task", status=FAILED)
    succeeded = record_run(
        fingerprint,
        scenario["state"],
        task_key="task",
        status=SUCCEEDED,
        delivered=True,
        previous=failed,
    )

    assert succeeded.attempt == 2, "a retry of the same fingerprint counts as a new attempt"
    assert (
        evaluate_run(
            _fingerprint(scenario),
            load_run_record(scenario["state"], "task"),
            connector=scenario["connector"],
        ).decision
        == NO_OP
    )


def test_an_unknown_run_status_is_refused(tmp_path: Path):
    with pytest.raises(IdempotencyError):
        RunRecord(task_key="task", fingerprint="abc", status="probably-fine")


# --------------------------------------------------------------------------- #
# Acceptance criterion 5: a cloud target revision change triggers a conflict
# check (forward-compatible interface, exercised with a stub connector)
# --------------------------------------------------------------------------- #


def _cloud(identifier: str = "cloud:workspace/fabricated-sheet") -> ConnectorTarget:
    return ConnectorTarget(identifier=identifier, kind="cloud")


def test_a_cloud_target_revision_change_triggers_a_conflict_check(tmp_path: Path):
    scenario = _scenario(tmp_path)
    connector = _cloud()
    stub = _StubCloudTarget("rev-1")
    _succeed(
        scenario,
        _fingerprint(scenario, connector=connector, revision_source=stub),
    )

    stub.revision = "rev-2"
    before = stub.calls
    previous = load_run_record(scenario["state"], "task")
    decision = evaluate_run(
        _fingerprint(scenario, connector=connector, revision_source=stub),
        previous,
        connector=connector,
        revision_source=stub,
    )

    assert stub.calls > before, "the conflict-check hook must ask the connector"
    assert decision.conflict is not None
    assert decision.conflict.code == "remote_revision_changed"
    assert decision.conflict.expected_revision != decision.conflict.actual_revision
    assert decision.decision == CHANGED
    assert COMPONENT_REMOTE_REVISION in decision.changed_components
    assert decision.no_op is False
    json.dumps(decision.to_dict())


def test_an_unchanged_cloud_revision_raises_no_conflict(tmp_path: Path):
    scenario = _scenario(tmp_path)
    connector = _cloud()
    stub = _StubCloudTarget("rev-1")
    _succeed(scenario, _fingerprint(scenario, connector=connector, revision_source=stub))

    previous = load_run_record(scenario["state"], "task")
    assert check_connector_conflict(previous, connector, revision_source=stub) is None
    decision = evaluate_run(
        _fingerprint(scenario, connector=connector, revision_source=stub),
        previous,
        connector=connector,
        revision_source=stub,
    )
    assert decision.decision == NO_OP


def test_an_unknown_current_revision_is_never_treated_as_unchanged(tmp_path: Path):
    scenario = _scenario(tmp_path)
    connector = _cloud()
    known = _StubCloudTarget("rev-1")
    _succeed(scenario, _fingerprint(scenario, connector=connector, revision_source=known))

    unknown = _StubCloudTarget(None)
    conflict = check_connector_conflict(
        load_run_record(scenario["state"], "task"), connector, revision_source=unknown
    )

    assert conflict is not None
    assert conflict.code == "remote_revision_changed"


def test_a_conflict_can_never_be_overridden_into_a_no_op(tmp_path: Path):
    """Even a byte-identical fingerprint cannot short-circuit past a conflict."""

    scenario = _scenario(tmp_path)
    fingerprint = _fingerprint(scenario)
    _succeed(scenario, fingerprint)
    previous = load_run_record(scenario["state"], "task")
    conflict = check_connector_conflict(
        previous, _cloud(), revision_source=_StubCloudTarget("moved")
    )

    decision = decide_run(fingerprint, previous, conflict=conflict)

    assert decision.decision == CHANGED
    assert decision.conflict is conflict


def test_the_conflict_check_is_inert_without_a_connector_or_a_previous_run(tmp_path: Path):
    scenario = _scenario(tmp_path)
    assert check_connector_conflict(None, _cloud()) is None
    _succeed(scenario, _fingerprint(scenario))
    assert check_connector_conflict(load_run_record(scenario["state"], "task"), None) is None


# --------------------------------------------------------------------------- #
# Acceptance criterion 6: no raw sensitive field or credential is persisted
# --------------------------------------------------------------------------- #


def test_the_run_record_never_contains_raw_sensitive_fields_or_credentials(tmp_path: Path):
    secret = "fabricated-token-DO-NOT-STORE-9f3c"
    connector = ConnectorTarget(
        identifier="cloud:workspace/fabricated-sheet",
        kind="cloud",
        revision="rev-1",
        credential_material=(secret,),
    )
    source = _write(tmp_path / "薪酬明细.csv", "employee,salary\n张三,123456\n")
    fingerprint = compute_fingerprint(inputs=[source], connector=connector)
    path = state_path(tmp_path / "delivery")

    record_run(
        fingerprint,
        path,
        task_key="task",
        status=SUCCEEDED,
        delivered=True,
        detail={
            "api_token": secret,
            "authorization": f"Bearer {secret}",
            "counts": {"written": 1},
            "note": "结果已交付",
        },
    )
    text = Path(path).read_text(encoding="utf-8")

    assert secret not in text
    assert "Bearer" not in text
    assert "salary" not in text and "123456" not in text
    assert "张三" not in text
    assert "薪酬明细" not in text
    # Every persisted component is a digest, never a value.
    stored = load_run_record(path, "task")
    assert set(stored.components) >= {COMPONENT_INPUTS, COMPONENT_CONNECTOR}
    assert all(len(value) == 64 and value.isalnum() for value in stored.components.values())
    assert stored.detail["api_token"].startswith("sha256:")
    assert stored.detail["counts"] == {"written": 1}


def test_a_credential_is_part_of_the_task_identity_without_being_stored(tmp_path: Path):
    first = ConnectorTarget(identifier="cloud:x", kind="cloud", credential_material=("token-a",))
    second = ConnectorTarget(identifier="cloud:x", kind="cloud", credential_material=("token-b",))

    assert compute_fingerprint(connector=first).digest != compute_fingerprint(
        connector=second
    ).digest
    assert "token-a" not in json.dumps(compute_fingerprint(connector=first).to_dict())
    assert "token-a" not in repr(first)


def test_scrub_detail_reduces_unexpected_structures_instead_of_storing_them():
    scrubbed = scrub_detail(
        {
            "password": "fabricated",
            "counts": {"written": 2, "secret_key": "fabricated"},
            "paths": [Path("C:/交付/北区.xlsx")],
            "handle": object(),
            "note": "x" * 500,
        }
    )

    assert scrubbed["password"].startswith("sha256:")
    assert scrubbed["counts"]["secret_key"].startswith("sha256:")
    assert scrubbed["handle"] == "object"
    assert len(scrubbed["note"]) == 200
    json.dumps(scrubbed)


# --------------------------------------------------------------------------- #
# Run-record store
# --------------------------------------------------------------------------- #


def test_the_state_file_follows_the_refresh_state_idiom(tmp_path: Path):
    path = state_path(tmp_path / "delivery")

    assert path.parent.name == ".excel-ops"
    assert path.name == "idempotency.json"
    assert load_run_records(path) == {}, "a missing state file is simply no previous run"


def test_several_tasks_share_one_state_file_without_overwriting_each_other(tmp_path: Path):
    scenario = _scenario(tmp_path)
    fingerprint = _fingerprint(scenario)
    record_run(fingerprint, scenario["state"], task_key="weekly", status=SUCCEEDED, delivered=True)
    record_run(fingerprint, scenario["state"], task_key="monthly", status=FAILED)

    records = load_run_records(scenario["state"])

    assert set(records) == {"weekly", "monthly"}
    assert records["weekly"].successful is True
    assert records["monthly"].successful is False
    assert json.loads(Path(scenario["state"]).read_text(encoding="utf-8"))["format"] == (
        RUN_RECORD_FORMAT
    )


def test_a_foreign_state_format_is_refused_instead_of_being_reinterpreted(tmp_path: Path):
    path = tmp_path / "state.json"
    path.write_text(
        json.dumps({"format": "excel-ops-workdir-recipe-v1", "overrides": []}), encoding="utf-8"
    )

    with pytest.raises(IdempotencyError):
        load_run_records(path)


def test_a_corrupt_state_file_is_refused_rather_than_read_as_no_previous_run(tmp_path: Path):
    path = tmp_path / "state.json"
    path.write_text("{not json", encoding="utf-8")

    with pytest.raises(IdempotencyError):
        load_run_records(path)


def test_the_default_task_key_is_an_opaque_digest_of_the_target(tmp_path: Path):
    options = IdempotencyOptions()
    connector = ConnectorTarget.local([tmp_path / "交付" / "北区仓库.xlsx"])

    key = options.resolved_task_key(connector)

    assert len(key) == 64 and key.isalnum()
    assert "北区仓库" not in key
    assert options.resolved_state_path(tmp_path / "delivery") == state_path(tmp_path / "delivery")
    assert IdempotencyOptions(task_key="weekly").resolved_task_key(connector) == "weekly"


# --------------------------------------------------------------------------- #
# Windows paths and non-ASCII names (standing repository convention)
# --------------------------------------------------------------------------- #


def test_non_ascii_filenames_round_trip(tmp_path: Path):
    """OS-agnostic: real filesystem interaction via tmp_path in its native form.

    Non-ASCII (Chinese) characters are not separator characters on any OS, so
    this exercises fingerprinting/no-op detection identically on Windows and
    Linux.
    """

    directory = tmp_path / "共享盘" / "本期交付"
    source = _write(directory / "北区巡检-源数据.csv", CONTENT)
    template = _write(directory / "巡检模板（最终）.csv", TEMPLATE_CONTENT)
    output = directory / "巡检模板-北区仓库.xlsx"
    connector = ConnectorTarget.local([output])

    fingerprint = compute_fingerprint(
        inputs=[source], templates=[template], connector=connector
    )
    path = state_path(directory)
    record_run(fingerprint, path, task_key="周报", status=SUCCEEDED, delivered=True)

    again = compute_fingerprint(inputs=[source], templates=[template], connector=connector)
    decision = evaluate_run(again, load_run_record(path, "周报"), connector=connector)

    assert decision.decision == NO_OP
    assert connector.identifier.startswith("local:")
    assert "\\" not in connector.identifier, "output targets are POSIX-normalised"
    assert Path(path).read_text(encoding="utf-8")  # written and readable as UTF-8
    json.dumps(decision.to_dict(), ensure_ascii=False)


@pytest.mark.skipif(
    sys.platform != "win32",
    reason="backslash-separated path strings are only meaningful to pathlib on Windows",
)
def test_windows_style_backslash_path_strings_are_accepted(tmp_path: Path):
    """Windows-only: ``file_content_digest`` has no explicit backslash-handling
    of its own -- it opens ``Path(path)`` directly, which only treats '\\' as a
    separator on Windows. On Linux a literal-backslash string never resolves to
    the real file, so this is stdlib platform behaviour, not a bug in
    ``idempotency.py``, and is only meaningfully testable on win32.
    """

    directory = tmp_path / "共享盘" / "本期交付"
    source = _write(directory / "北区巡检-源数据.csv", CONTENT)
    template = _write(directory / "巡检模板（最终）.csv", TEMPLATE_CONTENT)
    output = directory / "巡检模板-北区仓库.xlsx"
    # A path spelled the way Windows hands it to us, back-slashes and all.
    windows_spelling = Path(str(source).replace("/", "\\"))
    connector = ConnectorTarget.local([output])

    fingerprint = compute_fingerprint(
        inputs=[windows_spelling], templates=[template], connector=connector
    )
    path = state_path(directory)
    record_run(fingerprint, path, task_key="周报", status=SUCCEEDED, delivered=True)

    again = compute_fingerprint(inputs=[source], templates=[template], connector=connector)
    decision = evaluate_run(again, load_run_record(path, "周报"), connector=connector)

    assert decision.decision == NO_OP
