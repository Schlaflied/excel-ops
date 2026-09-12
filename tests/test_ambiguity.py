import json

import pytest

from excel_ops.ambiguity import (
    Ambiguity,
    build_confirmation_batch,
    decide,
    load_project_recipe,
    review_rows_from_ambiguities,
    save_project_recipe,
)


def ambiguity(field, question, samples, candidates, recommendation, count=1):
    return Ambiguity(
        field,
        question,
        tuple(samples),
        tuple(candidates),
        count,
        recommendation,
        "synthetic evidence supports the recommendation",
        0.72,
        (f"input.xlsx:Sheet1!{field}",),
    )


def test_multiple_rows_are_grouped_into_one_review_question():
    first = ambiguity("amount", "currency", ["$12.00"], ["CAD", "USD"], "CAD", 2)
    second = ambiguity("amount", "currency", ["$20.00"], ["CAD", "USD"], "CAD", 3)
    batch = build_confirmation_batch([first, second])
    assert len(batch.items) == 1
    assert batch.items[0].ambiguity.affected_count == 5
    assert batch.items[0].ambiguity.samples == ("$12.00", "$20.00")
    assert batch.delivery_blocked


@pytest.mark.parametrize(
    ("field", "question", "samples", "candidates", "selected"),
    [
        ("amount", "currency", ["$1,250.00"], ["CAD", "USD"], "CAD"),
        ("service_date", "date_order", ["03/04/2026"], ["DD/MM/YYYY", "MM/DD/YYYY"], "DD/MM/YYYY"),
        ("utilization", "unit", ["12.5"], ["percent", "hours"], "percent"),
        ("employee_id", "leading_zero", ["00123"], ["preserve_text", "integer"], "preserve_text"),
    ],
)
def test_project_recipe_round_trip_for_core_ambiguities(tmp_path, field, question, samples, candidates, selected):
    item = ambiguity(field, question, samples, candidates, selected)
    decision = decide(item, selected, scope="project")
    path = save_project_recipe([decision], tmp_path / "recipe.json")
    loaded = load_project_recipe(path)
    batch = build_confirmation_batch([item], project_recipe=loaded)
    assert batch.items[0].status == "resolved"
    assert batch.items[0].selected == selected
    assert batch.items[0].decision_source == "project:user"
    assert batch.delivery_blocked is False
    assert batch.manifest_decisions()[0]["source"] == "project:user"


def test_this_run_decision_is_applied_but_not_persisted(tmp_path):
    item = ambiguity("amount", "currency", ["$10"], ["CAD", "USD"], "CAD")
    run = decide(item, "USD", scope="this-run")
    batch = build_confirmation_batch([item], run_decisions={run.key: run})
    assert batch.items[0].selected == "USD"
    path = save_project_recipe([run], tmp_path / "recipe.json")
    assert json.loads(path.read_text(encoding="utf-8"))["decisions"] == []


def test_unknown_remains_unknown_and_blocks_delivery():
    item = ambiguity("amount", "currency", ["$10"], ["CAD", "USD"], "CAD")
    unknown = decide(item, "不知道", scope="this-run")
    batch = build_confirmation_batch([item], run_decisions={unknown.key: unknown})
    assert batch.items[0].status == "unknown"
    assert batch.delivery_blocked
    assert batch.manifest_decisions()[0]["selected"] == "unknown"


def test_saved_rule_conflict_reopens_confirmation():
    old = ambiguity("amount", "currency", ["$10"], ["CAD", "USD"], "CAD")
    saved = decide(old, "CAD", scope="project")
    new = ambiguity("amount", "currency", ["€10"], ["EUR"], "EUR")
    batch = build_confirmation_batch([new], project_recipe={saved.key: saved})
    assert batch.items[0].status == "conflict"
    assert "冲突" in batch.items[0].conflict_reason
    assert batch.delivery_blocked


def test_conflicting_new_evidence_reopens_saved_rule_even_with_same_candidates():
    first = ambiguity("amount", "currency", ["$10"], ["CAD", "USD"], "CAD")
    saved = decide(first, "CAD", scope="project")
    second = ambiguity("amount", "currency", ["$20"], ["CAD", "USD"], "USD")
    batch = build_confirmation_batch([first, second], project_recipe={saved.key: saved})
    assert batch.items[0].status == "conflict"


def test_custom_rule_is_supported_and_reconfirmed_when_candidates_change():
    item = ambiguity("amount", "currency", ["$10"], ["CAD", "USD"], "CAD")
    custom = decide(item, "use workbook metadata", scope="project", allow_custom=True)
    assert custom.custom is True
    assert build_confirmation_batch([item], project_recipe={custom.key: custom}).items[0].status == "resolved"

    changed = ambiguity("amount", "currency", ["€10"], ["EUR"], "EUR")
    assert build_confirmation_batch([changed], project_recipe={custom.key: custom}).items[0].status == "conflict"


def test_review_pack_adapter_contains_samples_impact_and_reason():
    item = ambiguity("service_date", "date_order", ["03/04/2026"], ["DD/MM/YYYY", "MM/DD/YYYY"], "DD/MM/YYYY", 12)
    row = review_rows_from_ambiguities([item])[0]
    assert row.record_id == "ambiguity:service_date:date_order"
    assert row.source_value == "03/04/2026"
    assert "affected=12" in row.review_reason
    assert row.candidates == ("DD/MM/YYYY", "MM/DD/YYYY")
