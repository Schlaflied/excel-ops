import hashlib
import json
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_ops.delivery import load_delivery_targets
from excel_ops.plan_preparation import PlanPreparationError, prepare_delivery_plan


def _scenario(tmp_path: Path) -> dict:
    source = tmp_path / "来源.csv"
    source.write_text(
        "record_id,location,event_date,identifier,category,source\n"
        "R-1,North,2026-09-22,I-1,Demo,synthetic\n",
        encoding="utf-8",
    )
    template = tmp_path / "模板.xlsx"
    workbook = Workbook()
    workbook.active.title = "Data"
    workbook.save(template)
    return {
        "directory": str(tmp_path),
        "planPath": "delivery-plan.json",
        "inputs": [source.name],
        "targets": [
            {
                "key": "North",
                "template": template.name,
                "sheet": "Data",
                "fieldColumns": {
                    "record_id": "A",
                    "location": "B",
                    "event_date": "C",
                    "identifier": "D",
                    "category": "E",
                    "source": "F",
                },
                "requiredFields": ["record_id", "location"],
            }
        ],
    }


def test_prepare_writes_one_valid_plan_without_touching_workbooks(tmp_path: Path):
    request = _scenario(tmp_path)
    template = tmp_path / "模板.xlsx"
    before = template.read_bytes()

    result = prepare_delivery_plan(request)

    plan = tmp_path / "delivery-plan.json"
    payload = json.loads(plan.read_text(encoding="utf-8"))
    targets, options = load_delivery_targets(payload, base_dir=plan.parent)
    assert result["status"] == "prepared"
    assert result["next_tool"] == "excel_ops.plan_delivery"
    assert result["plan_digest"] == hashlib.sha256(plan.read_bytes()).hexdigest()
    assert len(targets) == 1
    assert options["inputs"] == [tmp_path / "来源.csv"]
    assert template.read_bytes() == before
    assert not (tmp_path / "delivery").exists()


def test_missing_business_decisions_return_review_without_a_plan(tmp_path: Path):
    result = prepare_delivery_plan(
        {
            "directory": str(tmp_path),
            "planPath": "delivery-plan.json",
            "inputs": [],
            "targets": [{"key": "North"}],
        }
    )

    assert result["status"] == "needs_review"
    assert {item["field"] for item in result["review"]} == {
        "inputs",
        "targets[0].template",
        "targets[0].sheet",
        "targets[0].fieldColumns",
    }
    assert not (tmp_path / "delivery-plan.json").exists()


def test_prepare_rejects_escape_and_symlink_escape(tmp_path: Path):
    request = _scenario(tmp_path)
    request["planPath"] = "../outside.json"
    with pytest.raises(PlanPreparationError, match="outside authorized roots"):
        prepare_delivery_plan(request)

    outside = tmp_path.parent / f"{tmp_path.name}-outside"
    outside.mkdir()
    link = tmp_path / "linked"
    try:
        link.symlink_to(outside, target_is_directory=True)
    except OSError:
        pytest.skip("directory symlinks are not available")
    request["planPath"] = "linked/plan.json"
    with pytest.raises(PlanPreparationError, match="outside authorized roots"):
        prepare_delivery_plan(request)


def test_replace_requires_the_current_digest(tmp_path: Path):
    request = _scenario(tmp_path)
    first = prepare_delivery_plan(request)

    with pytest.raises(PlanPreparationError) as conflict:
        prepare_delivery_plan(request)
    assert conflict.value.code == "plan_already_exists"

    request["replace"] = True
    request["expectedDigest"] = "0" * 64
    with pytest.raises(PlanPreparationError) as changed:
        prepare_delivery_plan(request)
    assert changed.value.code == "plan_changed"

    request["expectedDigest"] = first["plan_digest"]
    request["confidenceThreshold"] = 0.9
    replaced = prepare_delivery_plan(request)
    assert replaced["status"] == "prepared"
    assert replaced["plan_digest"] != first["plan_digest"]


def test_invalid_nested_shape_is_a_structured_request_error(tmp_path: Path):
    with pytest.raises(PlanPreparationError) as invalid:
        prepare_delivery_plan(
            {
                "directory": str(tmp_path),
                "planPath": "delivery-plan.json",
                "alsoAllow": None,
                "inputs": [],
                "targets": [],
            }
        )
    assert invalid.value.code == "invalid_request"


def test_digest_check_and_replace_are_serialized_per_plan(tmp_path: Path):
    request = _scenario(tmp_path)
    first = prepare_delivery_plan(request)
    low = deepcopy(request)
    high = deepcopy(request)
    for candidate, confidence in ((low, 0.8), (high, 0.9)):
        candidate["replace"] = True
        candidate["expectedDigest"] = first["plan_digest"]
        candidate["confidenceThreshold"] = confidence

    def attempt(candidate):
        try:
            return prepare_delivery_plan(candidate)
        except PlanPreparationError as error:
            return error

    with ThreadPoolExecutor(max_workers=2) as executor:
        outcomes = list(executor.map(attempt, (low, high)))

    assert sum(isinstance(item, dict) for item in outcomes) == 1
    errors = [item for item in outcomes if isinstance(item, PlanPreparationError)]
    assert len(errors) == 1
    assert errors[0].code == "plan_changed"
