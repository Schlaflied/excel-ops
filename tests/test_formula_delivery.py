from decimal import Decimal
from pathlib import Path
import shutil
import zipfile
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook, load_workbook

from excel_ops.formula_delivery import (
    FormulaDeliveryError,
    FormulaDeliveryRule,
    apply_formula_delivery,
    formula_contract_fingerprint,
    formula_rule_from_mapping,
)
from excel_ops.formula_planning import FormulaPlan
from excel_ops.formula_recalculation import FormulaValueExpectation


def _formula_plan() -> FormulaPlan:
    return FormulaPlan(
        business_rule="Add three to the source value.",
        operation="addition",
        output_mode="formula",
        target_excel_version="365",
        value="=C2+3",
        function="addition",
        compatibility_strategy="basic arithmetic works in all target versions",
        requires_independent_recalculation=True,
    )


def test_formula_delivery_applies_and_independently_recalculates(tmp_path: Path):
    workbook_path = tmp_path / "delivery.xlsx"
    workbook = Workbook()
    workbook.active.title = "Report"
    workbook["Report"]["C2"] = 2
    workbook.save(workbook_path)
    workbook.close()

    def recalculate(source: Path, destination: Path) -> str:
        shutil.copy2(source, destination)
        _replace_cached_value(destination, "D2", "5")
        return "test-engine"

    evidence = apply_formula_delivery(
        workbook_path,
        (
            FormulaDeliveryRule(
                _formula_plan(),
                "Report",
                "D2",
                (FormulaValueExpectation("Report", "D2", 5),),
            ),
        ),
        recalculator=recalculate,
    )

    assert evidence["status"] == "verified"
    assert evidence["engine"] == "test-engine"
    assert evidence["rules"][0]["written_cells"] == 1
    assert "expected" not in str(evidence)


def _replace_cached_value(path: Path, cell_reference: str, value: str) -> None:
    replacement = path.with_suffix(".replacement.xlsx")
    worksheet_path = "xl/worksheets/sheet1.xml"
    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        replacement, "w"
    ) as destination:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == worksheet_path:
                root = ET.fromstring(data)
                cell = root.find(f".//{namespace}c[@r='{cell_reference}']")
                assert cell is not None
                cached = cell.find(f"{namespace}v")
                assert cached is not None
                cached.text = value
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            destination.writestr(item, data)
    replacement.replace(path)


def test_formula_rule_requires_independent_expectations():
    with pytest.raises(FormulaDeliveryError, match="requires independently derived"):
        formula_rule_from_mapping(
            {
                "plan": _formula_plan().to_dict(),
                "sheet": "Report",
                "target_range": "D2",
            }
        )


def test_expected_values_are_part_of_the_idempotency_fingerprint():
    first = FormulaDeliveryRule(
        _formula_plan(),
        "Report",
        "D2",
        (FormulaValueExpectation("Report", "D2", 5),),
    )
    second = FormulaDeliveryRule(
        _formula_plan(),
        "Report",
        "D2",
        (FormulaValueExpectation("Report", "D2", 6),),
    )

    assert first.fingerprint_payload() != second.fingerprint_payload()
    assert first.manifest_payload() == second.manifest_payload()


def test_static_value_types_are_part_of_the_contract_fingerprint():
    numeric = FormulaPlan(
        business_rule="Keep a numeric marker.",
        operation="static_marker",
        output_mode="static",
        target_excel_version="365",
        value=Decimal("1"),
        function=None,
        compatibility_strategy="static value",
        requires_independent_recalculation=False,
    )
    text = FormulaPlan(**{**numeric.to_dict(), "value": "1"})

    numeric_rule = FormulaDeliveryRule(numeric, "Report", "D2")
    text_rule = FormulaDeliveryRule(text, "Report", "D2")

    assert formula_contract_fingerprint((numeric_rule,)) != (
        formula_contract_fingerprint((text_rule,))
    )


def test_static_formula_rule_rejects_unverified_expectations():
    plan = _formula_plan().to_dict()
    plan.update(
        output_mode="static",
        value="approved",
        requires_independent_recalculation=False,
    )

    with pytest.raises(FormulaDeliveryError, match="static mode expectations"):
        formula_rule_from_mapping(
            {
                "plan": plan,
                "sheet": "Report",
                "target_range": "D2",
                "expectations": [
                    {"sheet": "Report", "cell": "D2", "expected": "approved"}
                ],
            }
        )


@pytest.mark.parametrize(
    ("field", "value", "message"),
    [
        (
            "requires_independent_recalculation",
            "false",
            "requires_independent_recalculation must be a boolean",
        ),
        ("overwrite_formulas", "false", "overwrite_formulas must be a boolean"),
    ],
)
def test_formula_rule_requires_strict_booleans(field, value, message):
    payload = {
        "plan": _formula_plan().to_dict(),
        "sheet": "Report",
        "target_range": "D2",
        "expectations": [{"sheet": "Report", "cell": "D2", "expected": 5}],
    }
    if field == "requires_independent_recalculation":
        payload["plan"][field] = value
    else:
        payload[field] = value

    with pytest.raises(FormulaDeliveryError, match=message):
        formula_rule_from_mapping(payload)
