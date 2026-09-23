import pytest

from excel_ops.delivery import DeliveryPlanError, load_delivery_targets


def _payload(**delivery):
    return {
        "inputs": ["input.csv"],
        "targets": [
            {
                "key": "demo",
                "template": "template.xlsx",
                "sheet": "Sheet1",
                "field_columns": {"record_id": "A", "name": "B"},
            }
        ],
        "delivery": delivery,
    }


def test_legacy_config_keeps_xlsx_default(tmp_path):
    _, options = load_delivery_targets(_payload(), base_dir=tmp_path)
    assert options["export_formats"] == ("xlsx",)


def test_csv_requires_explicit_single_sheet_or_per_sheet_mode(tmp_path):
    with pytest.raises(DeliveryPlanError, match=r"csv\.mode"):
        load_delivery_targets(_payload(formats=["csv"]), base_dir=tmp_path)


def test_csv_and_pdf_selection_is_json_safe_and_reports_warnings(tmp_path):
    _, options = load_delivery_targets(
        _payload(
            formats=["xlsx", "csv", "pdf"],
            csv={"mode": "one-file-per-sheet"},
            pdf={"sheets": ["Sheet1"]},
        ),
        base_dir=tmp_path,
    )
    assert options["export_formats"] == ("xlsx", "csv", "pdf")
    assert options["export_selection"]["warnings"] == [
        "csv_drops_workbook_features",
        "pdf_is_not_editable",
    ]


def test_pdf_all_sheets_is_supported(tmp_path):
    _, options = load_delivery_targets(
        _payload(formats=["pdf"], pdf={"sheets": "all"}), base_dir=tmp_path
    )
    assert options["export_selection"]["pdf"] is None


def test_csv_sheet_must_be_a_string(tmp_path):
    with pytest.raises(DeliveryPlanError, match="csv.sheet must be a string"):
        load_delivery_targets(
            _payload(formats=["csv"], csv={"mode": "single-sheet", "sheet": 3}),
            base_dir=tmp_path,
        )


def test_pdf_sheet_list_must_not_be_empty(tmp_path):
    with pytest.raises(DeliveryPlanError, match="pdf.sheets"):
        load_delivery_targets(
            _payload(formats=["pdf"], pdf={"sheets": []}), base_dir=tmp_path
        )
