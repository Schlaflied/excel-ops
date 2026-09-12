from pathlib import Path
from uuid import uuid4

from openpyxl import load_workbook

from excel_ops.pipeline import run_pipeline


def test_pipeline_routes_uncertain_rows():
    output = Path.cwd() / f".test-delivery-{uuid4().hex}.xlsx"
    try:
        result = run_pipeline("examples/extracted-records.json", output)
        assert result["total"] == 2
        assert result["accepted"] == 1
        assert result["review"] == 1
        assert {item["field"] for item in result["type_inference"]} == {
            "location", "event_date", "identifier", "category", "confidence"
        }

        wb = load_workbook(output, data_only=True)
        assert wb["Accepted"].max_row == 2
        assert wb["Review"].max_row == 2
        assert "low_confidence" in wb["Review"]["G2"].value
        assert "Type Inference" in wb.sheetnames
        assert wb["Type Inference"]["A2"].value == "location"
        wb.close()
    finally:
        output.unlink(missing_ok=True)
