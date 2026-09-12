from __future__ import annotations

from pathlib import Path
from typing import Any

from .inference import infer_fields
from .ingestion import load_input_records
from .review import review_record
from .workbook import verify_workbook, write_workbook


def run_pipeline(
    input_path: str | Path,
    output_path: str | Path,
    confidence_threshold: float = 0.85,
    locale: str | None = None,
) -> dict[str, Any]:
    extracted = load_input_records(input_path)
    field_inferences = infer_fields(
        [
            {
                "location": item.location,
                "event_date": item.event_date,
                "identifier": item.identifier,
                "category": item.category,
                "confidence": item.confidence,
            }
            for item in extracted
        ],
        locale=locale,
    )
    reviewed = [review_record(record, confidence_threshold) for record in extracted]
    output = write_workbook(reviewed, output_path, field_inferences)
    verify_workbook(output, len(reviewed))
    accepted = sum(item.status == "accepted" for item in reviewed)
    return {
        "output": str(output),
        "total": len(reviewed),
        "accepted": accepted,
        "review": len(reviewed) - accepted,
        "type_inference": [item.to_dict() for item in field_inferences],
        "low_confidence_fields": [
            item.field for item in field_inferences if item.low_confidence_count or item.ambiguous_count
        ],
    }
