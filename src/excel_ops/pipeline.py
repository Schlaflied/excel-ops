from __future__ import annotations

from pathlib import Path

from .extraction import load_extracted_json
from .review import review_record
from .workbook import verify_workbook, write_workbook


def run_pipeline(input_path: str | Path, output_path: str | Path, confidence_threshold: float = 0.85) -> dict[str, int | str]:
    extracted = load_extracted_json(input_path)
    reviewed = [review_record(record, confidence_threshold) for record in extracted]
    output = write_workbook(reviewed, output_path)
    verify_workbook(output, len(reviewed))
    accepted = sum(item.status == "accepted" for item in reviewed)
    return {"output": str(output), "total": len(reviewed), "accepted": accepted, "review": len(reviewed) - accepted}
