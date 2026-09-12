from __future__ import annotations

import json
from pathlib import Path

from .models import ExtractedRecord


def load_extracted_json(path: str | Path) -> list[ExtractedRecord]:
    """Load provider-neutral extraction output without trusting its fields."""
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or not isinstance(payload.get("records"), list):
        raise ValueError("Input must be an object containing a records array")
    source = str(payload.get("source") or Path(path).name)
    records = []
    for item in payload["records"]:
        if not isinstance(item, dict):
            continue
        enriched = dict(item)
        enriched.setdefault("source_file", source)
        records.append(ExtractedRecord.from_dict(enriched, source))
    return records
