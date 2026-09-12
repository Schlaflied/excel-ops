from __future__ import annotations

from .models import ExtractedRecord, ReviewedRecord
from .inference import infer_value


def review_record(record: ExtractedRecord, confidence_threshold: float = 0.85) -> ReviewedRecord:
    reasons: list[str] = []
    for field_name in ("location", "event_date", "identifier", "category"):
        if not getattr(record, field_name):
            reasons.append(f"missing:{field_name}")
    event_date = infer_value("event_date", record.event_date)
    if event_date.ambiguous:
        reasons.append(f"ambiguous:event_date:{event_date.reason}")
    elif event_date.inferred_type not in {"date", "datetime", "excel_serial_date"}:
        reasons.append("invalid:event_date")
    if record.confidence < confidence_threshold:
        reasons.append("low_confidence")
    return ReviewedRecord(record, "review" if reasons else "accepted", tuple(reasons))
