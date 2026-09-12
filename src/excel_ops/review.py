from __future__ import annotations

from datetime import date

from .models import ExtractedRecord, ReviewedRecord


def review_record(record: ExtractedRecord, confidence_threshold: float = 0.85) -> ReviewedRecord:
    reasons: list[str] = []
    for field_name in ("location", "event_date", "identifier", "category"):
        if not getattr(record, field_name):
            reasons.append(f"missing:{field_name}")
    try:
        date.fromisoformat(record.event_date)
    except ValueError:
        reasons.append("invalid:event_date")
    if record.confidence < confidence_threshold:
        reasons.append("low_confidence")
    return ReviewedRecord(record, "review" if reasons else "accepted", tuple(reasons))
