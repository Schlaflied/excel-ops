from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class ExtractedRecord:
    location: str
    event_date: str
    identifier: str
    category: str
    confidence: float
    source: str

    @classmethod
    def from_dict(cls, value: dict[str, Any], source: str) -> "ExtractedRecord":
        raw_confidence = value.get("confidence", 0)
        try:
            confidence = float(raw_confidence)
        except (TypeError, ValueError):
            confidence = 0.0
        return cls(
            location=str(value.get("location", "")).strip(),
            event_date=str(value.get("event_date", "")).strip(),
            identifier=str(value.get("identifier", "")).strip(),
            category=str(value.get("category", "")).strip(),
            confidence=max(0.0, min(1.0, confidence)),
            source=source,
        )


@dataclass(frozen=True)
class ReviewedRecord:
    record: ExtractedRecord
    status: str
    reasons: tuple[str, ...]
