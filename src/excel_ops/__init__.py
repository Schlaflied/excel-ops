"""Excel-Ops public package."""

from .pipeline import run_pipeline

__all__ = ["run_pipeline"]
from .matching import Destination, MatchCandidate, MatchResult, preallocate_records, stable_record_id

__all__ = [
    "Destination",
    "MatchCandidate",
    "MatchResult",
    "preallocate_records",
    "stable_record_id",
]
