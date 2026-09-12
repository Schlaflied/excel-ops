"""Excel-Ops public package."""

from .inference import FieldInference, ValueInference, infer_fields, infer_value
from .matching import Destination, MatchCandidate, MatchResult, preallocate_records, stable_record_id
from .pipeline import run_pipeline

__all__ = [
    "Destination",
    "FieldInference",
    "MatchCandidate",
    "MatchResult",
    "ValueInference",
    "infer_fields",
    "infer_value",
    "preallocate_records",
    "run_pipeline",
    "stable_record_id",
]
