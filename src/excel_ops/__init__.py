"""Excel-Ops public package."""

from .inference import FieldInference, ValueInference, infer_fields, infer_value
from .matching import Destination, MatchCandidate, MatchResult, preallocate_records, stable_record_id
from .period_refresh import DateRole, DateSlot, RefreshChange, RefreshResult, SlotTarget, refresh_workbook_period
from .periods import AmbiguousPeriodError, PeriodResolutionError, PeriodResult, resolve_period
from .naming import ResolvedOutput, resolve_output_path, verify_output_name
from .template_writer import (
    TemplateChange,
    TemplateMapping,
    TemplateSkip,
    TemplateWriteError,
    TemplateWriteResult,
    write_template,
)
from .delivery_verification import (
    DeliveryContract,
    DeliveryVerificationError,
    DeliveryVerificationResult,
    PeriodExpectation,
    StageCounts,
    VerificationFinding,
    WorkbookVerifier,
    verify_and_deliver,
)
from .formula_verification import (
    ExpectedErrorMarker,
    FormulaRegion,
    FormulaVerifier,
    SummaryReconciliation,
)
from .pipeline import run_pipeline
from .review_pack import (
    ReviewDecision,
    ReviewHistoryEntry,
    ReviewImportResult,
    ReviewPackRow,
    import_review_pack,
    review_rows_from_match_results,
    write_review_pack,
)

__all__ = [
    "Destination",
    "DateRole",
    "DateSlot",
    "AmbiguousPeriodError",
    "FieldInference",
    "MatchCandidate",
    "MatchResult",
    "PeriodResolutionError",
    "PeriodResult",
    "ReviewDecision",
    "ReviewHistoryEntry",
    "ReviewImportResult",
    "ReviewPackRow",
    "RefreshChange",
    "RefreshResult",
    "SlotTarget",
    "ResolvedOutput",
    "TemplateChange",
    "TemplateMapping",
    "TemplateSkip",
    "TemplateWriteError",
    "TemplateWriteResult",
    "ValueInference",
    "import_review_pack",
    "infer_fields",
    "infer_value",
    "preallocate_records",
    "refresh_workbook_period",
    "resolve_period",
    "review_rows_from_match_results",
    "resolve_output_path",
    "run_pipeline",
    "stable_record_id",
    "write_review_pack",
    "write_template",
    "DeliveryContract",
    "DeliveryVerificationError",
    "DeliveryVerificationResult",
    "PeriodExpectation",
    "StageCounts",
    "VerificationFinding",
    "WorkbookVerifier",
    "verify_and_deliver",
    "ExpectedErrorMarker",
    "FormulaRegion",
    "FormulaVerifier",
    "SummaryReconciliation",
    "verify_output_name",
]
