"""Stable, JSON-serializable contracts for read-only workbook health scans."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from hashlib import sha256
import json
from typing import Any, Literal, Mapping

HealthSeverity = Literal["info", "warning", "error"]
HealthFixability = Literal["safe", "review", "unsupported"]
HealthActionDecision = Literal["proposed", "selected", "rejected"]


@dataclass(frozen=True)
class HealthFinding:
    """One observed workbook condition, without authorizing a repair."""

    code: str
    severity: HealthSeverity
    sheet: str
    location: str
    evidence: str
    suggestion: str
    fixability: HealthFixability
    details: Mapping[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not self.code or not self.sheet or not self.location:
            raise ValueError("health findings require code, sheet, and location")
        if self.severity not in {"info", "warning", "error"}:
            raise ValueError(f"unsupported health severity: {self.severity}")
        if self.fixability not in {"safe", "review", "unsupported"}:
            raise ValueError(f"unsupported health fixability: {self.fixability}")

    @property
    def finding_id(self) -> str:
        payload = json.dumps(asdict(self), ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return f"finding-{sha256(payload.encode('utf-8')).hexdigest()[:16]}"

    def to_dict(self) -> dict[str, Any]:
        return {"finding_id": self.finding_id, **asdict(self)}


@dataclass(frozen=True)
class HealthScanResult:
    """Evidence returned by a completed scan of one unchanged source file."""

    source: str
    source_sha256: str
    source_size: int
    sheets: tuple[str, ...]
    coverage: tuple[str, ...]
    findings: tuple[HealthFinding, ...]

    @property
    def assessment(self) -> Literal["findings_detected", "no_known_issues"]:
        return "findings_detected" if self.findings else "no_known_issues"

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema": 1,
            "source": self.source,
            "source_sha256": self.source_sha256,
            "source_size": self.source_size,
            "sheets": list(self.sheets),
            "coverage": list(self.coverage),
            "assessment": self.assessment,
            "findings": [item.to_dict() for item in self.findings],
        }


class HealthScanError(ValueError):
    """Raised when a workbook cannot be safely and completely scanned."""


@dataclass(frozen=True)
class HealthRepairAction:
    """One reviewable proposal derived from exactly one health finding."""

    action_id: str
    finding_id: str
    finding_code: str
    classification: HealthFixability
    decision: HealthActionDecision
    operation: str
    sheet: str
    location: str
    impact_scope: str
    current_state: Mapping[str, Any]
    proposed_state: Mapping[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class HealthRepairPlan:
    """A deterministic dry-run plan that never authorizes or applies writes."""

    source: str
    source_sha256: str
    actions: tuple[HealthRepairAction, ...]

    def actions_for(self, classification: HealthFixability) -> tuple[HealthRepairAction, ...]:
        return tuple(action for action in self.actions if action.classification == classification)

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema": 1,
            "mode": "dry_run",
            "source": self.source,
            "source_sha256": self.source_sha256,
            "counts": {
                classification: len(self.actions_for(classification))
                for classification in ("safe", "review", "unsupported")
            },
            "actions": [action.to_dict() for action in self.actions],
        }


class HealthRepairPlanError(ValueError):
    """Raised when a repair plan or requested action disposition is invalid."""
