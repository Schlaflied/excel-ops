"""Stable, JSON-serializable contracts for read-only workbook health scans."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any, Literal

HealthSeverity = Literal["info", "warning", "error"]
HealthFixability = Literal["safe", "review", "unsupported"]


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

    def __post_init__(self) -> None:
        if not self.code or not self.sheet or not self.location:
            raise ValueError("health findings require code, sheet, and location")
        if self.severity not in {"info", "warning", "error"}:
            raise ValueError(f"unsupported health severity: {self.severity}")
        if self.fixability not in {"safe", "review", "unsupported"}:
            raise ValueError(f"unsupported health fixability: {self.fixability}")

    def to_dict(self) -> dict[str, str]:
        return asdict(self)


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
