"""Re-scan, reconcile, and publish only verified workbook health repairs."""

from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import secrets
from typing import Any, Mapping, Sequence

from .health_models import HealthRepairAction, HealthRepairApplyError, HealthRepairChange
from .health_repair_apply import (
    _DirectoryGuard,
    _descriptor_digest,
    _identity,
    _opened_digest,
    apply_workbook_health_repairs,
)
from .health_repair_plan import ActionBaseline, plan_workbook_health_repairs
from .health_scan import scan_workbook_health


@dataclass(frozen=True)
class HealthRepairVerification:
    source: str
    source_sha256: str
    output: str | None
    output_sha256: str | None
    verified: bool
    before_findings: tuple[Mapping[str, Any], ...]
    authorized_actions: tuple[Mapping[str, Any], ...]
    actual_changes: tuple[HealthRepairChange, ...]
    remaining_findings: tuple[Mapping[str, Any], ...]
    new_findings: tuple[Mapping[str, Any], ...]
    failures: tuple[Mapping[str, str], ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema": 1,
            "mode": "verified_repair",
            "source": self.source,
            "source_sha256": self.source_sha256,
            "output": self.output,
            "output_sha256": self.output_sha256,
            "verified": self.verified,
            "before": list(self.before_findings),
            "actions": list(self.authorized_actions),
            "changes": [item.to_dict() for item in self.actual_changes],
            "remaining": list(self.remaining_findings),
            "new": list(self.new_findings),
            "failures": list(self.failures),
        }


def health_workflow_contract(scan: Any, plan: Any, *, mode: str) -> dict[str, Any]:
    """Return the common CLI/MCP health contract for scan and dry-run modes."""

    if mode not in {"scan", "dry_run"}:
        raise ValueError(f"unsupported health workflow mode: {mode}")
    if scan.source_sha256 != plan.source_sha256:
        raise HealthRepairApplyError("source changed between health scan and repair planning")
    return {
        "schema": 1,
        "mode": mode,
        "source": scan.source,
        "source_sha256": scan.source_sha256,
        "output": None,
        "output_sha256": None,
        "verified": None,
        "before": [item.to_dict() for item in scan.findings],
        "actions": [item.to_dict() for item in plan.actions],
        "changes": [],
        "remaining": [item.to_dict() for item in scan.findings],
        "new": [],
        "failures": [],
    }


def _candidate_path(output: Path) -> Path:
    return output.parent / f".{output.stem}-health-verify-{secrets.token_hex(12)}{output.suffix}"


def _failure(code: str, message: str) -> dict[str, str]:
    return {"code": code, "message": message}


def _reconcile(before, after, actions: Sequence[HealthRepairAction]):
    before_by_id = {item.finding_id: item for item in before.findings}
    after_by_id = {item.finding_id: item for item in after.findings}
    selected_ids = {item.finding_id for item in actions}
    failures: list[dict[str, str]] = []
    unresolved = sorted(selected_ids & set(after_by_id))
    unauthorized_ids = set(before_by_id) - selected_ids
    missing_unauthorized = sorted(unauthorized_ids - set(after_by_id))
    new_ids = sorted(set(after_by_id) - set(before_by_id))
    if unresolved:
        failures.append(_failure("authorized_finding_remains", repr(unresolved)))
    if missing_unauthorized:
        failures.append(_failure("unauthorized_finding_changed", repr(missing_unauthorized)))
    if new_ids:
        failures.append(_failure("new_finding_introduced", repr(new_ids)))
    remaining = tuple(after_by_id[key].to_dict() for key in sorted(after_by_id))
    new = tuple(after_by_id[key].to_dict() for key in new_ids)
    return remaining, new, tuple(failures)


def _publish_candidate(
    guard: _DirectoryGuard,
    candidate: Path,
    output: Path,
    descriptor: int,
    candidate_identity: tuple[int, int],
    expected_digest: str,
    source: Path,
    source_identity: tuple[int, int],
    source_digest: str,
) -> None:
    if output.exists() or output.is_symlink():
        raise HealthRepairApplyError(f"output already exists: {output}")
    if guard.identity(candidate) != candidate_identity:
        raise HealthRepairApplyError("verified candidate path changed before publication")
    if _descriptor_digest(descriptor, candidate_identity) != expected_digest:
        raise HealthRepairApplyError("verified candidate changed before publication")
    if _opened_digest(source, source_identity) != source_digest:
        raise HealthRepairApplyError("source changed before verified output publication")
    guard.ensure_visible()
    guard.link(candidate, output)
    published_identity = guard.identity(output)
    if published_identity != candidate_identity:
        guard.unlink(output, published_identity)
        raise HealthRepairApplyError("published output identity does not match verified candidate")
    try:
        if _descriptor_digest(descriptor, candidate_identity) != expected_digest:
            raise HealthRepairApplyError("published output changed during publication")
        if _opened_digest(source, source_identity) != source_digest:
            raise HealthRepairApplyError("source changed during verified output publication")
        guard.ensure_visible()
    except Exception:
        guard.unlink(output, candidate_identity)
        raise


def repair_and_verify_workbook_health(
    source: str | Path,
    output: str | Path,
    *,
    selected_action_ids: Sequence[str],
    baselines: Mapping[str, ActionBaseline] | None = None,
) -> HealthRepairVerification:
    """Apply to a candidate, re-scan it, reconcile evidence, then publish."""

    source_path = Path(source).absolute()
    output_path = Path(output).absolute()
    source_identity = _identity(source_path)
    before = scan_workbook_health(source_path)
    if _opened_digest(source_path, source_identity) != before.source_sha256:
        raise HealthRepairApplyError("source changed during initial health scan")
    dispositions = {action_id: "selected" for action_id in selected_action_ids}
    plan = plan_workbook_health_repairs(
        source_path, dispositions=dispositions, baselines=baselines
    )
    if plan.source_sha256 != before.source_sha256:
        raise HealthRepairApplyError("source changed between health scan and repair planning")
    selected = tuple(action for action in plan.actions if action.action_id in dispositions)
    candidate = _candidate_path(output_path)
    parent_identity = _identity(output_path.parent)
    with _DirectoryGuard(output_path.parent, parent_identity) as guard:
        applied = apply_workbook_health_repairs(
            source_path, candidate, plan, selected_action_ids=selected_action_ids
        )
        candidate_identity = guard.identity(candidate)
        flags = os.O_RDONLY | getattr(os, "O_BINARY", 0) | getattr(os, "O_NOFOLLOW", 0)
        descriptor = os.open(candidate, flags)
        try:
            if _descriptor_digest(descriptor, candidate_identity) != applied.output_sha256:
                raise HealthRepairApplyError("candidate changed before verification")
            after = scan_workbook_health(candidate)
            if after.source_sha256 != applied.output_sha256:
                raise HealthRepairApplyError("candidate changed during verification scan")
            if guard.identity(candidate) != candidate_identity:
                raise HealthRepairApplyError("candidate path changed during verification scan")
            remaining, new, failures = _reconcile(before, after, selected)
            if not failures:
                if _opened_digest(source_path, source_identity) != before.source_sha256:
                    raise HealthRepairApplyError("source changed before verified output publication")
                _publish_candidate(
                    guard, candidate, output_path, descriptor,
                    candidate_identity, applied.output_sha256,
                    source_path, source_identity, before.source_sha256,
                )
            return HealthRepairVerification(
                source=str(source_path), source_sha256=before.source_sha256,
                output=str(output_path) if not failures else None,
                output_sha256=applied.output_sha256 if not failures else None,
                verified=not failures,
                before_findings=tuple(item.to_dict() for item in before.findings),
                authorized_actions=tuple(item.to_dict() for item in selected),
                actual_changes=applied.changes, remaining_findings=remaining,
                new_findings=new, failures=failures,
            )
        finally:
            os.close(descriptor)
            guard.unlink(candidate, candidate_identity)
