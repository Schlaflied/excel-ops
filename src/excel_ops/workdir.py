"""Read-only working-directory scan, classification, and duplicate detection.

This module answers one question before any spreadsheet work starts: *which
files in this working directory are this period's inputs, and which must a human
look at first?*  It never moves, renames, writes, or deletes anything, and it
only ever touches paths inside an explicit allowlist of authorized roots.

Every file ends in exactly one classification -- ``input``, ``template``,
``prior_delivery``, ``review_return`` or ``unknown`` -- and exactly one
disposition -- ``include``, ``review`` or ``exclude`` -- with a machine-readable
reason for that disposition.  Nothing is silently picked:

* a file whose size or modification time has not been still for the stability
  window is never hashed, never opened, and never ``include``;
* byte-identical files are detected by content hash, so a copy under a
  different name is not processed twice;
* files that merely *look* like versions of each other (``final.xlsx``,
  ``final (1).xlsx``, ``final-final.xlsx``) but differ in content are grouped as
  candidates and all go to ``review``.  This module never elects "the real one";
* cloud-sync leftovers -- lock files, ``.crdownload``/``.part`` downloads,
  conflict copies -- are flagged, not ingested.  Resolving a sync conflict is
  explicitly out of scope;
* a user override can relabel a file and be saved as a reusable Recipe, but an
  override can never grant access outside the allowlist and can never promote an
  unstable or incomplete file to ``include``.

The scan is not wired into :mod:`excel_ops.delivery`; it is exposed on its own
as ``excel-ops scan-workdir`` for now.
"""
from __future__ import annotations

import json
import os
from dataclasses import asdict, replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from .workdir_classify import (
    _is_stable,
    _scan_file,
    sync_artifact_reason,
    version_key,
    workbook_metadata_signal,
)
from .workdir_models import (
    CLASSIFICATIONS,
    DATA_EXTENSIONS,
    DEFAULT_STABILITY_WINDOW_SECONDS,
    DISPOSITIONS,
    EXCLUDE,
    INCLUDE,
    INPUT,
    PRIOR_DELIVERY,
    READABLE_EXTENSIONS,
    REVIEW,
    REVIEW_RETURN,
    TEMPLATE,
    TEMPLATE_EXTENSIONS,
    UNKNOWN,
    WORKBOOK_EXTENSIONS,
    ClassificationOverride,
    DuplicateGroup,
    PeriodWindow,
    ScanResult,
    ScanScope,
    ScannedFile,
    StabilitySample,
    UnauthorizedPathError,
    VersionCandidateGroup,
    WorkdirScanError,
    _as_utc,
)

__all__ = [
    "CLASSIFICATIONS",
    "DATA_EXTENSIONS",
    "DEFAULT_STABILITY_WINDOW_SECONDS",
    "DISPOSITIONS",
    "EXCLUDE",
    "INCLUDE",
    "INPUT",
    "PRIOR_DELIVERY",
    "READABLE_EXTENSIONS",
    "RECIPE_FORMAT",
    "REVIEW",
    "REVIEW_RETURN",
    "TEMPLATE",
    "TEMPLATE_EXTENSIONS",
    "UNKNOWN",
    "WORKBOOK_EXTENSIONS",
    "ClassificationOverride",
    "DuplicateGroup",
    "PeriodWindow",
    "ScanResult",
    "ScanScope",
    "ScannedFile",
    "StabilitySample",
    "UnauthorizedPathError",
    "VersionCandidateGroup",
    "WorkdirScanError",
    "format_dry_run",
    "load_workdir_recipe",
    "override_from_entry",
    "save_workdir_recipe",
    "scan_workdir",
    "sync_artifact_reason",
    "version_key",
    "workbook_metadata_signal",
]


RECIPE_FORMAT = "excel-ops-workdir-recipe-v1"


def _walk(scope: ScanScope, skipped: list[dict[str, str]]) -> list[Path]:
    """Collect candidate files without ever leaving an authorized root.

    Containment is re-checked for every entry instead of being inferred from the
    parent directory, so a symlink or junction cannot widen the scan.
    """
    found: list[Path] = []
    seen: set[Path] = set()
    for root in scope.roots:
        pending = [root]
        while pending:
            directory = pending.pop()
            try:
                entries = list(os.scandir(directory))
            except OSError as error:
                skipped.append(
                    {"path": str(directory), "reason": f"unreadable_directory:{error.strerror or 'error'}"}
                )
                continue
            for entry in entries:
                candidate = Path(entry.path)
                if entry.is_symlink() and not scope.follow_symlinks:
                    skipped.append({"path": str(candidate), "reason": "symlink_not_followed"})
                    continue
                if scope.authorized_root(candidate) is None:
                    skipped.append({"path": str(candidate), "reason": "outside_authorized_roots"})
                    continue
                try:
                    is_directory = entry.is_dir(follow_symlinks=scope.follow_symlinks)
                except OSError as error:
                    skipped.append(
                        {"path": str(candidate), "reason": f"unreadable_entry:{error.strerror or 'error'}"}
                    )
                    continue
                if is_directory:
                    if scope.recursive:
                        pending.append(candidate)
                    else:
                        skipped.append({"path": str(candidate), "reason": "subdirectory_not_scanned"})
                    continue
                resolved = candidate.resolve()
                if resolved in seen:
                    continue
                seen.add(resolved)
                found.append(candidate)
    return sorted(found, key=lambda item: str(item).casefold())


def scan_workdir(
    roots: str | Path | Iterable[str | Path] | ScanScope,
    *,
    recursive: bool = True,
    period: Any | None = None,
    period_window: PeriodWindow | None = None,
    now: datetime | None = None,
    stability_window_seconds: float = DEFAULT_STABILITY_WINDOW_SECONDS,
    previous_samples: Mapping[str, StabilitySample] | None = None,
    overrides: Iterable[ClassificationOverride] | Mapping[str, ClassificationOverride] | None = None,
    recipe_path: str | Path | None = None,
) -> ScanResult:
    """Classify every authorized file and report what is included and why.

    The call is read-only.  ``period``/``period_window`` supply the authorized
    current-period window; without one, data files cannot be separated from
    historical files and are reported as ``unknown``/``review`` instead of being
    guessed into ``input``.
    """
    scope = roots if isinstance(roots, ScanScope) else ScanScope.of(roots, recursive=recursive)
    if period_window is None and period is not None:
        period_window = PeriodWindow.from_period(period)
    moment = datetime.now(timezone.utc) if now is None else _as_utc(now)
    if stability_window_seconds < 0:
        raise WorkdirScanError("stability window must not be negative")

    active_overrides = _collect_overrides(overrides, recipe_path)
    skipped: list[dict[str, str]] = []
    accessed: list[str] = []
    entries: list[ScannedFile] = []

    for candidate in _walk(scope, skipped):
        resolved = scope.require(candidate)
        relative = scope.relative(resolved)
        try:
            stat = resolved.stat()
        except OSError as error:
            skipped.append({"path": str(resolved), "reason": f"unreadable_file:{error.strerror or 'error'}"})
            continue
        previous = (previous_samples or {}).get(relative)
        stable = _is_stable(stat.st_size, stat.st_mtime, moment, stability_window_seconds, previous)
        entry, opened = _scan_file(resolved, relative, stat, stable, period_window)
        if opened:
            accessed.append(str(resolved))
        entries.append(entry)

    duplicates, deduplicated = _group_duplicates(entries)
    versions, grouped = _group_version_candidates(deduplicated)
    final_entries = tuple(_apply_overrides(item, active_overrides) for item in grouped)

    return ScanResult(
        scope=scope,
        files=final_entries,
        duplicate_groups=duplicates,
        version_groups=versions,
        skipped_paths=tuple(skipped),
        accessed_paths=tuple(accessed),
        period_window=period_window,
        stability_window_seconds=stability_window_seconds,
        scanned_at=moment.isoformat(),
    )


def _group_duplicates(
    entries: Sequence[ScannedFile],
) -> tuple[tuple[DuplicateGroup, ...], list[ScannedFile]]:
    """Collapse byte-identical files so the same content is processed once."""
    buckets: dict[str, list[ScannedFile]] = {}
    for item in entries:
        if item.content_hash:
            buckets.setdefault(item.content_hash, []).append(item)

    groups: list[DuplicateGroup] = []
    duplicate_of: dict[str, str] = {}
    for content_hash, members in buckets.items():
        if len(members) < 2:
            continue
        ordered = sorted(members, key=lambda item: item.relative_path)
        # The copies are byte-identical, so this picks which *name* to work
        # under, not which content is right: shortest path first, then
        # alphabetical, so repeated scans agree.
        representative = min(
            ordered, key=lambda item: (len(item.relative_path), item.relative_path)
        ).relative_path
        groups.append(
            DuplicateGroup(content_hash, tuple(item.relative_path for item in ordered), representative)
        )
        for item in ordered:
            if item.relative_path != representative:
                duplicate_of[item.relative_path] = representative

    updated: list[ScannedFile] = []
    for item in entries:
        representative = duplicate_of.get(item.relative_path)
        if representative is None:
            updated.append(item)
            continue
        # Identical bytes, so dropping the extra copy decides nothing about
        # content -- it only avoids doing the same work twice.
        updated.append(
            replace(
                item,
                disposition=EXCLUDE,
                reason="duplicate_content",
                duplicate_of=representative,
                notes=item.notes + (f"byte-identical to {representative}",),
            )
        )
    return tuple(sorted(groups, key=lambda group: group.representative)), updated


def _group_version_candidates(
    entries: Sequence[ScannedFile],
) -> tuple[tuple[VersionCandidateGroup, ...], list[ScannedFile]]:
    """Flag look-alike versions with differing content; never elect a winner."""
    buckets: dict[str, list[ScannedFile]] = {}
    for item in entries:
        if item.version_key and item.disposition != EXCLUDE:
            buckets.setdefault(item.version_key, []).append(item)

    groups: list[VersionCandidateGroup] = []
    flagged: set[str] = set()
    for key, members in buckets.items():
        hashes = {item.content_hash for item in members}
        if len(members) < 2 or len(hashes) < 2:
            continue
        ordered = sorted(members, key=lambda item: item.relative_path)
        groups.append(VersionCandidateGroup(key, tuple(item.relative_path for item in ordered)))
        flagged.update(item.relative_path for item in ordered)

    updated: list[ScannedFile] = []
    for item in entries:
        if item.relative_path not in flagged:
            updated.append(item)
            continue
        updated.append(
            replace(
                item,
                disposition=REVIEW,
                reason="version_candidates_require_confirmation",
                notes=item.notes + (f"one of several '{item.version_key}' versions",),
            )
        )
    return tuple(sorted(groups, key=lambda group: group.version_key)), updated


def _collect_overrides(
    overrides: Iterable[ClassificationOverride] | Mapping[str, ClassificationOverride] | None,
    recipe_path: str | Path | None,
) -> tuple[ClassificationOverride, ...]:
    collected: list[ClassificationOverride] = []
    if recipe_path is not None:
        collected.extend(load_workdir_recipe(recipe_path).values())
    if isinstance(overrides, Mapping):
        collected.extend(overrides.values())
    elif overrides is not None:
        collected.extend(overrides)
    merged: dict[str, ClassificationOverride] = {}
    for override in collected:
        merged[override.key] = override
    return tuple(merged.values())


def _apply_overrides(entry: ScannedFile, overrides: Sequence[ClassificationOverride]) -> ScannedFile:
    """Apply a human override without letting it defeat a safety floor."""
    match = next((item for item in overrides if item.matches(entry.relative_path)), None)
    if match is None:
        return entry
    notes = entry.notes
    if match.note:
        notes = notes + (match.note,)
    disposition = match.disposition or entry.disposition
    reason = f"override:{match.source}"
    if disposition == INCLUDE and not entry.stable:
        # A file that has not settled is still unread; an override cannot make it
        # final.
        disposition = REVIEW
        reason = "override_refused_unstable_file"
    elif disposition == INCLUDE and entry.reason.startswith("sync_artifact"):
        disposition = REVIEW
        reason = "override_refused_sync_artifact"
    return replace(
        entry,
        classification=match.classification,
        disposition=disposition,
        reason=reason,
        signal="override",
        overridden_from=entry.classification,
        notes=notes,
    )


def override_from_entry(
    entry: ScannedFile,
    classification: str,
    *,
    disposition: str | None = None,
    note: str = "",
    source: str = "user",
    decided_at: str | None = None,
) -> ClassificationOverride:
    """Build a reusable override from a scanned entry the user corrected."""
    return ClassificationOverride(
        path=entry.relative_path,
        classification=classification,
        disposition=disposition,
        note=note,
        decided_at=decided_at or datetime.now(timezone.utc).isoformat(),
        source=source,
    )


def save_workdir_recipe(overrides: Iterable[ClassificationOverride], path: str | Path) -> Path:
    """Persist overrides as a versioned, reusable Recipe file.

    The file layout mirrors :func:`excel_ops.ambiguity.save_project_recipe`, but
    the payload is keyed by relative path rather than by ``field:question``
    because a workdir decision is about a file, not about a data field.
    """
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    payload = {"format": RECIPE_FORMAT, "overrides": [asdict(item) for item in overrides]}
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output


def load_workdir_recipe(path: str | Path) -> dict[str, ClassificationOverride]:
    """Load a saved workdir Recipe, refusing any other format."""
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if payload.get("format") != RECIPE_FORMAT:
        raise WorkdirScanError("Unsupported workdir Recipe format")
    overrides: dict[str, ClassificationOverride] = {}
    for raw in payload.get("overrides", []):
        override = ClassificationOverride(**raw)
        overrides[override.key] = override
    return overrides


def format_dry_run(result: ScanResult) -> str:
    """Render the read-only inclusion/exclusion report with a reason per file."""
    lines = [
        "excel-ops working-directory scan (read-only)",
        f"roots: {', '.join(str(root) for root in result.scope.roots)}",
        f"recursive: {result.scope.recursive}  stability window: {result.stability_window_seconds}s",
    ]
    window = result.period_window
    lines.append(
        f"current period: {window.start.date()} .. {window.end.date()}"
        if window
        else "current period: not declared"
    )
    counts = result.counts()
    lines.append(f"include={counts[INCLUDE]} review={counts[REVIEW]} exclude={counts[EXCLUDE]}")
    for disposition in DISPOSITIONS:
        selected = result.by_disposition(disposition)
        lines.append(f"\n[{disposition}] {len(selected)} file(s)")
        for item in selected:
            detail = f"  {item.relative_path} -> {item.classification} ({item.reason}"
            detail += f", signal={item.signal}"
            if item.duplicate_of:
                detail += f", duplicate_of={item.duplicate_of}"
            lines.append(detail + ")")
    for group in result.version_groups:
        lines.append(
            f"\n[version candidates] '{group.version_key}': {', '.join(group.paths)}"
            " -> no version selected, human confirmation required"
        )
    for item in result.skipped_paths:
        lines.append(f"[skipped] {item['path']} ({item['reason']})")
    return "\n".join(lines)
