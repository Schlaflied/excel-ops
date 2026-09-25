"""Per-file stability, sync-artifact, and classification rules for the scan.

Content is only ever read through :func:`_read_bytes` and
:func:`workbook_metadata_signal`, so file access stays auditable.
"""

from __future__ import annotations

import hashlib
import os
import re
from dataclasses import replace
from datetime import datetime, timezone
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from .workdir_models import (
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
    PeriodWindow,
    ScannedFile,
    StabilitySample,
)


_LOCK_PREFIXES = ("~$", ".~lock.", "._")
_INCOMPLETE_EXTENSIONS = frozenset(
    {".tmp", ".temp", ".part", ".partial", ".partialdownload", ".crdownload", ".download", ".filepart", ".!ut"}
)
_CONFLICT_PATTERN = re.compile(
    r"conflicted\s+copy|conflict\s+copy|\(conflict(ed)?\b|冲突副本|沖突副本|_conflict-\d",
    re.IGNORECASE,
)
_TEMPLATE_PATTERN = re.compile(r"template|模板|範本|范本", re.IGNORECASE)
_REVIEW_RETURN_PATTERN = re.compile(
    r"review[\s_-]*(pack|return|returned|back)|reviewed|复核|覆核|審核|审核|回收|已批注",
    re.IGNORECASE,
)
_PRIOR_DELIVERY_PATTERN = re.compile(
    r"deliver(y|ed|able)|delivery[\s_-]*pack|已交付|交付件|交付|归档|歸檔|archive[d]?",
    re.IGNORECASE,
)
#: Trailing decorations people add when they save "one more" version.  They are
#: stripped to build a version key; they are never used to rank versions.
_VERSION_DECORATIONS = (
    re.compile(r"\s*\(\d+\)$"),
    re.compile(r"\s*-\s*copy(\s*\(\d+\))?$", re.IGNORECASE),
    re.compile(r"[\s_-]*(final|latest|new|old|draft)$", re.IGNORECASE),
    re.compile(r"[\s_-]*(?<![A-Za-z])v\d+(\.\d+)*$", re.IGNORECASE),
    re.compile(r"[\s_-]*(rev|version|ver)[\s_-]*\d+(\.\d+)*$", re.IGNORECASE),
    re.compile(r"[\s_-]*(最终版?|最新版?|终版|定稿|副本|修订版?)$"),
)


def _read_bytes(path: Path) -> bytes:
    """The single place file content is read, so access stays auditable."""
    return path.read_bytes()


def _hash_file(path: Path) -> str:
    return hashlib.sha256(_read_bytes(path)).hexdigest()


def _strip_version_decorations(stem: str) -> str:
    current = stem.strip()
    changed = True
    while changed:
        changed = False
        for pattern in _VERSION_DECORATIONS:
            stripped = pattern.sub("", current).strip(" _-")
            if stripped and stripped != current:
                current = stripped
                changed = True
    return current or stem.strip()


def version_key(name: str) -> str:
    """Return the shared key for ``final.xlsx`` / ``final (1).xlsx`` style names."""
    stem = Path(name).stem
    base = _strip_version_decorations(stem)
    return re.sub(r"[\s_-]+", " ", base).strip().casefold()


def sync_artifact_reason(name: str) -> str | None:
    """Return why ``name`` looks like a cloud-sync leftover, or ``None``."""
    lowered = name.casefold()
    if any(lowered.startswith(prefix.casefold()) for prefix in _LOCK_PREFIXES):
        return "sync_artifact_lock_file"
    suffix = Path(name).suffix.casefold()
    if suffix in _INCOMPLETE_EXTENSIONS:
        return "sync_artifact_incomplete_download"
    if lowered.endswith("~"):
        return "sync_artifact_editor_backup"
    if _CONFLICT_PATTERN.search(name):
        return "sync_artifact_conflict_copy"
    return None


def workbook_metadata_signal(path: Path) -> tuple[str, str] | None:
    """Return ``(classification, note)`` from workbook metadata, if it is decisive.

    Only sheet titles and document properties are inspected, with
    ``read_only=True``; no macro is executed and the workbook is not modified.
    """
    if path.suffix.casefold() not in WORKBOOK_EXTENSIONS:
        return None
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except (InvalidFileException, BadZipFile, OSError, KeyError, ValueError):
        # An unreadable or not-actually-a-workbook file yields no metadata
        # signal; it is never treated as evidence for a classification.
        return None
    try:
        titles = list(workbook.sheetnames)
        properties = workbook.properties
        descriptors = [
            str(getattr(properties, name, "") or "")
            for name in ("title", "category", "keywords", "subject")
        ]
    except (AttributeError, KeyError, ValueError):
        return None
    finally:
        workbook.close()

    haystack = " ".join(titles + descriptors)
    if _REVIEW_RETURN_PATTERN.search(haystack):
        return REVIEW_RETURN, "workbook metadata names a review pack"
    if _TEMPLATE_PATTERN.search(haystack):
        return TEMPLATE, "workbook metadata names a template"
    if _PRIOR_DELIVERY_PATTERN.search(haystack):
        return PRIOR_DELIVERY, "workbook metadata names a delivery"
    return None


def _name_signal(name: str) -> tuple[str, str] | None:
    stem = Path(name).stem
    if Path(name).suffix.casefold() in TEMPLATE_EXTENSIONS:
        return TEMPLATE, "template file extension"
    # Review returns are checked before deliveries: a returned review pack for a
    # delivered file usually carries both words.
    if _REVIEW_RETURN_PATTERN.search(stem):
        return REVIEW_RETURN, "filename names a review return"
    if _TEMPLATE_PATTERN.search(stem):
        return TEMPLATE, "filename names a template"
    if _PRIOR_DELIVERY_PATTERN.search(stem):
        return PRIOR_DELIVERY, "filename names a delivery"
    return None


def _scan_file(
    resolved: Path,
    relative: str,
    stat: os.stat_result,
    stable: bool,
    period_window: PeriodWindow | None,
) -> tuple[ScannedFile, bool]:
    """Classify one file; the flag says whether its content was opened."""

    modified = datetime.fromtimestamp(stat.st_mtime, timezone.utc)
    extension = resolved.suffix.casefold()
    entry = ScannedFile(
        relative_path=relative,
        path=str(resolved),
        classification=UNKNOWN,
        disposition=REVIEW,
        reason="not_classified",
        signal="none",
        extension=extension,
        size=stat.st_size,
        modified_at=modified.isoformat(),
        period_state="unknown",
        stable=stable,
    )

    artifact = sync_artifact_reason(resolved.name)
    if artifact is not None:
        disposition = REVIEW if artifact == "sync_artifact_conflict_copy" else EXCLUDE
        return replace(entry, disposition=disposition, reason=artifact, signal="name_pattern"), False

    if extension not in READABLE_EXTENSIONS:
        return (
            replace(entry, disposition=EXCLUDE, reason="unsupported_extension", signal="extension"),
            False,
        )

    if not stable:
        # Never opened: an unstable file may be mid-write or mid-sync, so it
        # is not hashed and not classified from content.
        return (
            replace(entry, disposition=REVIEW, reason="not_stable_yet", signal="stability_window"),
            False,
        )

    entry = replace(entry, content_hash=_hash_file(resolved), version_key=version_key(resolved.name))
    classification, signal, note = _classify_stable(resolved, modified, period_window)
    return _with_classification(entry, classification, signal, note, modified, period_window), True


def _with_classification(
    entry: ScannedFile,
    classification: str,
    signal: str,
    note: str,
    modified: datetime,
    period_window: PeriodWindow | None,
) -> ScannedFile:
    """Turn a stable file's classification into its disposition and reason."""

    if classification == INPUT:
        entry = replace(entry, disposition=INCLUDE, reason="current_period_input", period_state="current")
    elif classification == PRIOR_DELIVERY and signal == "modification_time":
        entry = replace(
            entry,
            disposition=EXCLUDE,
            reason="historical_outside_current_period",
            period_state="historical",
        )
    elif classification == UNKNOWN:
        entry = replace(entry, disposition=REVIEW, reason=note or "classification_undetermined")
    else:
        period_state = "current" if period_window and period_window.contains(modified) else "historical"
        entry = replace(
            entry,
            disposition=EXCLUDE if classification == PRIOR_DELIVERY else REVIEW,
            reason=f"classified_as_{classification}",
            period_state=period_state,
        )
    notes = (note,) if note else ()
    return replace(entry, classification=classification, signal=signal, notes=notes)


def _is_stable(
    size: int,
    modified_at: float,
    now: datetime,
    window_seconds: float,
    previous: StabilitySample | None,
) -> bool:
    # ``modified_at`` survives a round trip through an ISO string, which is
    # microsecond-precise, so compare within that resolution.
    if previous is not None and (
        previous.size != size or abs(previous.modified_at - modified_at) > 1e-6
    ):
        return False
    return (now.timestamp() - modified_at) >= window_seconds


def _classify_stable(
    path: Path,
    modified: datetime,
    period_window: PeriodWindow | None,
) -> tuple[str, str, str]:
    named = _name_signal(path.name)
    if named is not None:
        classification, note = named
        return classification, "name_pattern", note

    metadata = workbook_metadata_signal(path)
    if metadata is not None:
        classification, note = metadata
        return classification, "workbook_metadata", note

    if period_window is None:
        return UNKNOWN, "none", "no_period_window_declared"
    if period_window.contains(modified):
        return INPUT, "modification_time", "modified inside the declared current period"
    if modified < period_window.start:
        return PRIOR_DELIVERY, "modification_time", "modified before the declared current period"
    return UNKNOWN, "modification_time", "modified after the declared current period"
