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

import fnmatch
import hashlib
import json
import os
import re
from dataclasses import asdict, dataclass, replace
from datetime import date, datetime, time, timezone
from pathlib import Path, PurePosixPath
from typing import Any, Iterable, Mapping, Sequence
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


RECIPE_FORMAT = "excel-ops-workdir-recipe-v1"

INPUT = "input"
TEMPLATE = "template"
PRIOR_DELIVERY = "prior_delivery"
REVIEW_RETURN = "review_return"
UNKNOWN = "unknown"
CLASSIFICATIONS = (INPUT, TEMPLATE, PRIOR_DELIVERY, REVIEW_RETURN, UNKNOWN)

INCLUDE = "include"
REVIEW = "review"
EXCLUDE = "exclude"
DISPOSITIONS = (INCLUDE, REVIEW, EXCLUDE)

#: Extensions this project can actually read.  Anything else is ``unknown`` and
#: excluded rather than optimistically opened.
DATA_EXTENSIONS = frozenset({".xlsx", ".xlsm", ".csv", ".json"})
TEMPLATE_EXTENSIONS = frozenset({".xltx", ".xltm"})
WORKBOOK_EXTENSIONS = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})
READABLE_EXTENSIONS = DATA_EXTENSIONS | TEMPLATE_EXTENSIONS

#: Seconds a file's size and modification time must already have been unchanged
#: before its content may be read at all.
DEFAULT_STABILITY_WINDOW_SECONDS = 5.0

_DAY_START = time.min
_DAY_END = time.max

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


class WorkdirScanError(ValueError):
    """Raised when a scan cannot be performed without guessing."""


class UnauthorizedPathError(WorkdirScanError):
    """Raised when a path outside every authorized root would be touched."""


@dataclass(frozen=True)
class ScanScope:
    """The explicit allowlist of directories a scan may touch.

    Mirrors ``refresh.mjs``'s root-confinement idiom: a path is usable only when
    it resolves inside a declared root, symlinks are never followed out of a
    root, and containment is re-checked at the point of access rather than
    assumed from how the path was produced.
    """

    roots: tuple[Path, ...]
    recursive: bool = True
    follow_symlinks: bool = False

    @classmethod
    def of(
        cls,
        roots: str | Path | Iterable[str | Path],
        *,
        recursive: bool = True,
        follow_symlinks: bool = False,
    ) -> "ScanScope":
        if isinstance(roots, (str, Path)):
            candidates: list[str | Path] = [roots]
        else:
            candidates = list(roots)
        if not candidates:
            raise WorkdirScanError("a scan needs at least one authorized root directory")
        resolved: list[Path] = []
        for candidate in candidates:
            root = Path(candidate).expanduser().resolve()
            if not root.is_dir():
                raise WorkdirScanError(f"authorized root is not an existing directory: {root}")
            if root not in resolved:
                resolved.append(root)
        return cls(tuple(resolved), recursive, follow_symlinks)

    def authorized_root(self, path: str | Path) -> Path | None:
        """Return the root containing ``path``, or ``None`` when unauthorized."""
        try:
            target = Path(path).expanduser().resolve()
        except OSError:
            return None
        for root in self.roots:
            if target == root:
                return root
            try:
                target.relative_to(root)
            except ValueError:
                continue
            return root
        return None

    def require(self, path: str | Path) -> Path:
        """Resolve ``path`` inside the allowlist or refuse to touch it."""
        root = self.authorized_root(path)
        if root is None:
            raise UnauthorizedPathError(f"path is outside every authorized root: {path}")
        return Path(path).expanduser().resolve()

    def relative(self, path: str | Path) -> str:
        target = self.require(path)
        for root in self.roots:
            try:
                return PurePosixPath(target.relative_to(root)).as_posix()
            except ValueError:
                continue
        return target.name

    def to_dict(self) -> dict[str, Any]:
        return {
            "roots": [str(root) for root in self.roots],
            "recursive": self.recursive,
            "follow_symlinks": self.follow_symlinks,
        }


@dataclass(frozen=True)
class PeriodWindow:
    """The authorized current-period window used to separate historical files."""

    start: datetime
    end: datetime

    @classmethod
    def of(cls, start: date | datetime, end: date | datetime) -> "PeriodWindow":
        first = _as_utc(start)
        last = _as_utc(end, end_of_day=True)
        if last < first:
            raise WorkdirScanError("period window end is before its start")
        return cls(first, last)

    @classmethod
    def from_period(cls, period: Any) -> "PeriodWindow":
        """Build a window from an already resolved :class:`PeriodResult`."""
        return cls.of(period.period_start, period.period_end)

    def contains(self, moment: datetime) -> bool:
        return self.start <= moment <= self.end

    def to_dict(self) -> dict[str, str]:
        return {"start": self.start.isoformat(), "end": self.end.isoformat()}


@dataclass(frozen=True)
class StabilitySample:
    """A previously observed size/mtime pair for one relative path."""

    size: int
    modified_at: float


@dataclass(frozen=True)
class ClassificationOverride:
    """A human decision that replaces the automatic classification."""

    path: str
    classification: str
    disposition: str | None = None
    note: str = ""
    decided_at: str = ""
    source: str = "user"

    def __post_init__(self) -> None:
        if self.classification not in CLASSIFICATIONS:
            raise WorkdirScanError(f"unknown classification in override: {self.classification}")
        if self.disposition is not None and self.disposition not in DISPOSITIONS:
            raise WorkdirScanError(f"unknown disposition in override: {self.disposition}")
        if not str(self.path).strip():
            raise WorkdirScanError("an override needs a relative path or glob pattern")

    @property
    def key(self) -> str:
        return PurePosixPath(str(self.path).replace("\\", "/")).as_posix()

    def matches(self, relative_path: str) -> bool:
        pattern = self.key
        target = PurePosixPath(relative_path).as_posix()
        if pattern == target:
            return True
        return fnmatch.fnmatch(target, pattern)


@dataclass(frozen=True)
class ScannedFile:
    """One classified file plus every signal that produced the decision."""

    relative_path: str
    path: str
    classification: str
    disposition: str
    reason: str
    signal: str
    extension: str
    size: int
    modified_at: str
    period_state: str
    stable: bool
    content_hash: str | None = None
    version_key: str | None = None
    duplicate_of: str | None = None
    overridden_from: str | None = None
    notes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["notes"] = list(self.notes)
        return payload


@dataclass(frozen=True)
class DuplicateGroup:
    """Byte-identical files found under different names."""

    content_hash: str
    paths: tuple[str, ...]
    representative: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "content_hash": self.content_hash,
            "paths": list(self.paths),
            "representative": self.representative,
        }


@dataclass(frozen=True)
class VersionCandidateGroup:
    """Files that look like versions of one another but differ in content."""

    version_key: str
    paths: tuple[str, ...]
    selected: None = None
    reason: str = "version_candidates_require_confirmation"

    def to_dict(self) -> dict[str, Any]:
        return {
            "version_key": self.version_key,
            "paths": list(self.paths),
            "selected": None,
            "reason": self.reason,
        }


@dataclass(frozen=True)
class ScanResult:
    """The complete read-only report for one working-directory scan."""

    scope: ScanScope
    files: tuple[ScannedFile, ...]
    duplicate_groups: tuple[DuplicateGroup, ...] = ()
    version_groups: tuple[VersionCandidateGroup, ...] = ()
    skipped_paths: tuple[dict[str, str], ...] = ()
    accessed_paths: tuple[str, ...] = ()
    period_window: PeriodWindow | None = None
    stability_window_seconds: float = DEFAULT_STABILITY_WINDOW_SECONDS
    scanned_at: str = ""

    def by_disposition(self, disposition: str) -> tuple[ScannedFile, ...]:
        return tuple(item for item in self.files if item.disposition == disposition)

    def by_classification(self, classification: str) -> tuple[ScannedFile, ...]:
        return tuple(item for item in self.files if item.classification == classification)

    @property
    def included(self) -> tuple[ScannedFile, ...]:
        return self.by_disposition(INCLUDE)

    @property
    def review(self) -> tuple[ScannedFile, ...]:
        return self.by_disposition(REVIEW)

    @property
    def excluded(self) -> tuple[ScannedFile, ...]:
        return self.by_disposition(EXCLUDE)

    def entry(self, relative_path: str) -> ScannedFile | None:
        target = PurePosixPath(relative_path.replace("\\", "/")).as_posix()
        for item in self.files:
            if item.relative_path == target:
                return item
        return None

    def counts(self) -> dict[str, int]:
        counts = {name: 0 for name in DISPOSITIONS}
        for item in self.files:
            counts[item.disposition] += 1
        return counts

    def to_dict(self) -> dict[str, Any]:
        return {
            "format": "excel-ops-workdir-scan-v1",
            "scanned_at": self.scanned_at,
            "scope": self.scope.to_dict(),
            "period_window": self.period_window.to_dict() if self.period_window else None,
            "stability_window_seconds": self.stability_window_seconds,
            "counts": self.counts(),
            "classification_counts": {
                name: len(self.by_classification(name)) for name in CLASSIFICATIONS
            },
            "files": [item.to_dict() for item in self.files],
            "duplicate_groups": [group.to_dict() for group in self.duplicate_groups],
            "version_groups": [group.to_dict() for group in self.version_groups],
            "skipped_paths": [dict(item) for item in self.skipped_paths],
            "accessed_paths": list(self.accessed_paths),
        }

    def stability_samples(self) -> dict[str, StabilitySample]:
        """Return this run's size/mtime observations for a follow-up scan."""
        samples: dict[str, StabilitySample] = {}
        for item in self.files:
            samples[item.relative_path] = StabilitySample(
                item.size, datetime.fromisoformat(item.modified_at).timestamp()
            )
        return samples


def _as_utc(value: date | datetime, *, end_of_day: bool = False) -> datetime:
    """Normalise a date or datetime to an aware UTC instant.

    A bare ``date`` marks the whole day, so a period end covers that day rather
    than cutting it off at midnight and silently dropping the final day's files.
    """
    if isinstance(value, datetime):
        moment = value
    else:
        boundary = _DAY_END if end_of_day else _DAY_START
        moment = datetime.combine(value, boundary)
    if moment.tzinfo is None:
        return moment.replace(tzinfo=timezone.utc)
    return moment.astimezone(timezone.utc)


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
    if now is None:
        moment = datetime.now(timezone.utc)
    elif now.tzinfo is None:
        moment = now.replace(tzinfo=timezone.utc)
    else:
        moment = now.astimezone(timezone.utc)
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
        modified = datetime.fromtimestamp(stat.st_mtime, timezone.utc)
        extension = resolved.suffix.casefold()
        previous = (previous_samples or {}).get(relative)
        stable = _is_stable(stat.st_size, stat.st_mtime, moment, stability_window_seconds, previous)
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
            entries.append(replace(entry, disposition=disposition, reason=artifact, signal="name_pattern"))
            continue

        if extension not in READABLE_EXTENSIONS:
            entries.append(
                replace(entry, disposition=EXCLUDE, reason="unsupported_extension", signal="extension")
            )
            continue

        if not stable:
            # Never opened: an unstable file may be mid-write or mid-sync, so it
            # is not hashed and not classified from content.
            entries.append(
                replace(entry, disposition=REVIEW, reason="not_stable_yet", signal="stability_window")
            )
            continue

        accessed.append(str(resolved))
        content_hash = _hash_file(resolved)
        entry = replace(entry, content_hash=content_hash, version_key=version_key(resolved.name))

        classification, signal, note = _classify_stable(resolved, modified, period_window)
        notes = (note,) if note else ()
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
        entries.append(replace(entry, classification=classification, signal=signal, notes=notes))

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
