"""Classifications, dispositions, and result types for the working-directory scan."""

from __future__ import annotations

import fnmatch
from dataclasses import asdict, dataclass
from datetime import date, datetime, time, timezone
from pathlib import Path, PurePosixPath
from typing import Any, Iterable


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
