from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Mapping, Protocol


class PeriodLike(Protocol):
    period_start: date
    period_end: date
    as_of_date: date


class OutputNamingError(ValueError):
    """Raised when an output naming rule cannot be resolved safely."""


_BRACED_TOKEN = re.compile(r"\{([A-Za-z][A-Za-z0-9_-]*)\}")
_RESERVED_WINDOWS_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}
_ILLEGAL_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass(frozen=True)
class ResolvedOutput:
    original_pattern: str
    path: Path
    collision: str
    tokens: Mapping[str, str]

    def to_dict(self) -> dict[str, object]:
        return {
            "original_pattern": self.original_pattern,
            "resolved_path": str(self.path),
            "collision": self.collision,
            "tokens": dict(self.tokens),
        }


def resolve_output_path(
    pattern: str,
    *,
    output_dir: str | Path,
    extension: str,
    period: PeriodLike,
    values: Mapping[str, object] | None = None,
    directory_mode: bool = False,
    content: bytes | None = None,
) -> ResolvedOutput:
    """Resolve a safe, deterministic delivery path from an already parsed period.

    ``directory_mode`` is intentionally opt-in. Otherwise slashes in a prompt such
    as ``payroll-YYYY/MM/DD`` become filename separators (hyphens), not folders.
    Existing identical content is a no-op; differing content receives ``-rN``.
    """

    if not pattern or not pattern.strip():
        raise OutputNamingError("output naming pattern must not be empty")

    suffix = _normalise_extension(extension)
    tokens = _build_tokens(period, values or {})
    rendered = _render_pattern(pattern.strip(), tokens)
    parts = re.split(r"[/\\]+", rendered) if directory_mode else [rendered]
    safe_parts = [_clean_component(part) for part in parts]
    if any(not part for part in safe_parts):
        raise OutputNamingError("output naming pattern contains an empty path component")

    leaf = safe_parts[-1]
    if Path(leaf).suffix.lower() != suffix.lower():
        leaf = f"{leaf}{suffix}"
    safe_parts[-1] = leaf

    root = Path(output_dir)
    candidate = root.joinpath(*safe_parts)
    collision = "new"
    if candidate.exists():
        if content is not None and _same_content(candidate, content):
            collision = "identical_noop"
        else:
            candidate = _next_revision(candidate)
            collision = "revision"

    return ResolvedOutput(pattern, candidate, collision, tokens)


def verify_output_name(
    result: ResolvedOutput,
    *,
    period: PeriodLike,
    manifest: Mapping[str, object] | None = None,
) -> None:
    """Verify the expected file exists and its recorded date tokens still agree."""

    if not result.path.is_file():
        raise OutputNamingError(f"resolved output was not written: {result.path}")
    expected = _build_tokens(period, {})
    for key in ("YYYY", "MM", "DD", "period_start", "period_end"):
        recorded = result.tokens.get(key)
        if recorded is not None and recorded != expected[key]:
            raise OutputNamingError(f"output token {key} does not match the resolved period")
    if manifest is not None:
        for key in ("period_start", "period_end"):
            if str(manifest.get(key, "")) != expected[key]:
                raise OutputNamingError(f"manifest {key} does not match the resolved period")


def _build_tokens(period: PeriodLike, values: Mapping[str, object]) -> dict[str, str]:
    delivery_date = period.as_of_date
    tokens = {
        "YYYY": f"{delivery_date.year:04d}",
        "MM": f"{delivery_date.month:02d}",
        "DD": f"{delivery_date.day:02d}",
        "period_start": period.period_start.isoformat(),
        "period_end": period.period_end.isoformat(),
    }
    aliases = {
        "client": "client",
        "customer": "client",
        "department": "department",
        "workflow": "workflow",
        "workflow_name": "workflow",
    }
    for supplied, value in values.items():
        canonical = aliases.get(str(supplied), str(supplied))
        tokens[canonical] = str(value).strip()
    return tokens


def _render_pattern(pattern: str, tokens: Mapping[str, str]) -> str:
    def replace_braced(match: re.Match[str]) -> str:
        token = match.group(1)
        if token in tokens:
            return tokens[token]
        raise OutputNamingError(f"unknown or unavailable output token: {token}")

    rendered = _BRACED_TOKEN.sub(replace_braced, pattern)
    # Bare date tokens are supported for the common payroll-YYYY/MM/DD form.
    # Longer tokens are braced so ordinary words such as PAYROLL are untouched.
    for token in ("YYYY", "MM", "DD"):
        rendered = rendered.replace(token, tokens[token])
    return rendered


def _clean_component(value: str) -> str:
    cleaned = _ILLEGAL_FILENAME_CHARS.sub("-", value).strip().rstrip(". ")
    cleaned = re.sub(r"-+", "-", cleaned)
    if cleaned.upper() in _RESERVED_WINDOWS_NAMES:
        cleaned = f"_{cleaned}"
    return cleaned


def _normalise_extension(extension: str) -> str:
    extension = extension.strip()
    if not extension:
        raise OutputNamingError("delivery extension must not be empty")
    if any(character in extension for character in "/\\"):
        raise OutputNamingError("delivery extension must not contain a path separator")
    return extension if extension.startswith(".") else f".{extension}"


def _same_content(path: Path, content: bytes) -> bool:
    existing = hashlib.sha256(path.read_bytes()).digest()
    proposed = hashlib.sha256(content).digest()
    return existing == proposed


def _next_revision(path: Path) -> Path:
    revision = 2
    while True:
        candidate = path.with_name(f"{path.stem}-r{revision}{path.suffix}")
        if not candidate.exists():
            return candidate
        revision += 1
