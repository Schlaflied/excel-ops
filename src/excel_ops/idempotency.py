"""Whole-run idempotency for periodic delivery tasks.

This module answers one question *before* a delivery run does any work: *has
this exact task already completed successfully?*  It fingerprints everything
that can change the outcome of a run, compares that fingerprint with a persisted
run record, and returns one of three decisions:

* ``no_op`` -- the fingerprint matches a previously **successful** run, so there
  is nothing to do and the run short-circuits;
* ``changed`` -- something in the fingerprint scope differs (or there is no
  previous record at all), so the run proceeds;
* ``retry`` -- the fingerprint matches a previous run that **failed** or was
  never confirmed successful.  The run always proceeds: a failed run is never a
  completed no-op baseline.

This is deliberately a *different* mechanism from the per-record deduplication
already in :mod:`excel_ops.delivery` (``_existing_record_ids`` /
``_build_plan``).  That one stops an individual row being appended twice inside
a target workbook, and it only finds out at plan time, after matching.  This one
sits above it and stops the whole run.  Both stay in force.

Fingerprint scope (issue #19):

===========================  ===================================================
Component                    What it covers
===========================  ===================================================
``inputs``                   the SHA-256 of every input file's **content**, so a
                             renamed or moved but byte-identical file produces
                             the same fingerprint, and a same-named but edited
                             file does not
``templates``               each template's content hash plus the declared
                             Template Profile version marker
``mapping``                 the declared sheet/column mapping and target
                             contract -- the "mapping version"
``recipe``                  the reusable project Recipe decisions, taken from
                             :mod:`excel_ops.ambiguity`'s already-parsed
                             ``RecipeDecision`` objects rather than re-parsing
                             the file independently
``confirmations``           the human confirmation records actually applied to
                             this run
``period``                  the declared report period
``connector``               the output target's connector identity
``output``                  the content hash of the output that already exists
                             at the target, so editing or deleting a delivered
                             file is a change
``remote_revision``         the target's remote revision, for the conflict check
===========================  ===================================================

Nothing raw is persisted.  Every component of a fingerprint is stored as a
SHA-256 digest, so a run record never contains a path, a cell value, a cloud
identifier, or a credential.  The free-form ``detail`` on a run record is
scrubbed the same way: nested structures are reduced to their type name and any
value whose key looks like a secret is replaced by a digest.

**Forward compatibility.** Phase 1 is local only; there is no cloud connector in
this repository yet.  :class:`ConnectorTarget` therefore describes a local target
as ``local:<output path>`` and the remote-revision conflict check is exposed as
an extension point (:class:`RevisionSource`, :func:`check_connector_conflict`) a
future cloud connector plugs into.  It is an interface with test coverage, not a
working cloud integration.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Mapping, Protocol, Sequence, runtime_checkable


FINGERPRINT_FORMAT = "excel-ops-run-fingerprint-v1"
RUN_RECORD_FORMAT = "excel-ops-run-record-v1"

#: Mirrors ``refresh.mjs``'s ``.refresh/state.json`` idiom: one hidden state
#: directory next to the work it describes, one small JSON state file inside it.
STATE_DIRNAME = ".excel-ops"
STATE_FILENAME = "idempotency.json"

NO_OP = "no_op"
CHANGED = "changed"
RETRY = "retry"
DECISIONS = (NO_OP, CHANGED, RETRY)

STARTED = "started"
SUCCEEDED = "succeeded"
FAILED = "failed"
RUN_STATUSES = (STARTED, SUCCEEDED, FAILED)

COMPONENT_INPUTS = "inputs"
COMPONENT_TEMPLATES = "templates"
COMPONENT_MAPPING = "mapping"
COMPONENT_RECIPE = "recipe"
COMPONENT_CONFIRMATIONS = "confirmations"
COMPONENT_PERIOD = "period"
COMPONENT_CONNECTOR = "connector"
COMPONENT_OUTPUT = "output"
COMPONENT_REMOTE_REVISION = "remote_revision"
COMPONENTS = (
    COMPONENT_INPUTS,
    COMPONENT_TEMPLATES,
    COMPONENT_MAPPING,
    COMPONENT_RECIPE,
    COMPONENT_CONFIRMATIONS,
    COMPONENT_PERIOD,
    COMPONENT_CONNECTOR,
    COMPONENT_OUTPUT,
    COMPONENT_REMOTE_REVISION,
)

#: A file that cannot be read has no content hash.  A run over an unreadable
#: input fails, and a failed run is never a no-op baseline, so this sentinel can
#: never make two different unreadable files look like a completed run.
UNREADABLE = "unreadable"

#: Key fragments whose value is replaced by a digest before it is persisted.
_SECRET_HINTS = (
    "token",
    "secret",
    "password",
    "passwd",
    "credential",
    "apikey",
    "api_key",
    "authorization",
    "auth",
    "cookie",
    "session",
    "signature",
    "private",
)

_HASH_CHUNK = 1 << 20
_DETAIL_STRING_LIMIT = 200


class IdempotencyError(ValueError):
    """Raised when a run record cannot be read or written without guessing."""


def _canonical(value: Any) -> str:
    """Serialize ``value`` so equal meaning always yields equal bytes."""

    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)


def _digest(value: Any) -> str:
    return hashlib.sha256(_canonical(value).encode("utf-8")).hexdigest()


def file_content_digest(path: str | Path) -> str:
    """Hash a file's **content**, never its name, size, or modification time."""

    target = Path(path)
    hasher = hashlib.sha256()
    try:
        with target.open("rb") as handle:
            while True:
                chunk = handle.read(_HASH_CHUNK)
                if not chunk:
                    break
                hasher.update(chunk)
    except OSError:
        return UNREADABLE
    return f"sha256:{hasher.hexdigest()}"


def _content_digests(paths: Iterable[str | Path]) -> list[str]:
    """Return the content digests of ``paths``, sorted and independent of order.

    Sorting is what makes "the same files listed in a different order" one
    fingerprint: the inputs are identified by content, so their listing order is
    not part of the task's identity.
    """

    return sorted(file_content_digest(item) for item in paths)


@runtime_checkable
class RevisionSource(Protocol):
    """A target that can report its own current revision.

    The extension point a future cloud connector implements.  ``None`` means the
    revision is unknown, which is never treated as "unchanged".
    """

    def current_revision(self) -> str | None:  # pragma: no cover - protocol
        ...


@dataclass(frozen=True)
class ConnectorTarget:
    """Where a run delivers, described without persisting anything raw.

    ``kind`` is ``local`` for Phase 1.  ``identifier`` is a stable, opaque name
    for the destination (``local:<output path>``); a future cloud connector uses
    the same field for its own stable target id.  ``credential_material`` exists
    so a connector that is only reachable with a given credential can make that
    part of the task's identity *without* the credential ever being stored: it is
    hashed into the connector component and never written out.
    """

    identifier: str
    kind: str = "local"
    outputs: tuple[str, ...] = ()
    revision: str | None = None
    credential_material: tuple[str, ...] = field(default=(), repr=False)

    @classmethod
    def local(
        cls,
        outputs: Iterable[str | Path],
        *,
        identifier: str | None = None,
        revision: str | None = None,
    ) -> "ConnectorTarget":
        """Describe a Phase-1 local output target.

        Paths are normalised to POSIX separators so a task fingerprinted on
        Windows matches the same task elsewhere.
        """

        listed = tuple(sorted(Path(item).as_posix() for item in outputs))
        name = identifier or f"local:{listed[0] if listed else ''}"
        return cls(identifier=name, kind="local", outputs=listed, revision=revision)

    def identity(self) -> dict[str, Any]:
        return {
            "kind": self.kind,
            "identifier": self.identifier,
            "outputs": list(self.outputs),
            "credential": _digest(list(self.credential_material)) if self.credential_material else None,
        }

    def current_revision(self, revision_source: RevisionSource | None = None) -> str | None:
        """Prefer a live revision from the connector over the declared one."""

        if revision_source is not None:
            return revision_source.current_revision()
        return self.revision


@dataclass(frozen=True)
class ConnectorConflict:
    """A remote revision that moved under a previously recorded run."""

    code: str
    message: str
    expected_revision: str | None
    actual_revision: str | None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class RunFingerprint:
    """One run's identity: a digest per component plus their combined digest."""

    digest: str
    components: Mapping[str, str]
    format: str = FINGERPRINT_FORMAT

    def changed_components(self, other: Mapping[str, str] | None) -> tuple[str, ...]:
        """Name the components that differ from ``other``, for an explanation."""

        if other is None:
            return ()
        names = sorted(set(self.components) | set(other))
        return tuple(name for name in names if self.components.get(name) != other.get(name))

    def to_dict(self) -> dict[str, Any]:
        return {
            "format": self.format,
            "digest": self.digest,
            "components": dict(self.components),
        }


def _decision_payload(item: Any) -> dict[str, Any]:
    """Canonicalize one human decision, provenance excluded.

    A Recipe decision's ``source`` and ``decided_at`` say *who* recorded it and
    *when*; re-saving the same answer must not look like a rule change.  What
    matters is the question, the answer, its scope, and the candidate set the
    answer was given against.
    """

    if isinstance(item, Mapping):
        read = item.get
    else:
        def read(name: str, default: Any = None) -> Any:
            return getattr(item, name, default)
    return {
        "field": str(read("field", "") or ""),
        "question": str(read("question", "") or ""),
        "selected": str(read("selected", "") or ""),
        "scope": str(read("scope", "") or ""),
        "candidates": sorted(str(value) for value in (read("candidates", ()) or ())),
        "custom": bool(read("custom", False)),
    }


def _decisions_payload(decisions: Iterable[Any] | Mapping[str, Any] | None) -> list[dict[str, Any]]:
    if decisions is None:
        return []
    items = decisions.values() if isinstance(decisions, Mapping) else decisions
    return sorted(
        (_decision_payload(item) for item in items),
        key=lambda payload: (payload["field"], payload["question"], payload["selected"]),
    )


def _period_payload(period: Any) -> Any:
    """Canonicalize a declared report period.

    Only the resolved window and its display text are taken.  A
    :class:`~excel_ops.periods.PeriodResult` also carries ``as_of_date``, which
    moves every day even when the resolved period is identical; including it
    would make a weekly task non-idempotent by the clock alone.
    """

    if period is None:
        return None
    start = getattr(period, "period_start", None) or getattr(period, "start", None)
    end = getattr(period, "period_end", None) or getattr(period, "end", None)
    if start is None and end is None:
        if isinstance(period, Mapping):
            return {key: str(value) for key, value in sorted(period.items())}
        return str(period)
    return {
        "start": str(start),
        "end": str(end),
        "display_text": str(getattr(period, "display_text", "") or ""),
    }


def compute_fingerprint(
    *,
    inputs: Iterable[str | Path] = (),
    templates: Iterable[str | Path] = (),
    template_profile_version: str | None = None,
    mapping: Any = None,
    recipe_decisions: Iterable[Any] | Mapping[str, Any] | None = None,
    confirmations: Iterable[Any] | Mapping[str, Any] | None = None,
    period: Any = None,
    connector: ConnectorTarget | None = None,
    revision_source: RevisionSource | None = None,
) -> RunFingerprint:
    """Fingerprint everything that can change a run's outcome.

    Every argument is optional so a caller can fingerprint the scope it actually
    declares; an absent component still gets a stable digest, so adding a period
    or a connector later is itself a change.
    """

    outputs = connector.outputs if connector is not None else ()
    revision = connector.current_revision(revision_source) if connector is not None else None
    components = {
        COMPONENT_INPUTS: _digest(_content_digests(inputs)),
        COMPONENT_TEMPLATES: _digest(
            {
                "contents": _content_digests(templates),
                "profile_version": template_profile_version,
            }
        ),
        COMPONENT_MAPPING: _digest(mapping),
        COMPONENT_RECIPE: _digest(_decisions_payload(recipe_decisions)),
        COMPONENT_CONFIRMATIONS: _digest(_decisions_payload(confirmations)),
        COMPONENT_PERIOD: _digest(_period_payload(period)),
        COMPONENT_CONNECTOR: _digest(connector.identity() if connector is not None else None),
        COMPONENT_OUTPUT: _digest(_content_digests(outputs)),
        COMPONENT_REMOTE_REVISION: _digest(revision),
    }
    combined = _digest({"format": FINGERPRINT_FORMAT, "components": components})
    return RunFingerprint(combined, components)


def _scrub(value: Any, *, secret: bool = False) -> Any:
    """Reduce one detail value to something safe to persist."""

    if secret:
        return f"sha256:{_digest(value)}"
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        return value[:_DETAIL_STRING_LIMIT]
    if isinstance(value, Mapping):
        return {str(key): _scrub(item, secret=_is_secret(str(key))) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_scrub(item) for item in value]
    return type(value).__name__


def _is_secret(name: str) -> bool:
    lowered = name.casefold()
    return any(hint in lowered for hint in _SECRET_HINTS)


def scrub_detail(detail: Mapping[str, Any] | None) -> dict[str, Any]:
    """Return ``detail`` with anything credential-shaped replaced by a digest."""

    if not detail:
        return {}
    return {str(key): _scrub(value, secret=_is_secret(str(key))) for key, value in detail.items()}


@dataclass(frozen=True)
class RunRecord:
    """What the last run of one task did, and whether it truly completed."""

    task_key: str
    fingerprint: str
    status: str
    components: Mapping[str, str] = field(default_factory=dict)
    recorded_at: str = ""
    delivered: bool = False
    attempt: int = 1
    detail: Mapping[str, Any] = field(default_factory=dict)
    format: str = RUN_RECORD_FORMAT

    def __post_init__(self) -> None:
        if self.status not in RUN_STATUSES:
            raise IdempotencyError(f"unknown run status: {self.status}")
        if not str(self.task_key).strip():
            raise IdempotencyError("a run record needs a task key")
        if not str(self.fingerprint).strip():
            raise IdempotencyError("a run record needs a fingerprint")

    @property
    def successful(self) -> bool:
        """Only an explicitly recorded success can ever justify a no-op."""

        return self.status == SUCCEEDED

    def to_dict(self) -> dict[str, Any]:
        return {
            "format": self.format,
            "task_key": self.task_key,
            "fingerprint": self.fingerprint,
            "status": self.status,
            "components": dict(self.components),
            "recorded_at": self.recorded_at,
            "delivered": self.delivered,
            "attempt": self.attempt,
            "detail": dict(self.detail),
        }


def default_task_key(connector: ConnectorTarget | None, *, extra: Any = None) -> str:
    """Derive a stable, opaque task key from the output target.

    The key is a digest, so the state file names no path and no cloud
    identifier even when several tasks share one state file.
    """

    return _digest(
        {
            "connector": connector.identity() if connector is not None else None,
            "extra": extra,
        }
    )


def state_path(base_dir: str | Path) -> Path:
    """Return the conventional run-state file for a delivery directory."""

    return Path(base_dir) / STATE_DIRNAME / STATE_FILENAME


def load_run_records(path: str | Path) -> dict[str, RunRecord]:
    """Load every task's run record, refusing any other format."""

    target = Path(path)
    if not target.is_file():
        return {}
    try:
        payload = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise IdempotencyError(f"{target.name}: unreadable run state ({error})") from error
    if not isinstance(payload, Mapping) or payload.get("format") != RUN_RECORD_FORMAT:
        raise IdempotencyError(f"{target.name}: unsupported run-state format")
    records: dict[str, RunRecord] = {}
    raw_runs = payload.get("runs")
    if not isinstance(raw_runs, Mapping):
        raise IdempotencyError(f"{target.name}: run state has no runs object")
    for key, raw in raw_runs.items():
        if not isinstance(raw, Mapping):
            raise IdempotencyError(f"{target.name}: run record {key} is not an object")
        record = RunRecord(
            task_key=str(raw.get("task_key") or key),
            fingerprint=str(raw.get("fingerprint") or ""),
            status=str(raw.get("status") or ""),
            components={str(name): str(value) for name, value in (raw.get("components") or {}).items()},
            recorded_at=str(raw.get("recorded_at") or ""),
            delivered=bool(raw.get("delivered", False)),
            attempt=int(raw.get("attempt", 1) or 1),
            detail=dict(raw.get("detail") or {}),
        )
        records[record.task_key] = record
    return records


def load_run_record(path: str | Path, task_key: str) -> RunRecord | None:
    """Load one task's record, or ``None`` when this task has never run."""

    return load_run_records(path).get(task_key)


def save_run_record(record: RunRecord, path: str | Path) -> Path:
    """Merge one task's record into the state file without losing the others."""

    target = Path(path)
    existing = load_run_records(target) if target.is_file() else {}
    existing[record.task_key] = record
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "format": RUN_RECORD_FORMAT,
        "runs": {key: item.to_dict() for key, item in sorted(existing.items())},
    }
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return target


def record_run(
    fingerprint: RunFingerprint,
    path: str | Path,
    *,
    task_key: str,
    status: str,
    delivered: bool = False,
    detail: Mapping[str, Any] | None = None,
    previous: RunRecord | None = None,
    recorded_at: str | None = None,
) -> RunRecord:
    """Persist the outcome of one run and return the stored record."""

    attempt = 1
    if previous is not None and previous.fingerprint == fingerprint.digest:
        attempt = previous.attempt + 1
    record = RunRecord(
        task_key=task_key,
        fingerprint=fingerprint.digest,
        status=status,
        components=dict(fingerprint.components),
        recorded_at=recorded_at or datetime.now(timezone.utc).isoformat(),
        delivered=delivered,
        attempt=attempt,
        detail=scrub_detail(detail),
    )
    save_run_record(record, path)
    return record


def check_connector_conflict(
    previous: RunRecord | None,
    connector: ConnectorTarget | None = None,
    *,
    revision_source: RevisionSource | None = None,
) -> ConnectorConflict | None:
    """Compare the target's current revision with the recorded one.

    This is the forward-compatible hook a cloud connector calls into.  A moved
    revision means the delivered artifact was changed by something other than
    this pipeline, so the run must never short-circuit and the caller must
    reconcile before overwriting.  An unknown current revision (``None``) is
    never treated as "unchanged" once a revision was recorded.
    """

    if previous is None or connector is None:
        return None
    recorded = previous.components.get(COMPONENT_REMOTE_REVISION)
    if recorded is None:
        return None
    current = connector.current_revision(revision_source)
    actual = _digest(current)
    if actual == recorded:
        return None
    return ConnectorConflict(
        "remote_revision_changed",
        (
            f"the target {connector.identifier} reports a different revision than the "
            "one recorded for the last run"
        ),
        recorded,
        actual,
    )


@dataclass(frozen=True)
class RunDecision:
    """Whether this run may short-circuit, and the reason either way."""

    decision: str
    reason: str
    fingerprint: RunFingerprint
    previous: RunRecord | None = None
    changed_components: tuple[str, ...] = ()
    conflict: ConnectorConflict | None = None

    def __post_init__(self) -> None:
        if self.decision not in DECISIONS:
            raise IdempotencyError(f"unknown run decision: {self.decision}")

    @property
    def no_op(self) -> bool:
        return self.decision == NO_OP

    @property
    def should_run(self) -> bool:
        return self.decision != NO_OP

    def explain(self) -> str:
        detail = f"{self.decision}: {self.reason} (fingerprint {self.fingerprint.digest[:12]})"
        if self.changed_components:
            detail += f"; changed: {', '.join(self.changed_components)}"
        if self.conflict is not None:
            detail += f"; conflict: {self.conflict.code}"
        return detail

    def to_dict(self) -> dict[str, Any]:
        return {
            "decision": self.decision,
            "reason": self.reason,
            "explanation": self.explain(),
            "fingerprint": self.fingerprint.to_dict(),
            "changed_components": list(self.changed_components),
            "previous_status": self.previous.status if self.previous else None,
            "previous_recorded_at": self.previous.recorded_at if self.previous else None,
            "previous_attempt": self.previous.attempt if self.previous else None,
            "conflict": self.conflict.to_dict() if self.conflict else None,
        }


def decide_run(
    fingerprint: RunFingerprint,
    previous: RunRecord | None,
    *,
    conflict: ConnectorConflict | None = None,
) -> RunDecision:
    """Return ``no_op``, ``changed`` or ``retry`` for this fingerprint."""

    if conflict is not None:
        # A moved remote revision can never short-circuit, whatever the
        # fingerprint says.  The conflict is surfaced so the caller reconciles.
        return RunDecision(
            CHANGED,
            conflict.code,
            fingerprint,
            previous,
            fingerprint.changed_components(previous.components if previous else None),
            conflict,
        )
    if previous is None:
        return RunDecision(CHANGED, "no_previous_run", fingerprint)
    if previous.fingerprint != fingerprint.digest:
        return RunDecision(
            CHANGED,
            "fingerprint_changed",
            fingerprint,
            previous,
            fingerprint.changed_components(previous.components),
        )
    if not previous.successful:
        # Same task, but the last attempt did not finish successfully.  It is a
        # retry, never a completed baseline.
        return RunDecision(RETRY, f"previous_run_{previous.status}", fingerprint, previous)
    return RunDecision(NO_OP, "unchanged_since_successful_run", fingerprint, previous)


def evaluate_run(
    fingerprint: RunFingerprint,
    previous: RunRecord | None,
    *,
    connector: ConnectorTarget | None = None,
    revision_source: RevisionSource | None = None,
) -> RunDecision:
    """Run the connector conflict check, then decide."""

    conflict = check_connector_conflict(previous, connector, revision_source=revision_source)
    return decide_run(fingerprint, previous, conflict=conflict)


@dataclass(frozen=True)
class IdempotencyOptions:
    """Opt-in whole-run idempotency for one periodic delivery task.

    Passing this to :func:`excel_ops.delivery.run_delivery` turns on the
    short-circuit.  Callers that do not pass it keep the previous behaviour
    exactly, including the independent per-record deduplication.
    """

    state_path: str | Path | None = None
    task_key: str | None = None
    template_profile_version: str | None = None
    period: Any = None
    connector: ConnectorTarget | None = None
    revision_source: RevisionSource | None = None

    def resolved_state_path(self, base_dir: str | Path) -> Path:
        return Path(self.state_path) if self.state_path is not None else state_path(base_dir)

    def resolved_task_key(self, connector: ConnectorTarget | None) -> str:
        if self.task_key:
            return self.task_key
        return default_task_key(connector or self.connector)


def format_decision(decision: RunDecision) -> str:
    """Render one readable line for a conversation or a CLI report."""

    lines = [f"excel-ops run decision: {decision.explain()}"]
    if decision.previous is not None:
        lines.append(
            f"previous run: status={decision.previous.status} "
            f"delivered={decision.previous.delivered} at={decision.previous.recorded_at or 'unknown'}"
        )
    if decision.conflict is not None:
        lines.append(f"conflict: {decision.conflict.message}")
    return "\n".join(lines)


def fingerprint_paths(paths: Sequence[str | Path]) -> tuple[str, ...]:
    """Expose the per-file content digests, for reporting and tests."""

    return tuple(_content_digests(paths))


def recipe_version_digest(decisions: Iterable[Any] | Mapping[str, Any] | None) -> str:
    """Version the reusable Recipe decisions by their content.

    This is the ``recipe`` fingerprint component (see :func:`compute_fingerprint`)
    exposed on its own, so a report can record *which* rule version produced an
    artifact without recomputing a whole run fingerprint or reimplementing the
    canonicalization.  Like the component, it excludes a decision's ``source``
    and ``decided_at``: re-saving the same answer is not a rule change.

    The returned value is ``sha256:<hex>``; the fingerprint component is the same
    hex without the prefix, so the two remain directly comparable.
    """

    return f"sha256:{_digest(_decisions_payload(decisions))}"
