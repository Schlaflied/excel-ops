[中文](idempotency.zh-CN.md) | English

# Idempotent execution: run fingerprints and the no-op short-circuit

A periodic task gets run again. Sometimes because the schedule fired, sometimes
because a person was not sure the last run finished, sometimes because a file was
re-downloaded under a new name. `excel_ops.idempotency` answers one question
before any work happens: *has this exact task already completed successfully?*

```text
compute the run fingerprint            (content hashes only, no filenames)
  -> load the persisted run record     (.excel-ops/idempotency.json)
  -> connector revision conflict check (extension point)
  -> decide: no_op | changed | retry
       no_op   -> return immediately, nothing was redone
       changed -> run the full pipeline
       retry   -> run the full pipeline (the last attempt did not finish)
```

This is **not** the same mechanism as the per-record deduplication already in
`run_delivery(...)`. That one compares stable record IDs against the record IDs a
delivered workbook already holds, so the same row is never appended twice — but
it only finds that out after ingesting, matching and planning. The fingerprint
sits above it and stops the whole run. Both are in force at the same time, and
neither replaces the other.

## Fingerprint scope

| Component | What it covers |
|---|---|
| `inputs` | the SHA-256 of every input file's **content** |
| `templates` | each template's content hash plus the declared Template Profile version marker |
| `mapping` | the declared sheet, column mapping, required fields, record-ID field and aliases — the "mapping version" |
| `recipe` | the reusable project Recipe decisions, taken from `ambiguity.py`'s already-parsed `RecipeDecision` objects rather than re-parsing the file |
| `confirmations` | the human confirmation records actually applied to this run |
| `period` | the declared report period, including the in-workbook period expectations |
| `connector` | the output target's connector identity |
| `output` | the content hash of whatever already exists at the target |
| `remote_revision` | the target's remote revision, for the conflict check |

Each component is stored as its own SHA-256 digest and the fingerprint is the
digest of that set, so a decision can name *which* component changed
(`changed_components`) without storing anything about it.

## Guarantees

- **Content, not names.** Inputs and templates are identified by content hash.
  A file that was renamed, moved or re-downloaded but is byte-identical produces
  the same fingerprint and does not trigger reprocessing; a file that kept its
  name but was edited does. Filenames, paths, sizes and modification times are
  never part of a fingerprint.
- **Listing order is not identity.** The per-file digests are sorted, so the same
  files declared in a different order are one task.
- **A failed run is never a completed baseline.** Only a record explicitly stored
  as `succeeded` can justify a `no_op`. A run that failed, was blocked by its
  plan, failed verification, or was interrupted between the write and the end is
  stored as `failed` / `started`, and a later identical run returns `retry` and
  proceeds. `retry` is a distinct decision from `changed` so the reason stays
  legible in a report.
- **A new human decision always re-runs.** Changing a confirmed review decision,
  adding one, or withdrawing one changes the `confirmations` component. Re-saving
  the *same* answer does not: a decision's `source` and `decided_at` are
  provenance, not rule content, so they are excluded.
- **The same period resolved on a later day is the same period.** A
  `PeriodResult` also carries `as_of_date`, which moves every day; including it
  would make a weekly task non-idempotent by the clock alone. Only the resolved
  window and its display text are fingerprinted.
- **Editing the delivered file re-runs.** The output content hash is in scope, so
  a delivery that was hand-edited or deleted after the fact is a change.
- **Nothing raw is persisted.** Every component is a digest. A run record
  contains no path, no destination name, no cell value, no cloud identifier and
  no credential. The free-form `detail` is scrubbed the same way: nested
  structures are reduced to their type name and any value whose key looks like a
  secret (`token`, `password`, `authorization`, …) is replaced by a digest. A
  credential can still be *part of the task identity* through
  `ConnectorTarget.credential_material`, which is hashed into the connector
  component and never written out.
- **A corrupt or foreign run-state file never authorizes a no-op.** The loader
  refuses anything that is not `excel-ops-run-record-v1`; `run_delivery(...)`
  then proceeds with reason `unreadable_run_state` and rewrites the record.

## Where the run record lives

```text
<delivery_dir>/.excel-ops/idempotency.json
```

This mirrors `refresh.mjs`'s `.refresh/state.json` idiom: one hidden state
directory next to the work it describes, one small versioned JSON file inside it,
a `format` sentinel that refuses anything else. The file is keyed by task, so
several periodic tasks can share one delivery directory without overwriting each
other's record. Task keys default to an opaque digest of the output target; pass
`task_key=` to name one yourself.

```json
{
  "format": "excel-ops-run-record-v1",
  "runs": {
    "weekly-north": {
      "fingerprint": "…64 hex…",
      "status": "succeeded",
      "components": {"inputs": "…", "templates": "…", "…": "…"},
      "recorded_at": "2026-09-21T09:30:00+00:00",
      "delivered": true,
      "attempt": 1,
      "detail": {"counts": {"written": 2}, "failure_codes": [], "targets": 1}
    }
  }
}
```

The record is written **twice** per run: once as `started` before the first
write, so an interrupted run is found and retried, and once at the end with the
final status. The final record stores the fingerprint recomputed *after* the
delivery, so it already accounts for the output the run produced — without that,
the output content hash could never match on a rerun.

## Using it

Whole-run idempotency is opt-in. A caller that does not pass `idempotency=`
behaves exactly as before, per-record deduplication included.

```python
from excel_ops import IdempotencyOptions, run_delivery

options = IdempotencyOptions(task_key="weekly-north", template_profile_version="profile-v1")

first = run_delivery(inputs, targets, staging_dir=..., delivery_dir=..., idempotency=options)
second = run_delivery(inputs, targets, staging_dir=..., delivery_dir=..., idempotency=options)

assert second.no_op is True
assert second.run_decision.decision == "no_op"
print(second.run_decision.explain())
# no_op: unchanged_since_successful_run (fingerprint 5af5cf177346)
```

On the CLI:

```bash
excel-ops deliver delivery-plan.json --run-state --task-key weekly-north
excel-ops deliver delivery-plan.json --run-state runs/state.json --task-key weekly-north
```

A no-op exits `0`: the declared delivery is already in place, nothing changed,
nothing failed. `no_op` is reported as its own field rather than as
`delivered: true`, because this run delivered nothing — the previous one did.

A **dry run never short-circuits.** `plan_delivery(...)` still returns the full,
checkable plan; it just also reports the verdict in `run_decision`.

## The connector conflict interface

The issue's last acceptance criterion is that a **cloud target revision change
triggers a conflict check**. There is no cloud connector in this repository yet —
that is Phase 2 — so what ships here is the interface a cloud connector plugs
into, with test coverage against a stub. It is a forward-compatible extension
point, **not a working cloud integration**.

```python
class RevisionSource(Protocol):
    def current_revision(self) -> str | None: ...
```

`check_connector_conflict(previous_record, connector, revision_source=...)`
compares the target's current revision with the one recorded for the last run and
returns a `ConnectorConflict(code="remote_revision_changed", …)` when they
differ. `evaluate_run(...)` calls it first, and a conflict can never resolve to
`no_op`: the decision becomes `changed` with the conflict attached, whatever the
rest of the fingerprint says. An unknown current revision (`None`) is never
treated as "unchanged" once a revision has been recorded.

For Phase 1 the target is local and `ConnectorTarget.local([output_path])`
describes it as `local:<output path>`, POSIX-normalised so a task fingerprinted
on Windows matches the same task elsewhere. A future cloud connector supplies its
own stable target id and its own `RevisionSource` without the fingerprint format
changing.

## Not implemented

- **No cloud connector.** The conflict check is an interface plus a stub-backed
  test. Nothing here talks to a cloud spreadsheet, and nothing here decides *how*
  to reconcile a conflict — it reports that one exists.
- **No sub-run granularity.** The decision is for the whole run. When something
  changed, the run proceeds in full and the existing per-record deduplication
  decides what is actually appended; this module does not deliver "only the
  changed records".
- **Not wired to the working-directory scan.** `scan_workdir(...)` (#15) decides
  *which* files are this period's inputs; this module takes the input list as an
  argument and never scans or classifies files itself.
- **A recorded success is trusted.** The module does not re-verify a previously
  delivered file beyond its content hash. A delivery that was edited and then
  edited back to byte-identical content is, correctly, indistinguishable from an
  untouched one.
- **No lock.** The run record is state, not a mutex. Two runs started at the same
  moment can both see "no previous success"; concurrency control is out of scope.
