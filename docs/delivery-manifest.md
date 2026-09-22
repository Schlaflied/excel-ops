[中文](delivery-manifest.zh-CN.md) | English

# Delivery manifests: evidence for a completed delivery

A periodic delivery raises an obvious question afterward: *what exactly did
this run deliver?* `excel_ops.delivery_manifest` answers it with one
machine-readable and one human-readable evidence file per delivered workbook,
built entirely from a completed `run_delivery(...)` call and from the file it
actually wrote to disk.

```text
run_delivery(...) finishes and a workbook is really on disk
  -> for each delivered target, build one DeliveryManifest
       sources        content hash + record count + status breakdown per input
       tabs           per worksheet: written count vs. rows re-counted in the file
       template       content hash of the template actually used
       recipe         content-digest version of the Recipe decisions applied
       verification   #4's verdict, reduced to codes/severities/counts
       output         delivered filename + content hash
  -> write <output>.manifest.json  (machine-readable)
  -> write <output>.manifest.txt   (human-readable, same facts as lines)
```

A Manifest **computes nothing new**. Every number it reports either comes
straight out of the `DeliveryRun` that just finished, or is read back from the
persisted workbook after the write — never assumed, never carried over from an
in-memory counter alone.

## Why the row counts are read back, not trusted

The run already tracks how many records it believes it wrote per target
(`written`). A Manifest additionally **reopens the delivered file** and counts
real rows by their mapped record-ID column. When the two disagree — the file
was hand-edited after delivery, a row went missing, anything — `reconciled` is
`False` and the mismatch is named in `discrepancies` (for example
`row_count_mismatch:North` or `written_count_mismatch`) instead of being
smoothed over. This is what makes a Manifest falsifiable evidence rather than
a restated log line.

## What is in a Manifest

| Field | Meaning |
|---|---|
| `output`, `output_name`, `output_hash` | The delivered file's path, name, and content hash — hashed *after* the write and after verification published it. |
| `destination_key` | Which declared target this Manifest is for. |
| `template`, `template_version` | The template file used and its content hash. |
| `recipe_path`, `recipe_version` | The project Recipe used (if any) and a content digest of the decisions actually applied — see [Ambiguity Recipes](ambiguity-recipes.md). Re-saving the same answer with a new timestamp or reviewer does not move the version; only a changed decision does. |
| `sources[]` | One entry per input file that contributed: its name, path, content hash, records ingested from it, records written into *this* output, and a run/status breakdown. |
| `tabs[]` (keyed by sheet name in the JSON) | Per worksheet: `written` (the run's own count) vs. `rows` (re-counted from the real file), and which sources fed it. |
| `accepted`, `review`, `rejected`, `written`, `rows` | The run's terminal-state counts, narrowed to what is meaningful for one output — see below. |
| `run_counts`, `target_counts` | The full breakdown the two summary fields above are drawn from. |
| `verification` | `status` (`passed`/`failed`), `findings` (a count), `codes` (finding codes only), `severities` (counts by severity), and the path to the full verification report. |
| `period_start`, `period_end`, `period_display_text` | The resolved report period passed to `run_delivery(..., period=...)`, or `null` when none was declared — never guessed. |
| `run_fingerprint` | #19's whole-run fingerprint, when idempotency was enabled — a cross-reference only. |
| `reconciled`, `discrepancies` | Whether the file backs up every number above, and what does not reconcile when it doesn't. |

`accepted` and `rejected` are run-level: a rejected record never reached
matching, so it does not belong to any one destination. `review` is
attributed by *candidacy* — a record held back for review names every
destination it could still land on, so it counts there without claiming a
delivery that has not happened yet. `written` and the per-tab `rows` are the
only counts that are single-output *and* re-verified against the file.

## What is never in a Manifest

No cell value, and nothing derived from one, crosses into a Manifest:

- No location, identifier, category, or any other field value.
- No record ID — it is derived from source values, so it is excluded on the
  same grounds.
- No verification finding **message**. A finding message can quote the very
  cell that failed (for example `written_value_mismatch: expected 'X', got
  'Y'`); only the finding's **code** and **severity** cross over. The full
  message stays in the verification report `delivery_verification` already
  wrote, referenced by `report_path`.

What remains is exactly file names, file content hashes, sheet names,
declared field names, and integer counts.

## Using it

Manifests are built automatically; nothing needs to be enabled.

```python
from excel_ops import run_delivery

result = run_delivery(inputs, targets, staging_dir="staging", delivery_dir="delivery")

for manifest in result.manifests:
    print(manifest.output_name, manifest.reconciled, manifest.verification.status)
```

By default `run_delivery(...)` also writes `<output>.manifest.json` and
`<output>.manifest.txt` next to each delivered file. Pass
`write_manifest=False` (or `excel-ops deliver --no-manifest` on the CLI) to
keep the Manifests only in `result.manifests` — the evidence is still
computed and returned, just not persisted as sibling files.

```bash
excel-ops deliver delivery-plan.json --no-manifest
```

A run that delivers nothing — a dry run, a no-op, a target that failed
verification — produces no Manifest for that target: a Manifest describes a
file that is really on disk, never a plan or a failed attempt.

Read one back:

```python
from excel_ops.delivery_manifest import load_delivery_manifest, format_manifest

payload = load_delivery_manifest("delivery/North-Warehouse.xlsx.manifest.json")
```

`load_delivery_manifest` refuses anything that is not stamped
`excel-ops-delivery-manifest-v1`, the same idiom `idempotency.py` and the
Recipe loader already use for their own formats.

## Not the same as #4's verification or #19's idempotency fingerprint

These three systems sit next to each other in the pipeline and are easy to
conflate. They answer different questions and none of them replaces another:

| | Question it answers | Where it lives | Keyed by |
|---|---|---|---|
| **#4 verification** (`delivery_verification`) | "Does the reopened file actually match what should have been written?" Runs *before* a file counts as delivered — a failed check withholds delivery entirely. | `<output>.verification.json` | The delivery itself. |
| **#19 idempotency** (`idempotency.py`) | "Has this exact task already completed successfully, so this run can be skipped?" Runs *before* any ingest, match or write happens. | `.excel-ops/idempotency.json` | An opaque task-fingerprint digest — no path, name or cell value. |
| **#20 delivery manifest** (this module) | "What did this completed delivery actually contain?" Built *after* a file is delivered, as a record of it. | `<output>.manifest.json` / `.manifest.txt` | The delivered file it sits beside. |

A Manifest's `verification` block *reports* #4's verdict — it does not
re-verify anything. Its `run_fingerprint` field is a cross-reference to #19's
record when idempotency was enabled — the Manifest never reads or writes
`.excel-ops/idempotency.json` itself, and the two files are not
interchangeable: the fingerprint file is a component-hash digest with no
readable content, the Manifest is the human- and machine-readable evidence
about one delivery. Deleting or losing a Manifest never affects whether a
later run is treated as a no-op, and vice versa.

`delivery_manifest.py` is named to avoid a fourth collision: the unrelated
`refresh-manifest.json` tracked by `scripts/update-refresh-manifest.mjs` (see
[Refresh pipeline](refresh.md)) is a repository-tooling manifest of *tracked
source files*, with nothing to do with a spreadsheet delivery.

## Not implemented

- **Not a cross-run audit trail.** A Manifest describes exactly one delivery.
  Aggregating Manifests across a project's history is the later audit-package
  roadmap item (issue #12), not this one.
- **Not a second verification pass.** `verification` narrates #4's result; it
  never independently re-checks the file.
- **No signature.** A Manifest is evidence, not a tamper-proof attestation —
  nothing stops the sibling `.manifest.json` itself from being edited after
  the fact. What it protects against is *drift*: reopening the real delivered
  file and finding it disagrees with the run's own counters.
