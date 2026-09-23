<p align="center">
  <img src="assets/logo.png" alt="Excel-Ops logo" width="180">
</p>

# Excel-Ops

[中文说明](README.zh-CN.md) | English

### I hate VLOOKUP. This is why.

Excel-Ops is a conversation-driven workflow for AI agents, not another spreadsheet web app, and definitely not an AI formula generator. Give an agent the files and describe the outcome in plain language; the agent plans the operation, picks the right local or cloud connector, isolates whatever it isn't sure about, verifies the result, and hands back a delivery package — not a guess dressed up as an answer.

The reason VLOOKUP is annoying was never the formula. It's that real spreadsheets have inconsistent names, duplicate records, shifting headers, and dates that don't agree with each other. This design comes from a real, recurring reporting workflow — reconciling messy multi-source records by hand, week after week — not a hypothetical use case dreamed up for a demo. Excel-Ops treats that mess as an operational workflow instead of pretending a single function can paper over it:

**ingest → extract → normalize → match → human review → write → verify → deliver**

It's a companion project to [PPT-Ops](https://github.com/luochen211/ppt-ops): PPT-Ops turns source material into a reviewable presentation delivery, Excel-Ops turns images and messy tabular inputs into a reviewable data delivery.

Excel-Ops's verification-first core — data validation before any claim of success — and `refresh.mjs`'s own update-integrity design both trace back to patterns the same author first built for Career-Ops, a separate personal job-search automation project. Career-Ops's "Jurisdiction Umbrella" concept and Excel-Ops share that same lineage and the same author, even though the two codebases share no code or data.

## Use with an agent

In Codex or another agent that follows the Agent Skills standard, invoke the repository skill:

> Use `$excel-agent` to extract the rows from these screenshots, append only high-confidence records to the workbook, and put everything uncertain in a review sheet.

`$excel-agent` is the only user-facing entry point. The Python CLI remains an internal deterministic capability used by the agent and by automated tests.

Agents that support MCP can launch the local `excel-ops-mcp` stdio server. It exposes working-directory scan, validated delivery-plan preparation, read-only dry runs, and explicitly confirmed delivery while keeping all spreadsheet decisions in the existing Python core. See [Agent MCP server](docs/mcp.md).

The initial connector order is local XLSX/CSV, local cloud-sync folders, Google Sheets, Dropbox API, Feishu Sheets, and WPS Sheets. Cloud credentials belong to the user's environment and are never stored in a project or delivery package.

The first vertical slice focuses on a common business task: ingest XLSX or CSV rows (or provider-neutral image extraction JSON), detect their layout, validate them against a declared schema, and write only accepted records into an Excel workbook. Uncertain records go to a separate review sheet instead of being silently guessed. XLSX sheet selection is content-based, and an unrecognized layout stops the run instead of producing partial output.

## What the MVP actually does

Feed it provider-neutral extracted JSON from screenshots, scans, receipts, or forms, and it will keep the source image name and extraction confidence on every row, validate required fields, and split the result: accepted rows go to `Accepted`, anything low-confidence or incomplete goes to `Review` instead of getting silently guessed at, and an `Audit` sheet logs counts and processing metadata. Fuzzy matching never gets to make the final call on its own.

None of the code or examples here came from a real production workflow — no customer files, addresses, payroll records, or credentials. The design was shaped by a real recurring reporting pipeline, but everything in this repo is synthetic.

## Synthetic field example

Imagine a facilities team receiving weekly inspection records in several inconsistent spreadsheets and occasional phone screenshots. Each customer expects the results in a different workbook template.

Excel-Ops should:

1. identify the input layout;
2. extract and normalize each observation;
3. match it to a declared destination using strict rules;
4. place uncertain matches in a human review queue;
5. write accepted records into the appropriate workbook layout;
6. verify counts, dates and required fields before delivery.

Names, locations and identifiers in this repository are fictional. The example demonstrates the workflow without reproducing any organization's data or proprietary configuration.

## Quick start

```bash
python -m pip install -e .
excel-ops examples/extracted-records.json output.xlsx
# The same command accepts .xlsx, .xlsm, and .csv inputs.
```

Pass `--locale en-US` (or another explicit locale hint) when date order or
decimal separators are known. The delivery includes a `Type Inference` sheet,
and the JSON result exposes field-level confidence and ambiguity counts. Numeric
identifiers and leading zeroes remain text; ambiguous dates are sent to review
rather than silently converted.

For the integrated ingest-to-verified-delivery run against declared workbook templates:

```bash
excel-ops deliver delivery-plan.json --dry-run   # the checkable plan, no file touched
excel-ops deliver delivery-plan.json             # write a staging copy, verify it, then deliver
```

See [End-to-end delivery pipeline](docs/delivery-pipeline.md) for the target declaration, the plan contract, the failure codes, and what is deliberately left out.

Every delivered output also gets a source-and-verification Manifest — which inputs, template, and Recipe version produced it, a per-tab breakdown, and the persisted file's own hash and row counts, re-verified against the real file rather than only the run's counters. See [Delivery manifests](docs/delivery-manifest.md).

See [Agent workflow](docs/agent-workflow.md) for the conversation contract and [Connector contract](docs/connectors.md) for local and cloud spreadsheet behavior.

See [Batch ambiguity confirmation and Recipes](docs/ambiguity-recipes.md) for grouped decisions, run/project scopes, and conflict handling.

To find out which files in a working directory are this period's inputs before any of them is read:

```bash
excel-ops scan-workdir ./september --period-start 2026-09-01 --period-end 2026-09-30
```

The scan is read-only, only touches directories you list, never reads a file that is still being written or synced, and never picks a winner among `final.xlsx` / `final (1).xlsx` / `final-final.xlsx`. See [Working-directory scan](docs/workdir-scan.md); it is not yet wired into `excel-ops deliver`.

To make a periodic task idempotent, so that re-running it with nothing changed returns a no-op instead of redoing the work:

```bash
excel-ops deliver delivery-plan.json --run-state --task-key weekly-north
```

Inputs and templates are fingerprinted by content, so a renamed but identical file is not reprocessed, while a changed template, Recipe, or human review decision always triggers a new run — and a failed or interrupted run is never recorded as a completed one. See [Idempotent execution](docs/idempotency.md).

For comparing workbook structures before append, join, or write-back, see [Schema drift](docs/schema-drift.md).

For system update checks, confirmed updates, and rollback, see [Refresh](docs/refresh.md).

## Input format

```json
{
  "source": "synthetic-patrol-sheet.png",
  "records": [
    {
      "location": "100 Example Avenue",
      "event_date": "2026-09-08",
      "identifier": "DEMO 123",
      "category": "synthetic example",
      "confidence": 0.97
    }
  ]
}
```

Image-to-JSON adapters are deliberately separated from workbook writing. A future adapter may use a local OCR engine or a vision model, but it must emit this same reviewable contract.

## Product principles

1. **No silent guesses.** Uncertainty is an output, not an implementation detail.
2. **Preserve provenance.** Every written row points back to its source.
3. **Separate extraction from acceptance.** A model may propose data; validation decides where it goes.
4. **Verify the delivery.** Output counts and required fields are checked after writing.
5. **Make recurring work reproducible.** Rules belong in configuration, not in somebody's memory.

## Roadmap

The detailed, acceptance-test-driven roadmap lives in [Roadmap issue #5](https://github.com/Schlaflied/excel-ops/issues/5):

1. **Phase 1 — Local XLSX:** reliable local ingest, normalization, matching, review, template write-back, verification, and recipes.
2. **Phase 2 — Cloud connectors:** the same workflow over synced folders, Google Sheets, Dropbox, Feishu, WPS, and later Microsoft Graph where demand justifies it.
3. **Phase 3 — Prompt-to-analysis:** safe workbook joins, multi-tab delivery grouping, verified summaries/pivots, and repeatable automation.

**Done so far, all in Phase 1:** ingestion and layout detection ([#1](https://github.com/Schlaflied/excel-ops/issues/1)), type/locale inference ([#24](https://github.com/Schlaflied/excel-ops/issues/24)), schema-drift detection ([#10](https://github.com/Schlaflied/excel-ops/issues/10)), strict matching and dedup ([#2](https://github.com/Schlaflied/excel-ops/issues/2)), offline review packs ([#18](https://github.com/Schlaflied/excel-ops/issues/18)), safe template write-back ([#3](https://github.com/Schlaflied/excel-ops/issues/3)), independent delivery verification ([#4](https://github.com/Schlaflied/excel-ops/issues/4)), business-period resolution and date refresh ([#16](https://github.com/Schlaflied/excel-ops/issues/16), [#9](https://github.com/Schlaflied/excel-ops/issues/9)), safe output naming ([#17](https://github.com/Schlaflied/excel-ops/issues/17)), static formula integrity checks ([#11](https://github.com/Schlaflied/excel-ops/issues/11)), batch ambiguity confirmation and Recipes ([#26](https://github.com/Schlaflied/excel-ops/issues/26)), working-directory scan ([#15](https://github.com/Schlaflied/excel-ops/issues/15)), idempotent execution ([#19](https://github.com/Schlaflied/excel-ops/issues/19)), delivery manifests ([#20](https://github.com/Schlaflied/excel-ops/issues/20)), semantic number/currency formats ([#23](https://github.com/Schlaflied/excel-ops/issues/23)), the integrated `run_delivery(...)` pipeline ([#46](https://github.com/Schlaflied/excel-ops/issues/46)), cross-platform CI ([#50](https://github.com/Schlaflied/excel-ops/issues/50)), and — most recently — a JS/TS Agent orchestration layer over MCP that forwards to the same Python CLI without reimplementing it ([#53](https://github.com/Schlaflied/excel-ops/issues/53)), plus an agent-callable delivery-plan preparation tool on top of it ([#59](https://github.com/Schlaflied/excel-ops/issues/59)).

**In progress:** natural-language formula generation ([#25](https://github.com/Schlaflied/excel-ops/issues/25)).

**Still open / not started:** natural-language summaries and pivot tables ([#6](https://github.com/Schlaflied/excel-ops/issues/6)), conditional formatting ([#7](https://github.com/Schlaflied/excel-ops/issues/7)), batch name-tag/label generation ([#8](https://github.com/Schlaflied/excel-ops/issues/8)), version-diff audit packages ([#12](https://github.com/Schlaflied/excel-ops/issues/12)), workbook health checks and auditable auto-repair ([#13](https://github.com/Schlaflied/excel-ops/issues/13)), role packs ([#14](https://github.com/Schlaflied/excel-ops/issues/14)), unattended file-arrival triggers ([#21](https://github.com/Schlaflied/excel-ops/issues/21)), all of Phase 2's cloud connectors, and the rest of Phase 3.

## Current capabilities

v0.3.0 released multi-source ingestion, type and locale inference, schema-drift detection, strict matching, and an offline human-review round trip.

`main` additionally contains, **merged but not part of any newer Release**: business-period resolution, period-aware date refresh, safe output naming, batch ambiguity confirmation and project Recipes, safe template write-back, independent delivery verification, static formula integrity checks, the integrated `run_delivery(...)` pipeline that chains them into one verified Phase 1 delivery run ([#46](https://github.com/Schlaflied/excel-ops/issues/46), [docs](docs/delivery-pipeline.md)), a working-directory scan, idempotent execution, delivery manifests, semantic number formats, cross-platform CI, and a local Agent MCP server exposing `scan_workdir`, `prepare_delivery`, `plan_delivery`, and `run_delivery` as structured tools over the same Python core ([#53](https://github.com/Schlaflied/excel-ops/issues/53), [#59](https://github.com/Schlaflied/excel-ops/issues/59), [docs](docs/mcp.md)).

Merged modules, repository tests, and a verified persisted file are separate kinds of evidence from a tagged Release. Nothing after v0.3.0 has been released.

[See the complete capability ledger, implementation status, safety boundaries, and linked issues/PRs](docs/capabilities.md).

`main` now includes multi-format XLSX/CSV/PDF delivery ([#22](https://github.com/Schlaflied/excel-ops/issues/22), [docs](docs/multi-format-export.md)). Natural-language formula generation remains in progress: typed, version-aware planning, safe workbook application, independent recalculation, and the Agent delivery contract are documented in [formula planning](docs/formula-planning.md). Every cloud connector is still out of scope. Semantic currency and precision rules are documented in [the number-format policy guide](docs/number-format-policy.md).

## License

MIT
