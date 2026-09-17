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

## Use with an agent

In Codex or another agent that follows the Agent Skills standard, invoke the repository skill:

> Use `$excel-agent` to extract the rows from these screenshots, append only high-confidence records to the workbook, and put everything uncertain in a review sheet.

`$excel-agent` is the only user-facing entry point. The Python CLI remains an internal deterministic capability used by the agent and by automated tests.

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

See [Agent workflow](docs/agent-workflow.md) for the conversation contract and [Connector contract](docs/connectors.md) for local and cloud spreadsheet behavior.

For comparing workbook structures before append, join, or write-back, see [Schema drift](docs/schema-drift.md).

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

### Phase 1 foundations in v0.3.0

| Capability | Issue | Pull request | Status |
|---|---:|---:|---|
| Multi-source, layout-aware ingestion | [#1](https://github.com/Schlaflied/excel-ops/issues/1) | [#29](https://github.com/Schlaflied/excel-ops/pull/29) | Merged |
| Type, unit, and locale inference | [#24](https://github.com/Schlaflied/excel-ops/issues/24) | [#28](https://github.com/Schlaflied/excel-ops/pull/28) | Merged |
| Schema-drift detection and mappings | [#10](https://github.com/Schlaflied/excel-ops/issues/10) | [#30](https://github.com/Schlaflied/excel-ops/pull/30) | Merged |
| Strict matching and deduplication | [#2](https://github.com/Schlaflied/excel-ops/issues/2) | [#27](https://github.com/Schlaflied/excel-ops/pull/27) | Merged |
| Offline human-review round trip | [#18](https://github.com/Schlaflied/excel-ops/issues/18) | [#31](https://github.com/Schlaflied/excel-ops/pull/31) | Merged |

These foundations are merged and covered by the repository test suite. They do not yet complete the full Phase 1 delivery loop: the next vertical-slice boundary is safe template write-back ([#3](https://github.com/Schlaflied/excel-ops/issues/3)) followed by independent delivery verification ([#4](https://github.com/Schlaflied/excel-ops/issues/4)).

## License

MIT
