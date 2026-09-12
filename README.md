# Excel-Ops

[中文说明](README.zh-CN.md) | English

### From images and messy tables to a checked, repeatable spreadsheet delivery

Excel-Ops is a conversation-driven workflow for AI agents, not another spreadsheet web app. Give an agent the files and describe the outcome in plain language; the agent plans the operation, uses the appropriate local or cloud connector, isolates uncertainty, verifies the result, and returns a delivery package.

Excel-Ops is not an AI formula generator. It treats spreadsheets as an operational workflow:

**ingest → extract → normalize → match → human review → write → verify → deliver**

It is designed as a companion project to [PPT-Ops](https://github.com/luochen211/ppt-ops): PPT-Ops turns source material into a reviewable presentation delivery, while Excel-Ops turns images and messy tabular inputs into a reviewable data delivery.

## Use with an agent

In Codex or another agent that follows the Agent Skills standard, invoke the repository skill:

> Use `$excel-agent` to extract the rows from these screenshots, append only high-confidence records to the workbook, and put everything uncertain in a review sheet.

`$excel-agent` is the only user-facing entry point. The Python CLI remains an internal deterministic capability used by the agent and by automated tests.

The initial connector order is local XLSX/CSV, local cloud-sync folders, Google Sheets, Dropbox API, Feishu Sheets, and WPS Sheets. Cloud credentials belong to the user's environment and are never stored in a project or delivery package.

The first vertical slice focuses on a common business task: extract rows from images or structured captures, validate them against a declared schema, and write only accepted records into an Excel workbook. Uncertain records go to a separate review sheet instead of being silently guessed.

## Why this exists

People rarely want a `VLOOKUP`. They want to connect records from two imperfect sources without losing data or hiding uncertainty. The hard part is not the formula; it is choosing a trustworthy key, handling duplicates, preserving provenance, and showing what did not match.

## MVP contract

- Accept provider-neutral extracted JSON from screenshots, scans, receipts, or forms.
- Preserve the source image name and extraction confidence for every row.
- Validate required fields before workbook output.
- Route low-confidence or incomplete rows to `Review`.
- Write accepted rows to `Accepted`.
- Add an `Audit` sheet with counts and processing metadata.
- Never use fuzzy matching as an invisible final decision.

The repository intentionally does not include customer files, addresses, payroll records, credentials, or code copied from a private production workflow. The design was informed by a real recurring reporting pipeline, but all examples here are synthetic.

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
```

See [Agent workflow](docs/agent-workflow.md) for the conversation contract and [Connector contract](docs/connectors.md) for local and cloud spreadsheet behavior.

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

### Current Phase 1 implementation PRs

| Capability | Issue | Pull request | Status |
|---|---:|---:|---|
| Multi-source, layout-aware ingestion | [#1](https://github.com/Schlaflied/excel-ops/issues/1) | [#29](https://github.com/Schlaflied/excel-ops/pull/29) | Open |
| Type, unit, and locale inference | [#24](https://github.com/Schlaflied/excel-ops/issues/24) | [#28](https://github.com/Schlaflied/excel-ops/pull/28) | Open |
| Schema-drift detection and mappings | [#10](https://github.com/Schlaflied/excel-ops/issues/10) | [#30](https://github.com/Schlaflied/excel-ops/pull/30) | Open; stacked on #28 |
| Strict matching and deduplication | [#2](https://github.com/Schlaflied/excel-ops/issues/2) | [#27](https://github.com/Schlaflied/excel-ops/pull/27) | Open |
| Offline human-review round trip | [#18](https://github.com/Schlaflied/excel-ops/issues/18) | [#31](https://github.com/Schlaflied/excel-ops/pull/31) | Open |

These are review branches, not released capabilities. The next vertical-slice boundary is safe template write-back ([#3](https://github.com/Schlaflied/excel-ops/issues/3)) followed by independent delivery verification ([#4](https://github.com/Schlaflied/excel-ops/issues/4)).

## License

MIT
