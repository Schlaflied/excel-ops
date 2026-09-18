[中文](capabilities.zh-CN.md) | English

# Current capabilities and implementation status

This page records **what Excel Ops can actually do today, where the implementation lives, its safety boundaries, and whether the capability has been released**. Future plans remain in [Roadmap #5](https://github.com/Schlaflied/excel-ops/issues/5).

## Status definitions

| Status | Meaning |
|---|---|
| Released | Merged into `main` and included in the named GitHub Release. |
| Merged to main | Merged into `main` and covered by repository tests, but not yet included in a newer Release. |
| Planned | Exists only as an Issue, interface proposal, or Roadmap item and must not be described as available. |

A merged PR, automated tests, verification of an actual persisted file, and business approval are separate kinds of evidence. A merged module does not establish that the complete business workflow works end to end.

## Released: v0.3.0

[v0.3.0](https://github.com/Schlaflied/excel-ops/releases/tag/v0.3.0) delivered the first group of local-spreadsheet foundations for Phase 1.

### 1. Multi-source, layout-aware ingestion

- **What it does:** reads `.xlsx`, `.xlsm`, `.csv`, and provider-neutral image-extraction JSON; detects the data sheet, header, and data region from content.
- **Output and evidence:** normalizes rows into `ExtractedRecord` values while retaining file, sheet, row, or image-region provenance.
- **Safety boundary:** an unknown layout stops the run instead of producing silent partial output; `.xlsm` files can be read, but macros are never executed.
- **Implementation:** [Issue #1](https://github.com/Schlaflied/excel-ops/issues/1) / [PR #29](https://github.com/Schlaflied/excel-ops/pull/29)

### 2. Data type, unit, and locale inference

- **What it does:** identifies dates, Excel serial dates, numbers, percentages, currencies, and identifiers; accepts an explicit locale hint.
- **Output and evidence:** reports field-level type, confidence, unit, locale, ambiguity counts, and low-confidence counts.
- **Safety boundary:** employee IDs, postal codes, phone numbers, and leading zeroes are preserved as text; ambiguous dates go to review instead of being silently converted.
- **Implementation:** [Issue #24](https://github.com/Schlaflied/excel-ops/issues/24) / [PR #28](https://github.com/Schlaflied/excel-ops/pull/28)

### 3. Schema-drift detection and confirmed mappings

- **What it does:** compares a baseline workbook with candidate workbooks and detects field additions, removals, reorderings, type changes, requiredness changes, enum changes, and changes in record granularity.
- **Output and evidence:** creates a machine-readable drift report and describes downstream impact on append, join, and delivery grouping.
- **Safety boundary:** heuristic rename matches remain review suggestions; only human-confirmed mappings can be saved and reused.
- **Implementation:** [Issue #10](https://github.com/Schlaflied/excel-ops/issues/10) / [PR #30](https://github.com/Schlaflied/excel-ops/pull/30) / [Detailed guide](schema-drift.md)

### 4. Strict matching, preallocation, and deduplication

- **What it does:** applies configured exact, case-insensitive, and normalized matching in order and preallocates each record to at most one destination.
- **Output and evidence:** returns a stable record ID, the applied rule, confidence information, candidate destinations, and exception reasons.
- **Safety boundary:** fuzzy matching is always review-only; conflicts, duplicates, and unmatched records are never written automatically.
- **Implementation:** [Issue #2](https://github.com/Schlaflied/excel-ops/issues/2) / [PR #27](https://github.com/Schlaflied/excel-ops/pull/27)

### 5. Offline human-review pack

- **What it does:** creates a review workbook that opens in ordinary Excel and supports `accept`, `correct`, `reject`, and `cannot determine` decisions.
- **Output and evidence:** imports decisions by stable record ID and retains an idempotent review history.
- **Safety boundary:** `correct` requires a replacement value; `cannot determine` never enters accepted results; a hidden manifest detects tampered headers, IDs, candidate values, and duplicate IDs.
- **Implementation:** [Issue #18](https://github.com/Schlaflied/excel-ops/issues/18) / [PR #31](https://github.com/Schlaflied/excel-ops/pull/31)

## Merged to main, not yet released

The following capabilities entered `main` after v0.3.0 and therefore must not be described as part of that release.

### 6. Deterministic business-period resolution

- **What it does:** resolves weeks, months, quarters, custom date ranges, fiscal years, and fiscal quarters with explicit `as_of`, timezone, week-start, and fiscal-year-start inputs.
- **Output and evidence:** `resolve_period(...)` returns one shared `PeriodResult` consumed by refresh and naming modules.
- **Safety boundary:** ambiguous expressions such as “last week” that depend on an implicit current date raise an error instead of being guessed.
- **Implementation:** [Issue #16](https://github.com/Schlaflied/excel-ops/issues/16) / [PR #33](https://github.com/Schlaflied/excel-ops/pull/33)

### 7. Period-aware date refresh

- **What it does:** updates declared date slots for an authorized business period, with dry run, zero-record-period updates, and post-write verification for every slot.
- **Output and evidence:** keeps the refresh plan separate from persisted results and treats `PeriodResult` as the authoritative date source.
- **Safety boundary:** historical dates, formula dates, and undeclared regions are not rewritten; repeated execution is idempotent.
- **Implementation:** [Issue #9](https://github.com/Schlaflied/excel-ops/issues/9) / [PR #34](https://github.com/Schlaflied/excel-ops/pull/34)

### 8. Safe, period-aware output naming

- **What it does:** creates safe names such as `payroll-2026-09-12.xlsx` from a business period and handles extensions, invalid characters, destination directories, and revision suffixes.
- **Output and evidence:** an identical existing target can produce a no-op; different content receives a `-rN` suffix, and the persisted path is verified.
- **Safety boundary:** the module does not read the system's current date; naming dates must come from the resolved upstream business period.
- **Implementation:** [Issue #17](https://github.com/Schlaflied/excel-ops/issues/17) / [PR #35](https://github.com/Schlaflied/excel-ops/pull/35)

## What can currently be composed

The current modules cover:

```text
ingest → normalize/type inference → schema check → strict match → human review
                                  period resolution → date refresh → safe naming
```

These capabilities have explicit data contracts and tests, but the full `write → verify → deliver` vertical slice is not yet complete. Excel Ops therefore cannot yet claim unattended end-to-end delivery into arbitrary enterprise templates.

## Not implemented or not yet complete end to end

- [Issue #3](https://github.com/Schlaflied/excel-ops/issues/3): safe write-back into an existing enterprise workbook template;
- [Issue #4](https://github.com/Schlaflied/excel-ops/issues/4): reopen and independently verify the actual delivery artifact;
- formula integrity, item-level reconciliation, source manifests, and complete recipes;
- local sync folders and Google Sheets, Dropbox, Feishu, and WPS connectors;
- prompt-driven append, join, multi-tab delivery grouping, summaries, and pivot tables;
- a full synthetic end-to-end fixture for `ingest → normalize → match → review → write → verify → deliver`.

## Maintenance rules

- Update the relevant entry when a feature PR merges, but initially mark it only as “Merged to main.”
- Move a capability to “Released” and name its version only after publishing a GitHub Release.
- Every capability must state its purpose, boundary, and implementation links rather than listing only a module name.
- Keep future plans in Roadmap #5; do not duplicate its changing backlog on this page.
