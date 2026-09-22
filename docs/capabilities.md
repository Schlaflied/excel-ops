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

### 9. Batch ambiguity confirmation and project Recipes

- **What it does:** groups repeated field-level questions into one confirmation batch and reuses saved decisions in `this-run` or `project` scope.
- **Output and evidence:** each item carries its status, selected value, decision provenance, and affected count; project decisions persist as a versioned Recipe file.
- **Safety boundary:** `unknown` and conflicting decisions block delivery instead of being applied; changed candidates invalidate a saved decision rather than silently reusing it.
- **Implementation:** [Issue #26](https://github.com/Schlaflied/excel-ops/issues/26) / [PR #41](https://github.com/Schlaflied/excel-ops/pull/41) / [Detailed guide](ambiguity-recipes.md)

### 10. Safe write-back into an existing template

- **What it does:** copies an enterprise template and changes only the cells a declared `TemplateMapping` authorizes.
- **Output and evidence:** returns the actual written path plus a JSON change log of every previous and new value, and every skipped cell with its reason.
- **Safety boundary:** the source template is never modified, an existing output is never overwritten, and formulas and non-anchor merged cells are skipped rather than replaced; formatting comes from a declared style row.
- **Implementation:** [Issue #3](https://github.com/Schlaflied/excel-ops/issues/3) / [PR #40](https://github.com/Schlaflied/excel-ops/pull/40)

### 11. Independent verification of the delivered file

- **What it does:** reopens the writer's real output, and then the published copy, and checks sheets, headers, required fields, record IDs, written values, and declared period slots.
- **Output and evidence:** writes a machine-readable verification report with a finding code, message, suggestion, and location for every problem.
- **Safety boundary:** fail-closed — unknown finding severities block delivery, staged and delivery paths must differ, and a failed run publishes nothing.
- **Implementation:** [Issue #4](https://github.com/Schlaflied/excel-ops/issues/4) / [PR #42](https://github.com/Schlaflied/excel-ops/pull/42)

### 12. Static formula integrity checks

- **What it does:** scans formula text for errors, external references, broken references, missing sheets, invalid ranges, unsupported dynamic arrays, and circular references; verifies declared formula regions and reconciles declared summaries.
- **Output and evidence:** returns findings usable as a delivery verifier.
- **Safety boundary:** openpyxl cannot calculate formulas, so recalculation is reported as an explicit `recalculation_not_verified` warning and cached values are never treated as proof.
- **Implementation:** [Issue #11](https://github.com/Schlaflied/excel-ops/issues/11) / [PR #43](https://github.com/Schlaflied/excel-ops/pull/43)

### 13. Integrated end-to-end delivery run

- **What it does:** `run_delivery(...)` chains ingestion, the data contract, strict matching, ambiguity confirmation and Recipes, a checkable plan or dry run, staged template write-back, independent verification, and optional formula checks into one agent-callable run, also exposed as `excel-ops deliver`.
- **Output and evidence:** one JSON-serializable result with per-record terminal state, stable record IDs, source provenance for every delivered cell, per-target actual delivery paths, verification findings, and failure codes; covered by a synthetic end-to-end fixture with two input layouts, an image-extraction JSON, and two templates.
- **Safety boundary:** a save is never reported as a delivery — only a target whose reopened file passed verification is published; unresolved ambiguities, conflicts, and duplicates never reach accepted; a skipped mapped cell fails the run closed instead of delivering a partial row; inputs and templates stay unmodified.
- **Implementation:** [Issue #46](https://github.com/Schlaflied/excel-ops/issues/46) / [Detailed guide](delivery-pipeline.md)

### 14. Working-directory scan and current-period file selection

- **What it does:** scans an explicit allowlist of directories and classifies every file as `input`, `template`, `prior_delivery`, `review_return`, or `unknown` from its extension, filename, modification time, content hash, and workbook metadata; groups byte-identical duplicates and look-alike version candidates; reports include/review/exclude with a reason per file; also exposed as `excel-ops scan-workdir`.
- **Output and evidence:** one JSON-serializable report with the classification, disposition, reason, and deciding signal for every file, the duplicate and version-candidate groups, skipped paths, and the list of paths actually accessed.
- **Safety boundary:** read-only — no file is moved, renamed, or deleted; nothing outside the authorized roots is accessed and symlinks are not followed; a file whose size and modification time have not been stable is never even opened; an undeterminable latest version is never auto-selected; cloud-sync lock files, incomplete downloads, and conflict copies never reach `include`, and resolving a conflict remains out of scope; a user override can relabel a file and be saved as a reusable Recipe but cannot promote an unstable file or a sync artifact.
- **Not yet:** not wired into `run_delivery(...)` or `excel-ops deliver`; that integration is a separate step.
- **Implementation:** [Issue #15](https://github.com/Schlaflied/excel-ops/issues/15) / [Detailed guide](workdir-scan.md)

### 15. Idempotent execution of a periodic task

- **What it does:** fingerprints a run from the content hash of its inputs, the template content plus Template Profile version, the mapping and project Recipe, the human confirmations actually applied, the declared report period, the output target, the existing output's content hash, and the target's remote revision; compares that fingerprint with a persisted run record and returns `no_op`, `changed`, or `retry`. Wired into `run_delivery(...)` as an early short-circuit (`idempotency=IdempotencyOptions(...)`, `excel-ops deliver --run-state`) that returns before any matching, write, or verification happens.
- **Output and evidence:** one versioned run record per task in `<delivery_dir>/.excel-ops/idempotency.json` holding the fingerprint, its per-component digests, the run status, the attempt count, and aggregate detail; the run result carries `no_op` plus a `run_decision` naming the reason and the components that changed.
- **Safety boundary:** only a run explicitly recorded as `succeeded` can justify a no-op — a failed, blocked, or interrupted run returns `retry` and re-runs; a changed template, Recipe, or human review decision always re-runs; files are identified by content, never by name, path, size, or modification time; every persisted component is a digest, so no path, destination name, cell value, cloud identifier, or credential is stored; a corrupt or foreign run-state file never authorizes a no-op; a dry run never short-circuits.
- **Not yet:** the cloud-target revision conflict check is a forward-compatible extension point with stub-backed tests, not a working cloud connector; the decision is whole-run, not per-record — the independent per-record deduplication in the delivery pipeline still decides what is appended.
- **Implementation:** [Issue #19](https://github.com/Schlaflied/excel-ops/issues/19) / [Detailed guide](idempotency.md)

## What can currently be composed

The current modules cover the full Phase 1 local loop:

```text
ingest → normalize/type inference → schema check → strict match → ambiguity/Recipe
       → plan/dry run → staged template write → independent verification → delivery
                                  period resolution → date refresh → safe naming
```

This loop is exercised end to end against synthetic fixtures and judged by the reopened delivery file. It is still not part of a tagged Release, and it only covers declared local workbook templates.

## Not implemented or not yet complete end to end

- cloud-target revision conflict handling for [Issue #19](https://github.com/Schlaflied/excel-ops/issues/19). Run-level fingerprinting and the no-op short-circuit are implemented (capability 15), but the revision conflict check is only a forward-compatible interface until a real cloud connector exists;
- [Issue #20](https://github.com/Schlaflied/excel-ops/issues/20) source and verification Manifests, [Issue #23](https://github.com/Schlaflied/excel-ops/issues/23) currency and precision rules, and [Issue #22](https://github.com/Schlaflied/excel-ops/issues/22) multi-format export;
- recalculated-formula proof, which requires Excel or LibreOffice rather than openpyxl;
- local sync folders and Google Sheets, Dropbox, Feishu, and WPS connectors;
- prompt-driven append, join, multi-tab delivery grouping, summaries, and pivot tables;
- cross-source fact checking and regional rule calculations.

## Maintenance rules

- Update the relevant entry when a feature PR merges, but initially mark it only as “Merged to main.”
- Move a capability to “Released” and name its version only after publishing a GitHub Release.
- Every capability must state its purpose, boundary, and implementation links rather than listing only a module name.
- Keep future plans in Roadmap #5; do not duplicate its changing backlog on this page.
