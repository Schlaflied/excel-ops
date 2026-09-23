[中文](multi-format-export.zh-CN.md) | English

# Multi-format delivery export contract

This document defines the implemented multi-format delivery boundary for [Issue #22](https://github.com/Schlaflied/excel-ops/issues/22). CSV and PDF adapters consume the verified XLSX delivery; they do not repeat matching or workbook business logic.

## User-facing request

The delivery request may select one or more formats:

```yaml
delivery:
  formats: [xlsx, pdf]
  csv:
    mode: one-file-per-sheet
  pdf:
    sheets: all
```

Supported format identifiers for the first vertical slice are `xlsx`, `csv`, and `pdf`. The default is intentionally unset. If the user does not choose a format, the Agent must recommend or ask before export; the exporter must not silently choose one.

## Capability contract

Every requested format produces a separate output artifact and an export result with:

- the requested and actual format;
- the source workbook and included sheets;
- the output path and content digest;
- warnings describing known information loss;
- verification status and structured findings.

The extension is derived from the actual writer, never from a user-provided filename. A writer must reject a path whose suffix does not match its format.

### XLSX

Retains the existing workbook delivery semantics: multiple sheets, formulas, styles, images, conditional formatting, and editable structure. The existing write, reread, verification, naming, idempotency, and manifest contracts remain authoritative.

### CSV

CSV is a flat, single-sheet interchange format. It cannot preserve formulas, styles, merged cells, images, charts, or additional sheets. For a workbook with multiple sheets, the request must explicitly choose one sheet or `one-file-per-sheet`; an exporter must return `needs_review` instead of selecting the active sheet implicitly.

### PDF

PDF is a fixed-layout review and archival artifact, not an editable workbook. The first implementation must make the selected sheet scope explicit (`all` or a named list) and verify page rendering for pagination, scaling, repeated headers, clipping, and unreadable output before reporting success.

## Pipeline boundary

Format selection belongs in the delivery plan and is resolved before writing. Format adapters must consume the verified delivery result; they must not reimplement source matching, ambiguity decisions, template mapping, or business calculations.

```text
plan → dry run → user confirmation → write XLSX → export CSV/PDF → reread/verify → manifest
```

An individual export failure must not be represented as a successful complete delivery. The result should identify successful, blocked, and review-required artifacts separately.

## Manifest additions

The delivery Manifest should record an `exports` entry per artifact containing the format, source workbook, included sheets, output digest, verification status, and loss warnings. It must not copy cell values, credentials, or tokens into the manifest.

## Planned implementation slices

1. Validate and normalize the format-selection contract in the delivery plan.
2. Add an XLSX pass-through artifact and shared export result/Manifest shape.
3. Add explicit single-sheet and one-file-per-sheet CSV export with tests.
4. Add PDF export behind a capability check and page-level verification.
5. CLI/MCP exposure, examples, and end-to-end tests are integrated through the existing confirmed `run_delivery` tool.

`prepare_delivery` accepts the `delivery` object above. `plan_delivery` returns an `export_plan`; after explicit approval, `run_delivery` writes and verifies XLSX first, creates the selected siblings, and adds an `exports` array to the delivery Manifest. Each entry records actual format, source workbook, included sheets, content digest, verification status, loss warnings, and a renderer summary. See [`examples/multi-format-delivery.json`](../examples/multi-format-delivery.json).
