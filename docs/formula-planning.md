[中文](formula-planning.zh-CN.md) | English

# Formula planning contract

Issue [#25](https://github.com/Schlaflied/excel-ops/issues/25) starts with a typed planning boundary. The Agent explains and interprets the user's natural-language business goal; the Python core accepts only explicit cell, range, sheet, Excel-version, and output-mode decisions. This separation keeps free-form text out of formula syntax.

The first slice supports lookup, conditional sum, and date-offset plans. Excel 2021 and Microsoft 365 use `XLOOKUP`; Excel 2016 and 2019 receive an orientation-aware `INDEX/MATCH` fallback and are limited to a single returned row or column because this contract does not carry legacy array-entry metadata. `SUMIFS` and `EDATE` work across all declared target versions.

Formula mode returns formula text and marks the plan as requiring independent recalculation. Static mode never evaluates a formula through openpyxl: the caller must provide a value from an independent calculation path.

## Safe workbook application

`plan_formula_application(...)` produces a value-free dry-run ledger for an explicit sheet and bounded target range. `apply_formula_plan(...)` requires `confirmed=True`, always writes a new workbook, preserves the source, refuses an existing output, skips protected formulas and merged-range followers by default, translates relative references while filling a range, then reopens the staged copy to verify every planned write before publishing it. Formula replacement requires the explicit `overwrite_formulas=True` option. Static values are limited to one target cell until a later contract can carry one independently computed value per row.

## Integrity and independent recalculation

`verify_formula_recalculation(...)` first runs the existing #11 static checks for broken references, missing sheets, invalid ranges, error tokens, declared fill gaps, and circular references. It then recalculates a staged copy with Microsoft Excel on Windows when available, otherwise with LibreOffice, and compares the engine-produced cached values against explicit independently derived expectations with optional numeric tolerances. The verified copy is published only when all checks pass.

No installed engine produces `recalculation_not_verified`; omitted expectations produce `recalculation_expectations_missing`. Both remain `unverified`, never `verified`. A formula error or expectation mismatch is `failed` and the staged output is discarded.

## Delivery and Agent integration

A delivery target may now declare `formulas`. Each rule contains the serialized typed plan, an explicit sheet and target range, independently derived expectations for formula mode, overwrite policy, and calculation engine. `run_delivery(...)` applies those rules only to the staged copy, fails closed on skipped cells or unverified recalculation, and then runs the ordinary persisted-workbook verification.

The rule identity is part of #19's mapping fingerprint, so changing the business rule, Excel version, output mode, target range, overwrite policy, or expectation scope invalidates a previous no-op baseline. The delivery Manifest records the business rule, compatibility strategy, written range, engine, and verification status without copying expected cell values. The same fields pass through `excel_ops.prepare_delivery`, so the existing MCP sequence can plan and execute formula delivery without invoking separate scripts.
