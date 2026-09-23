[中文](formula-planning.zh-CN.md) | English

# Formula planning contract

Issue [#25](https://github.com/Schlaflied/excel-ops/issues/25) starts with a typed planning boundary. The Agent explains and interprets the user's natural-language business goal; the Python core accepts only explicit cell, range, sheet, Excel-version, and output-mode decisions. This separation keeps free-form text out of formula syntax.

The first slice supports lookup, conditional sum, and date-offset plans. Excel 2021 and Microsoft 365 use `XLOOKUP`; Excel 2016 and 2019 receive an orientation-aware `INDEX/MATCH` fallback and are limited to a single returned row or column because this contract does not carry legacy array-entry metadata. `SUMIFS` and `EDATE` work across all declared target versions.

Formula mode returns formula text and marks the plan as requiring independent recalculation. Static mode never evaluates a formula through openpyxl: the caller must provide a value from an independent calculation path.

## Safe workbook application

`plan_formula_application(...)` produces a value-free dry-run ledger for an explicit sheet and bounded target range. `apply_formula_plan(...)` requires `confirmed=True`, always writes a new workbook, preserves the source, refuses an existing output, skips protected formulas and merged-range followers by default, translates relative references while filling a range, then reopens the staged copy to verify every planned write before publishing it. Formula replacement requires the explicit `overwrite_formulas=True` option. Static values are limited to one target cell until a later contract can carry one independently computed value per row.

Independent calculation-engine verification and integration with the #11 verifier and #19 idempotency record remain follow-up slices of #25.
