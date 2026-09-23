[中文](formula-planning.zh-CN.md) | English

# Formula planning contract

Issue [#25](https://github.com/Schlaflied/excel-ops/issues/25) starts with a typed planning boundary. The Agent explains and interprets the user's natural-language business goal; the Python core accepts only explicit cell, range, sheet, Excel-version, and output-mode decisions. This separation keeps free-form text out of formula syntax.

The first slice supports lookup, conditional sum, and date-offset plans. Excel 2021 and Microsoft 365 use `XLOOKUP`; Excel 2016 and 2019 receive an `INDEX/MATCH` fallback. `SUMIFS` and `EDATE` work across all declared target versions.

Formula mode returns formula text and marks the plan as requiring independent recalculation. Static mode never evaluates a formula through openpyxl: the caller must provide a value from an independent calculation path. Workbook application, existing-formula protection, static verification, and integration with the #11 verifier and #19 idempotency record remain follow-up slices of #25.
