[中文](ambiguity-recipes.zh-CN.md) | English

# Batch ambiguity confirmation and Recipes

Excel-Ops groups questions by field and rule so reviewers decide once instead of answering row by row. Each question carries representative samples, candidate interpretations, affected-row count, a recommendation, its rationale, and confidence. A recommendation remains advice rather than fact.

Decisions may use two scopes:

- `this-run`: applies only to the current run and is not persisted;
- `project`: is stored in a credential-free `excel-ops-recipe-v1` JSON Recipe.

An unavailable saved choice becomes a `conflict` and must be confirmed again. `pending`, `unknown`, and `conflict` all block formal delivery. `review_rows_from_ambiguities()` connects field-level questions to the #18 Review Pack; `manifest_decisions()` exposes provenance for the future #20 Manifest without implementing template write-back or delivery verification.
