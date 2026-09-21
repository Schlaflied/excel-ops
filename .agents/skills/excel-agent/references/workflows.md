# Workflow selection

Choose the smallest workflow that produces the requested business outcome. A request may pass through several workflows without asking the user to name them.

## Session start check

Once per session, before the first workbook workflow, run `node refresh.mjs agent-check`. It is read-only, caches a successful result for 24 hours, and returns one JSON object. When `updateAvailable` is true, mention it and offer `node refresh.mjs preview`; the user runs `node refresh.mjs apply --confirm` themselves. On `offline` or any failure, continue the workbook work unchanged. Never apply or roll back an update yourself, and keep this output out of the report you return for the spreadsheet task. See `docs/refresh.md`.

## Image or document to spreadsheet

Extract candidate records, preserve image/page/region provenance, normalize fields, route unreadable or low-confidence values to review, then write and verify accepted records. OCR text is evidence, not automatically accepted truth.

## Clean or normalize

Profile the source, propose normalization rules, preserve raw values, apply deterministic transformations, isolate exceptions, and reconcile row counts and key totals.

## Join or lookup

Identify candidate keys, report uniqueness and missingness, prefer exact stable identifiers, and make fuzzy results review-only. Produce matched, ambiguous, unmatched-left, unmatched-right, and duplicate-key counts.

## Update an existing workbook

Snapshot the target, map only authorized ranges, preserve formulas and formatting outside scope, write through a copy or conflict-aware revision, reread, and verify.

## Template-driven recurring delivery

Detect the input layout, normalize records, preassign each accepted record to exactly one destination, write using a declared template mapping, update zero-record periods, verify actual output files, and move only passing artifacts into delivery.

For declared local workbook templates this whole loop is one deterministic call. Run the plan first, read its counts and blocking items back to the user, then run the delivery and report the verified result:

```bash
excel-ops deliver delivery-plan.json --dry-run
excel-ops deliver delivery-plan.json --recipe recipes/project-recipe.json --result run.json
```

Equivalent in Python: `excel_ops.run_delivery(inputs, targets, staging_dir=..., delivery_dir=...)`. Treat `delivered: false`, any `failures[]` entry, and any record in `review` or `rejected` as items to explain, never to work around. See `docs/delivery-pipeline.md` for the target declaration and failure codes.

## Analyze or summarize

Keep source data separate from calculations and outputs. Explain the metric definition and source range. Do not build a dashboard when a compact table or spreadsheet result answers the question.

## Delivery contents

Return the primary spreadsheet plus, when applicable, a review queue, change log, verification report, and credential-free rerun recipe. Do not create empty ceremonial artifacts.
