# Workflow selection

Choose the smallest workflow that produces the requested business outcome. A request may pass through several workflows without asking the user to name them.

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

## Analyze or summarize

Keep source data separate from calculations and outputs. Explain the metric definition and source range. Do not build a dashboard when a compact table or spreadsheet result answers the question.

## Delivery contents

Return the primary spreadsheet plus, when applicable, a review queue, change log, verification report, and credential-free rerun recipe. Do not create empty ceremonial artifacts.
