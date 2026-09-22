[中文](delivery-pipeline.zh-CN.md) | English

# End-to-end delivery pipeline

`run_delivery(...)` is the single agent-callable entry point that runs one local
spreadsheet task from input files to a verified delivery file. It adds no new
matching, ambiguity, write-back or verification logic; it sequences modules that
were merged and unit-tested separately and returns one JSON-serializable result.

```text
ingest (xlsx / csv / extraction JSON)
  -> data-contract validation          (accepted candidate / rejected)
  -> strict matching + deduplication   (matching.preallocate_records)
  -> batch ambiguity confirmation      (ambiguity + project Recipe)
  -> checkable plan / dry run          (before any file is touched)
  -> staged template write             (template_writer.write_template)
  -> independent verification          (delivery_verification.verify_and_deliver)
  -> optional formula checks           (formula_verification.FormulaVerifier)
  -> delivery, only if verification passed
```

## Guarantees

- **Original files are never modified.** Inputs and enterprise templates are
  read only. Every write goes into a staging copy.
- **A save is not a delivery.** A target is reported as delivered only after
  `verify_and_deliver` reopens the persisted file and every blocking check
  passes. When verification fails, no delivery file is published and the
  machine-readable report explains why.
- **Every input record ends in exactly one state**: `written`, `accepted`
  (accepted but not written, for example in a dry run), `review`, `rejected`, or
  `skipped_existing`. The counts reconcile with the input count.
- **Unresolved questions never reach accepted.** A pending, unknown or
  conflicting ambiguity, a fuzzy candidate, a conflict and a duplicate record ID
  are all routed to review instead of being written.
- **Partial rows are never delivered.** If the writer skips a mapped cell — a
  protected formula or a non-anchor merged cell — the run fails closed with
  `incomplete_write` instead of publishing a half-written row.
- **Every delivered cell is traceable** to its source file, sheet and row or
  image region, and to a stable record ID.

## Declaring a target

```python
from excel_ops import (
    Destination, DeliveryTarget, PeriodExpectation, TemplateMapping, run_delivery,
)

target = DeliveryTarget(
    destination=Destination("North Warehouse", ("North Warehouse", "北区仓库")),
    template_path="templates/north-template.xlsx",
    mapping=TemplateMapping(
        sheet="Inspections",
        header_row=4,
        data_start_row=5,
        field_columns={
            "record_id": "A",
            "location": "B",
            "event_date": "C",
            "identifier": "D",
            "category": "E",
            "source": "F",
        },
    ),
    required_fields=("record_id", "location", "event_date", "identifier", "category", "source"),
    period_expectations=(PeriodExpectation("Inspections", "B2", "2026-09-07 to 2026-09-13"),),
)

result = run_delivery(
    ["sources/week-37.xlsx", "sources/extraction.json"],
    [target],
    staging_dir="staging",
    delivery_dir="delivery",
)
print(result.delivered, result.counts, result.delivery_paths)
```

Only the fields listed in `excel_ops.delivery.CONTRACT_FIELDS` may be mapped:
`record_id`, `location`, `event_date`, `identifier`, `category`, `confidence`,
`source`, `source_file`, `source_sheet`, `source_row`, `source_region`. Mapping
anything else is a blocking planning error, not an empty cell.

## Dry run first

`plan_delivery(...)` — or `run_delivery(..., dry_run=True)` — returns the same
result object with the plan filled in and no file created: input count, target
template, field mapping, staging and delivery paths, expected written and review
counts per target, unresolved ambiguities, and blocking items.

## Reusing human decisions

Ambiguous matches are grouped into one confirmation batch. A decision can be
saved as a project Recipe and reused on later runs:

```python
from excel_ops import decide, save_project_recipe

pending = [item for item in plan.confirmation_batch.items if item.status == "pending"]
save_project_recipe(
    [decide(pending[0].ambiguity, "North Warehouse", scope="project")],
    "recipes/project-recipe.json",
)
run_delivery(..., recipe_path="recipes/project-recipe.json")
```

## CLI

```bash
excel-ops deliver delivery-plan.json --dry-run
excel-ops deliver delivery-plan.json --recipe recipes/project-recipe.json --result run.json
```

The configuration file declares inputs and targets, with paths relative to the
configuration file:

```json
{
  "inputs": ["sources/week-37.xlsx", "sources/extraction.json"],
  "staging_dir": "staging",
  "delivery_dir": "delivery",
  "targets": [
    {
      "key": "North Warehouse",
      "aliases": ["North Warehouse", "北区仓库"],
      "template": "templates/north-template.xlsx",
      "sheet": "Inspections",
      "header_row": 4,
      "data_start_row": 5,
      "field_columns": {
        "record_id": "A",
        "location": "B",
        "event_date": "C",
        "identifier": "D",
        "category": "E",
        "source": "F"
      },
      "required_fields": ["record_id", "location", "event_date", "identifier", "category", "source"],
      "period_expectations": [
        { "sheet": "Inspections", "cell": "B2", "expected": "2026-09-07 to 2026-09-13" }
      ]
    }
  ]
}
```

A successful `--dry-run` always exits with status 0, even when `delivered` is `false` (a dry-run never delivers by design, so that's not a failure). A real (non-dry-run) run exits non-zero only when it produced failures, or when it ends up undelivered.

## Failure codes

| Code | Meaning |
|---|---|
| `unknown_layout` | An input header could not be identified; nothing was written. |
| `unreadable_input` | An input file could not be read or parsed. |
| `unreadable_recipe` | The project Recipe is missing its format marker or is corrupt. |
| `plan_blocked` | The plan has blocking items, such as a mapped field outside the record contract, an unmapped record ID, or a missing template. |
| `template_write_failed` | The declared mapping does not fit the template, for example a missing sheet. |
| `incomplete_write` | The writer skipped a mapped cell; no partially written row is delivered. |
| `verification_failed` | The reopened file failed a blocking check. The verification report lists every finding. |

## Not covered by this pipeline

- Whole-run idempotency is opt-in and lives in its own module: pass
  `idempotency=IdempotencyOptions(...)` (or `excel-ops deliver --run-state`) and
  an unchanged rerun of a successful run short-circuits to a no-op before any
  matching, write or verification happens. See
  [Idempotent execution](idempotency.md). Without it, the pipeline behaves as
  described here: re-running the same confirmed input still re-ingests,
  re-matches and re-plans, and the per-record deduplication then withholds the
  records the target already holds. The cloud-target revision conflict check
  from [#19](https://github.com/Schlaflied/excel-ops/issues/19) is a
  forward-compatible interface, not a working cloud connector.
- Source and verification Manifests ([#20](https://github.com/Schlaflied/excel-ops/issues/20)),
  currency and precision rules ([#23](https://github.com/Schlaflied/excel-ops/issues/23)),
  and multi-format export ([#22](https://github.com/Schlaflied/excel-ops/issues/22)).
- Cloud connectors, multi-tab orchestration, cross-source fact checking, and
  regional rule calculations.
- openpyxl does not calculate formulas. Declared formula checks are static, and
  the run reports `recalculation_not_verified` as an explicit warning rather than
  claiming a recalculated workbook.
