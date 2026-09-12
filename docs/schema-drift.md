# Schema drift checks

Compare a baseline workbook with one or more incoming workbooks before any data is written:

```bash
excel-ops schema-drift baseline.xlsx week-1.xlsx week-2.xlsx \
  --key-field "Employee ID" --output drift-report.json
```

The JSON report separates added and removed fields, confirmed renames, order and type changes, required-field changes, small enum changes, and changes in rows per key. Potential renames remain in `mapping_suggestions` with `status: review`; they block append and join rather than being silently accepted.

After a person confirms a candidate-to-baseline mapping, persist it for later runs:

```bash
excel-ops schema-drift baseline.xlsx incoming.xlsx --output drift-report.json \
  --mapping mappings.json --confirm "Staff Number=Employee ID"
```

Only mappings in this file are accepted automatically. A newly added or removed field triggers review even when previously confirmed mappings still apply.
