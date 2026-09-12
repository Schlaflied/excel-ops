# Shared operating rules

The durable source of truth is the user's original file or cloud document plus the recorded operation plan, not chat memory. Keep six states distinct:

1. original source;
2. extracted candidate;
3. normalized candidate;
4. accepted or human-confirmed record;
5. persisted workbook mutation;
6. independently verified delivery.

Automate routine work end to end. Human review should contain only records that cannot be safely resolved, with the source value, proposed value, reason, confidence, and available choices visible together.

Use stable business keys or generated record IDs for idempotency. Never rely on current row number after sorting, filtering, insertion, or cloud collaboration.

For recurring work, produce a recipe that records inputs, schema, mapping rules, matching policy, validation gates, destination, and conflict behavior without credentials or private source contents.

Report a compact status: sources, assumptions, changed artifact or ranges, accepted/review/rejected counts, verification evidence, unresolved decisions, and reusable next run.
