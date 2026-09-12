# Agent workflow

Excel-Ops uses the conversation as its interface. Users describe an outcome and provide sources; the agent owns internal routing and tool selection.

```text
User intent and sources
        ↓
Authority and destination check
        ↓
Snapshot and operation plan
        ↓
Extract → normalize → match
        ↓
Accepted ───────────────→ write
Review ─→ human choice ─→ write
Rejected ───────────────→ report
        ↓
Reread persisted destination
        ↓
Reconcile → package → reusable recipe
```

## Automation boundary

The agent should automate deterministic work without asking for confirmation at every step. It pauses only for a decision that can materially change the result: conflicting source authority, ambiguous identity, destructive replacement, unsupported formula behavior, or a cloud revision conflict.

An unreadable image value is not a request for the user to redo the entire task. The agent completes safe records and returns a minimal review queue for the uncertain cells.

## Example request

> Use `$excel-agent` to read these inspection screenshots, match them to last week's workbook, update the current period, and deliver a checked copy. Do not guess unreadable identifiers.

Expected result:

- updated workbook or cloud sheet;
- accepted, review, and rejected counts;
- source-to-output traceability;
- verification result from the persisted output;
- a reusable, credential-free recipe for the next period.
