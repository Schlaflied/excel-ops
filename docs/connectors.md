[中文](connectors.zh-CN.md) | English

# Connector contract

Excel-Ops keeps business logic independent from storage providers. A connector should provide equivalent behavior even when the underlying platform works differently.

| Connector | Read model | Write model | Conflict guard |
|---|---|---|---|
| Local XLSX/CSV | File snapshot | New file or authorized replace | Hash and modification time |
| Synced folder | Local file plus sync state | Sibling output, then client sync | Hash, modification time, sync observation |
| Dropbox API | Versioned file | Upload with expected revision | Remote revision |
| Google Sheets | Ranges and metadata | Batch range/document updates | Snapshot plus reread |
| Feishu Sheets | Sheet ranges | Platform API writes | Document version/reread where available |
| WPS Sheets | Type-specific ranges | Platform-specific batch writes | File type, version/reread |

Every connector must support a dry-run change set before mutation and a persisted-state verification after mutation. If a platform cannot provide safe replacement semantics, write a copy and return a reviewable handoff rather than claiming an in-place update.

Provider credentials are deployment concerns and must never enter recipes or repository fixtures.
