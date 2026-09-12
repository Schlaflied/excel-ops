# Connector behavior

Connectors implement the same conceptual operations: identify, snapshot, read, plan, write, verify, and create a recoverable revision or copy.

## Local XLSX and CSV

Use deterministic file parsing and writing. Preserve the original file. For XLSX, reopen the exported workbook and validate sheets, ranges, formulas, record counts, and required fields.

## Local cloud-sync folders

Dropbox, Google Drive, OneDrive, or similar desktop clients can expose ordinary local files. Confirm the file is stable before reading. Before replacement, compare the current size, modification time, and hash with the snapshot. Write a sibling output first and let the sync client upload it. Do not claim cloud persistence until the synced state can be observed.

## Dropbox API

Treat Dropbox as versioned file storage, not a cell API. Download a revision, edit and verify locally, then upload with the expected revision or conflict-safe mode. On conflict, stop and preserve both versions rather than overwriting collaborator changes.

## Google Sheets

Read the required ranges and relevant metadata, create a compact change set, batch compatible updates, then reread affected ranges. Detect collaborative changes between snapshot and write where practical. Prefer stable metadata or business IDs over row numbers.

## Feishu

Distinguish electronic spreadsheets from multidimensional tables. Use the connector matching the actual document type and its scopes. Never treat a Bitable record ID as a spreadsheet row identity or vice versa.

## WPS

Distinguish traditional spreadsheets, smart spreadsheets, and multidimensional tables; their file IDs, endpoints, and scopes are not interchangeable. Choose the connector from the actual document type before reading or writing.

## Authentication

Use OAuth, an installed app connector, service account, or platform credential store appropriate to the user's environment. Request the narrowest practical scopes. Never place access or refresh tokens in project files, recipes, logs, workbooks, or Git.
