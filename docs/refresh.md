[中文](refresh.zh-CN.md) | English

# System update checks, updates, and rollback

Issue [#44](https://github.com/Schlaflied/excel-ops/issues/44) is split into two PRs. The first added the read-only system-file manifest with `check` and `preview`. The second adds the explicitly confirmed `apply`, the `rollback` of that update, and the cached `agent-check`. Workbook processing remains in Python.

Run from the repository root:

```text
node refresh.mjs check
node refresh.mjs preview
node refresh.mjs apply --confirm
node refresh.mjs rollback
node refresh.mjs agent-check [--force]
```

Every command prints one JSON object on stdout. No command installs dependencies, runs Git, starts a background process, or changes the Python CLI's own output.

## check and preview

Both leave files untouched. `check` reports `up-to-date`, `update-available`, `local-system-changed`, `offline`, `local-manifest-invalid`, or `remote-manifest-invalid`. `preview` also lists system files that an update would add, update, or remove, plus local conflicts. A network failure returns `offline` without blocking spreadsheet work.

`refresh-manifest.json` explicitly lists system files and their SHA-256 hashes. The remote comparison uses the manifest published on GitHub `main`; equal package versions do not hide file changes. The check reads only listed local system files. It never enumerates user files, and the JSON output includes no file contents or unknown user paths.

## apply --confirm

Without `--confirm` the command returns `confirmation-required` with exit code 2 and changes nothing: no download, no write, no backup, no removal. There is no partial or implicit apply.

With `--confirm` the update runs in this order:

1. Run the `preview` comparison. `offline`, `local-manifest-invalid`, and `remote-manifest-invalid` stop the run before any download.
2. Stop with `conflict-blocked` when a local system file at an incoming path differs from the recorded hash. Conflicts are reported per file and resolved by you; `apply` never merges or picks a side.
3. Download each listed system file and verify its SHA-256 against the published manifest. A failure returns `download-failed` or `download-verification-failed` before the working tree is touched.
4. Copy every file about to change or be removed, plus the current manifest, into `.refresh/backups/<run-id>/files/`, then write the state record `.refresh/state.json`. The record exists before the first write, so an interrupted run is still recoverable.
5. Write the verified files, remove the files the published manifest dropped, and replace `refresh-manifest.json`.
6. Run a minimal doctor check: read every written file back, compare its hash, confirm removals are gone, and revalidate the local manifest.

A successful run returns `applied` with the file-level `added`, `updated`, and `removed` lists and `integrity.ok`. A failed doctor check or a failed write returns `apply-failed` with per-file diagnostics, keeps those diagnostics in `.refresh/state.json`, and sets `rollbackAvailable`. Only paths listed in the published manifest are written, so `.refresh/` and every user path stay outside the update.

## rollback

`rollback` reverts only the most recent update that `apply` itself recorded. It is not a general restore tool: it reads `.refresh/state.json`, touches only the paths listed there, and restores their bytes from that run's backup.

A file changed after the apply is left exactly as it is and reported under `skipped` with `changed-after-apply`. Files added after the apply are never seen at all. Statuses are `rolled-back`, `rolled-back-partial` (something was skipped), `rolled-back-unverified` (the integrity check did not come back clean), `already-rolled-back`, `no-apply-recorded`, or `state-invalid`. A path that could not be written back, for example because it is locked, is reported under `skipped` with `restore-failed` instead of aborting the run. A partial rollback is recorded as `rolled-back-partial` in `.refresh/state.json`, so running `rollback` again retries the skipped paths once their cause is resolved; entries already restored match the recorded hashes and are restored to the same bytes. Only a complete rollback is recorded as `rolled-back` and reports `already-rolled-back` on a second run. After restoring, the integrity check reruns and its result is returned as `integrity`.

## agent-check

`agent-check` is the once-per-session read-only reminder for an agent's first Excel-Ops call. It runs the same `check` comparison, caches a successful result in `.refresh/agent-check.json` for 24 hours, and returns `updateAvailable` with advice to tell the user. `--force` bypasses the cache.

It never calls `apply`. A missing, stale, unwritable, or corrupted cache falls through to one fresh read-only check, and a failed check returns `offline`, so the normal Excel-Ops workflow is never blocked. Its JSON is a separate command output and does not enter any other command's stdout contract.

## User data boundary

Company templates, Template Profiles, project Recipes, review decisions, confirmed mappings, source workbooks, review packs, Manifests, deliveries, output directories, connector configuration, credentials, customer and employee data, and locally added examples or business rules are user owned. They are outside the system list, so `apply` and `rollback` never update, delete, upload, log, or copy them into a backup or diagnostic record. Backups contain only listed system files. An unexpected local file at a newly listed system path is reported as a conflict instead of being overwritten.

Maintainers should update the explicit file list when adding a system file, then run `node scripts/update-refresh-manifest.mjs`. Run `node scripts/update-refresh-manifest.mjs --check` in validation. The manifest does not hash itself. `.refresh/` holds local run data only and is ignored by Git.
