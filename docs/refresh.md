[中文](refresh.zh-CN.md) | English

# Read-only system update checks

Issue [#44](https://github.com/Schlaflied/excel-ops/issues/44) is split into two PRs. This first PR adds only a read-only system-file manifest and update check. Workbook processing remains in Python.

Run from the repository root:

```text
node refresh.mjs check
node refresh.mjs preview
```

Both commands print one JSON object and leave files untouched. `check` reports `up-to-date`, `update-available`, `local-system-changed`, `offline`, `local-manifest-invalid`, or `remote-manifest-invalid`. `preview` also lists system files that a future update would add, update, or remove, plus local conflicts. A network failure returns `offline` without blocking spreadsheet work.

`refresh-manifest.json` explicitly lists system files and their SHA-256 hashes. The remote comparison uses the manifest published on GitHub `main`; equal package versions do not hide file changes. The check reads only listed local system files. It never enumerates user files, and the JSON output includes no file contents or unknown user paths.

Company templates, Template Profiles, project Recipes, review decisions, source workbooks, review packs, Manifests, outputs, connector configuration, credentials, and locally added examples or business rules are user owned. They are outside the system list. An unexpected local file at a newly listed system path is reported as a conflict.

Maintainers should update the explicit file list when adding a system file, then run `node scripts/update-refresh-manifest.mjs`. Run `node scripts/update-refresh-manifest.mjs --check` in validation. The manifest does not hash itself.

`apply`, `rollback`, and `agent-check` are reserved for the second PR under #44. This PR does not install dependencies, call Git, modify files during a check, or alter the Python CLI's JSON output.
