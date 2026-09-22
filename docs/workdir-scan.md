[中文](workdir-scan.zh-CN.md) | English

# Working-directory scan and current-period file selection

`scan_workdir(...)` answers one question before any spreadsheet work starts:
*which files in this working directory are this period's inputs, and which ones
must a human look at first?* It classifies every authorized file, reports why
each file was included or excluded, and never moves, renames, writes or deletes
anything.

```text
authorized roots (explicit allowlist)
  -> walk, re-checking containment per entry   (symlinks are not followed)
  -> sync-artifact check                        (lock / partial download / conflict copy)
  -> extension check                            (unreadable extension -> exclude)
  -> stability window                           (unstable -> review, never opened)
  -> name pattern -> workbook metadata -> modification time
  -> content-hash duplicate grouping            (identical bytes processed once)
  -> version-candidate grouping                 (look-alike names, no winner elected)
  -> user overrides / saved workdir Recipe
  -> include / review / exclude, with a reason per file
```

## Guarantees

- **Nothing outside the allowlist is accessed.** A scan takes an explicit list of
  authorized root directories. Containment is re-checked for every entry rather
  than inferred from its parent, and symlinks are reported as
  `symlink_not_followed` instead of being traversed — the same root-confinement
  idiom `refresh.mjs` uses for system files. Widening the scope is an explicit
  act; it is never inferred from the directory tree.
- **A file that is still being written or synced is never read as final.** Its
  size and modification time must already have been unchanged for the stability
  window (and must match the previous run's sample, when one is supplied) before
  the file is opened at all. An unstable file is not hashed, has no
  `content_hash`, and goes to `review` with reason `not_stable_yet`.
- **Duplicate content is detected by content hash, not by name.** Byte-identical
  files under different names form one `DuplicateGroup`; one representative stays
  `include` and the rest are excluded with `duplicate_content`. Because the bytes
  are identical, this decides which *name* to work under, not which content is
  right.
- **An undeterminable latest version is never auto-selected.** `final.xlsx`,
  `final (1).xlsx` and `final-final.xlsx` with differing content form a
  `VersionCandidateGroup` whose `selected` is always `None`; every member goes to
  `review`.
- **Cloud-sync leftovers are flagged, not ingested.** Lock files (`~$…`),
  incomplete downloads (`.crdownload`, `.part`, `.tmp`) and conflict copies
  (`… (conflicted copy from …)`, `…（冲突副本）`) never reach `include`.
- **A user override can relabel a file, but not defeat a safety floor.** An
  override cannot authorize a path outside the allowlist, cannot promote an
  unstable file to `include` (`override_refused_unstable_file`) and cannot
  promote a sync artifact (`override_refused_sync_artifact`).
- **The scan is read-only.** The only file this module ever writes is a workdir
  Recipe, and only when `save_workdir_recipe(...)` is called explicitly.

## Classifications and dispositions

Every file ends in exactly one classification and exactly one disposition.

| Classification | Meaning |
|---|---|
| `input` | A current-period input file. |
| `template` | An enterprise template, by extension, name, or workbook metadata. |
| `prior_delivery` | An earlier delivery or an archived historical file. |
| `review_return` | A returned review pack. |
| `unknown` | No signal was decisive; a human decides. |

| Disposition | Meaning |
|---|---|
| `include` | Safe to hand to the rest of the pipeline. |
| `review` | A human must decide before this file is used. |
| `exclude` | Deliberately not used in this run, with a recorded reason. |

The signal that produced a classification is recorded in `signal`
(`name_pattern`, `workbook_metadata`, `modification_time`, `extension`,
`stability_window`, `override`, or `none`), so a decision can always be audited.

## Usage

```python
from datetime import date
from excel_ops import PeriodWindow, ScanScope, format_dry_run, scan_workdir

scope = ScanScope.of(["/work/september", "/work/templates"], recursive=True)
result = scan_workdir(
    scope,
    period_window=PeriodWindow.of(date(2026, 9, 1), date(2026, 9, 30)),
    stability_window_seconds=5.0,
)

print(format_dry_run(result))            # readable include / review / exclude report
print(result.counts())                   # {'include': 3, 'review': 4, 'exclude': 6}
for item in result.review:
    print(item.relative_path, item.reason, item.notes)
```

`period_window` can also come from an already resolved business period:
`PeriodWindow.from_period(resolve_period("2026年9月"))`. Without a declared
period, data files cannot be separated from historical ones, so they are reported
as `unknown` / `review` with reason `no_period_window_declared` rather than
guessed into `input`.

To detect files that are still settling across two passes, feed the previous
run's observations back in:

```python
first = scan_workdir(scope, period_window=window)
second = scan_workdir(scope, period_window=window, previous_samples=first.stability_samples())
```

### Command line

```bash
excel-ops scan-workdir ./september --period-start 2026-09-01 --period-end 2026-09-30
excel-ops scan-workdir ./september --also-allow ./templates --no-recursive --json
excel-ops scan-workdir ./september --override "monthly.xlsx=template:exclude" --save-recipe workdir.json
excel-ops scan-workdir ./september --recipe workdir.json --result scan.json
```

The command is a dry run by definition: it reports what would be included,
excluded, and why, and touches nothing. `--save-recipe` is the only option that
writes a file, and it writes only the Recipe.

## Overriding a classification and saving it as a Recipe

```python
from excel_ops import override_from_entry, save_workdir_recipe, scan_workdir

result = scan_workdir(scope, period_window=window)
entry = result.entry("northern-monthly.xlsx")
override = override_from_entry(entry, "template", disposition="exclude", note="regional template")
save_workdir_recipe([override], "recipes/workdir.json")

reused = scan_workdir(scope, period_window=window, recipe_path="recipes/workdir.json")
```

A workdir Recipe is a separate file format (`excel-ops-workdir-recipe-v1`) from
the project Recipe used for field-level ambiguity decisions
(`excel-ops-recipe-v1`). The mechanism is deliberately the same shape — a
versioned JSON file, a format sentinel that refuses anything else, and a
`save_*` / `load_*` pair — but the payload is keyed by relative path rather than
by `field:question`, because a workdir decision is about a file rather than about
a data field. An override path may also be an `fnmatch` glob such as
`archive/*.xlsx`. Relative paths are always POSIX-separated, so a Recipe written
on Windows is reusable elsewhere.

## Not implemented

- **Not wired into the delivery pipeline.** `run_delivery(...)` and
  `excel-ops deliver` still take explicitly declared inputs. Feeding a scan's
  `include` set into the pipeline is a separate, later step.
- **No file is moved, renamed, or deleted.** The module classifies and reports.
- **Cloud-sync conflicts are not resolved.** A conflict copy is detected and
  routed to review; deciding which side wins is out of scope, exactly as it is
  for `refresh.mjs`.
- **Classification is evidence-based, not semantic.** The module reads filenames,
  extensions, modification times, content hashes, sheet titles and document
  properties. It does not read business content to decide what a file means.
