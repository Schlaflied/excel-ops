from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Sequence
from datetime import date
from pathlib import Path

from .delivery import load_delivery_targets, run_delivery
from .idempotency import IdempotencyOptions
from .pipeline import run_pipeline
from .schema_drift import save_confirmed_mapping, write_drift_report
from .workdir import (
    DEFAULT_STABILITY_WINDOW_SECONDS,
    ClassificationOverride,
    PeriodWindow,
    ScanScope,
    format_dry_run,
    save_workdir_recipe,
    scan_workdir,
)


def main(argv: Sequence[str] | None = None) -> None:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments and arguments[0] == "deliver":
        deliver = argparse.ArgumentParser(
            description="Ingest, match, confirm, write a template copy, and verify the delivered file"
        )
        deliver.add_argument("config", help="JSON delivery configuration with inputs and targets")
        deliver.add_argument("--dry-run", action="store_true", help="Return the plan without touching files")
        deliver.add_argument("--recipe", help="Project Recipe JSON used to reuse confirmed decisions")
        deliver.add_argument("--result", help="Write the machine-readable run result to this JSON file")
        deliver.add_argument(
            "--run-state",
            nargs="?",
            const="",
            help=(
                "Enable whole-run idempotency: an unchanged rerun of a successful run is a no-op. "
                "Optionally give the run-record path (default: <delivery_dir>/.excel-ops/idempotency.json)"
            ),
        )
        deliver.add_argument(
            "--task-key",
            help="Name this periodic task, so several tasks can share one run-record file",
        )
        deliver.add_argument(
            "--template-profile-version",
            help="Template Profile version marker folded into the run fingerprint",
        )
        args = deliver.parse_args(arguments[1:])
        if args.task_key and args.run_state is None:
            deliver.error("--task-key requires --run-state")
        config_path = Path(args.config)
        payload = json.loads(config_path.read_text(encoding="utf-8"))
        targets, options = load_delivery_targets(payload, base_dir=config_path.parent)
        if args.recipe:
            options["recipe_path"] = Path(args.recipe)
        inputs = options.pop("inputs")
        if args.run_state is not None:
            options["idempotency"] = IdempotencyOptions(
                state_path=args.run_state or None,
                task_key=args.task_key,
                template_profile_version=args.template_profile_version,
            )
        result = run_delivery(inputs, targets, dry_run=args.dry_run, **options)
        report = result.to_dict()
        if args.result:
            Path(args.result).write_text(
                json.dumps(report, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8"
            )
        print(json.dumps(report, ensure_ascii=False, default=str))
        # A no-op is a success: the declared delivery is already in place and
        # nothing changed, so there is nothing to deliver and nothing failed.
        if result.failures or (not args.dry_run and not result.delivered and not result.no_op):
            sys.exit(1)
        return

    if arguments and arguments[0] == "scan-workdir":
        scan = argparse.ArgumentParser(
            description="Read-only scan of authorized working directories: what is included, excluded, and why"
        )
        scan.add_argument("directory", help="An authorized directory to scan")
        scan.add_argument(
            "--also-allow",
            action="append",
            default=[],
            metavar="DIR",
            help="Additional authorized directory; nothing outside the allowlist is accessed",
        )
        scan.add_argument("--no-recursive", action="store_true", help="Scan only the top level of each root")
        scan.add_argument("--period-start", help="Current business period start, YYYY-MM-DD")
        scan.add_argument("--period-end", help="Current business period end, YYYY-MM-DD")
        scan.add_argument(
            "--stability-window",
            type=float,
            default=DEFAULT_STABILITY_WINDOW_SECONDS,
            help="Seconds a file's size and modification time must already be unchanged",
        )
        scan.add_argument("--recipe", help="Workdir Recipe JSON with saved classification overrides")
        scan.add_argument(
            "--override",
            action="append",
            default=[],
            metavar="PATH=CLASSIFICATION[:DISPOSITION]",
            help="Override one relative path or glob; repeatable",
        )
        scan.add_argument("--save-recipe", help="Write the supplied overrides to a reusable Recipe file")
        scan.add_argument("--result", help="Write the machine-readable scan report to this JSON file")
        scan.add_argument("--json", action="store_true", help="Print JSON instead of the readable dry-run report")
        args = scan.parse_args(arguments[1:])
        if bool(args.period_start) != bool(args.period_end):
            scan.error("--period-start and --period-end must be supplied together")
        window = None
        if args.period_start:
            window = PeriodWindow.of(
                date.fromisoformat(args.period_start), date.fromisoformat(args.period_end)
            )
        overrides = []
        for value in args.override:
            if "=" not in value:
                scan.error("--override must use PATH=CLASSIFICATION[:DISPOSITION]")
            target, decision = value.split("=", 1)
            classification, _, disposition = decision.partition(":")
            overrides.append(
                ClassificationOverride(
                    path=target.strip(),
                    classification=classification.strip(),
                    disposition=disposition.strip() or None,
                    source="cli",
                )
            )
        scope = ScanScope.of(
            [args.directory, *args.also_allow], recursive=not args.no_recursive
        )
        result = scan_workdir(
            scope,
            period_window=window,
            stability_window_seconds=args.stability_window,
            overrides=overrides,
            recipe_path=args.recipe,
        )
        if args.save_recipe:
            save_workdir_recipe(overrides, args.save_recipe)
        report = result.to_dict()
        if args.result:
            Path(args.result).write_text(
                json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
            )
        print(json.dumps(report, ensure_ascii=False) if args.json else format_dry_run(result))
        return

    if arguments and arguments[0] == "schema-drift":
        drift = argparse.ArgumentParser(description="Compare workbook schemas")
        drift.add_argument("inputs", nargs="+", help="Baseline workbook followed by candidates")
        drift.add_argument("--output", required=True, help="Machine-readable JSON report")
        drift.add_argument("--mapping", help="JSON file containing confirmed_mappings")
        drift.add_argument("--confirm", action="append", default=[], metavar="CANDIDATE=BASELINE", help="Persist a human-confirmed mapping before comparison")
        drift.add_argument("--key-field", help="Field used to detect row-granularity changes")
        args = drift.parse_args(arguments[1:])
        if args.confirm and not args.mapping:
            drift.error("--confirm requires --mapping")
        for value in args.confirm:
            if "=" not in value:
                drift.error("--confirm must use CANDIDATE=BASELINE")
            source, target = value.split("=", 1)
            save_confirmed_mapping(args.mapping, source.strip(), target.strip())
        output = write_drift_report(args.inputs, args.output, mapping_path=args.mapping, key_field=args.key_field)
        print(json.dumps({"output": str(output)}, ensure_ascii=False))
        return

    parser = argparse.ArgumentParser(description="Turn reviewed extraction records into an auditable Excel delivery")
    parser.add_argument("input", help="Input .xlsx, .xlsm, .csv, or provider-neutral image extraction .json")
    parser.add_argument("output", help="Output .xlsx path")
    parser.add_argument("--confidence-threshold", type=float, default=0.85)
    parser.add_argument("--locale", help="Locale hint such as en-US, en-GB, zh-CN, or de-DE")
    args = parser.parse_args(arguments)
    if not args.input or not args.output:
        parser.error("input and output are required")
    print(json.dumps(run_pipeline(args.input, args.output, args.confidence_threshold, args.locale), ensure_ascii=False))


if __name__ == "__main__":
    main()
