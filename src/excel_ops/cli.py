from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Sequence
from pathlib import Path

from .delivery import load_delivery_targets, run_delivery
from .pipeline import run_pipeline
from .schema_drift import save_confirmed_mapping, write_drift_report


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
        args = deliver.parse_args(arguments[1:])
        config_path = Path(args.config)
        payload = json.loads(config_path.read_text(encoding="utf-8"))
        targets, options = load_delivery_targets(payload, base_dir=config_path.parent)
        if args.recipe:
            options["recipe_path"] = Path(args.recipe)
        inputs = options.pop("inputs")
        result = run_delivery(inputs, targets, dry_run=args.dry_run, **options)
        report = result.to_dict()
        if args.result:
            Path(args.result).write_text(
                json.dumps(report, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8"
            )
        print(json.dumps(report, ensure_ascii=False, default=str))
        if result.failures or (not args.dry_run and not result.delivered):
            sys.exit(1)
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
