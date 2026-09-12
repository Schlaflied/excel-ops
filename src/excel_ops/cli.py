from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Sequence

from .pipeline import run_pipeline
from .schema_drift import save_confirmed_mapping, write_drift_report


def main(argv: Sequence[str] | None = None) -> None:
    arguments = list(sys.argv[1:] if argv is None else argv)
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
