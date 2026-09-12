from __future__ import annotations

import argparse
import json

from .pipeline import run_pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Turn reviewed extraction records into an auditable Excel delivery")
    parser.add_argument("input", help="Input .xlsx, .xlsm, .csv, or provider-neutral image extraction .json")
    parser.add_argument("output", help="Output .xlsx path")
    parser.add_argument("--confidence-threshold", type=float, default=0.85)
    parser.add_argument("--locale", help="Locale hint such as en-US, en-GB, zh-CN, or de-DE")
    args = parser.parse_args()
    print(json.dumps(run_pipeline(args.input, args.output, args.confidence_threshold, args.locale), ensure_ascii=False))


if __name__ == "__main__":
    main()
