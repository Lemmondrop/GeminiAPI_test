"""
Merge discount-rate source CSVs into one versioned input file.

The pipeline often starts from:
  - a baseline peer_group_discount_rates.csv
  - newly extracted discount_rate_table pages

This utility keeps the merge rule explicit and reproducible for each dataset
version folder under test_data/.
"""

import argparse
from pathlib import Path

import pandas as pd


KEY_COLS = ["company_name", "listing_date"]


def read_source(path: str, priority: int) -> pd.DataFrame:
    df = pd.read_csv(path, encoding="utf-8-sig")
    df["_merge_source_path"] = Path(path).as_posix()
    df["_merge_priority"] = priority
    return df


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base", required=True, help="baseline discount-rate CSV")
    parser.add_argument("--add", nargs="+", required=True, help="newly extracted CSVs")
    parser.add_argument("--output", required=True)
    parser.add_argument(
        "--prefer-new",
        action="store_true",
        help="prefer rows from --add when company_name/listing_date duplicates exist",
    )
    args = parser.parse_args()

    frames = [read_source(args.base, 1 if args.prefer_new else 0)]
    frames.extend(read_source(path, 0 if args.prefer_new else 1) for path in args.add)
    merged = pd.concat(frames, ignore_index=True, sort=False)

    missing = [col for col in KEY_COLS if col not in merged.columns]
    if missing:
        raise SystemExit(f"Missing required columns: {missing}")

    merged = (
        merged.sort_values(KEY_COLS + ["_merge_priority"])
        .drop_duplicates(KEY_COLS, keep="first")
        .drop(columns=["_merge_priority"])
    )

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    merged.to_csv(output, index=False, encoding="utf-8-sig")

    print(f"Saved: {output}")
    print(f"Rows: {len(merged)}")
    print("\nBy merge source:")
    print(merged["_merge_source_path"].value_counts(dropna=False).to_string())


if __name__ == "__main__":
    main()
