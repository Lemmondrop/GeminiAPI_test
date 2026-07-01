"""
Build a non-bio general-purpose training dataset.

Default policy:
  - input : training_features_segmented.csv
  - output: training_features_general.csv
  - exclude rows where segment_bio == 1 or segment_healthcare == 1

This keeps the segmented dataset intact and creates a separate dataset for the
early-stage model that intentionally does not try to learn bio valuation cases.
"""

import argparse
from pathlib import Path

import pandas as pd


DEFAULT_INPUT = "training_features_segmented.csv"
DEFAULT_OUTPUT = "training_features_general.csv"
DEFAULT_EXCLUDED_OUTPUT = "training_features_general_excluded.csv"


def _as_flag(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0).astype(int)


def build_exclusion_mask(df: pd.DataFrame, mode: str) -> tuple[pd.Series, list[str]]:
    reasons: list[str] = []
    mask = pd.Series(False, index=df.index)

    if "segment_bio" in df.columns:
        bio_mask = _as_flag(df["segment_bio"]).eq(1)
        mask = mask | bio_mask
        reasons.append("segment_bio")

    if mode == "bio_healthcare" and "segment_healthcare" in df.columns:
        healthcare_mask = _as_flag(df["segment_healthcare"]).eq(1)
        mask = mask | healthcare_mask
        reasons.append("segment_healthcare")

    return mask, reasons


def add_exclusion_reason(df: pd.DataFrame, mode: str) -> pd.Series:
    reason = pd.Series("", index=df.index, dtype="object")

    bio_mask = pd.Series(False, index=df.index)
    if "segment_bio" in df.columns:
        bio_mask = _as_flag(df["segment_bio"]).eq(1)
        reason.loc[bio_mask] = "segment_bio"

    if mode == "bio_healthcare" and "segment_healthcare" in df.columns:
        healthcare_mask = _as_flag(df["segment_healthcare"]).eq(1)
        reason.loc[healthcare_mask & ~bio_mask] = "segment_healthcare"
        reason.loc[healthcare_mask & bio_mask] = "segment_bio+segment_healthcare"

    return reason


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create a general training dataset by excluding bio rows."
    )
    parser.add_argument("--input", default=DEFAULT_INPUT)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--excluded-output", default=DEFAULT_EXCLUDED_OUTPUT)
    parser.add_argument(
        "--mode",
        choices=["bio_healthcare", "bio_only"],
        default="bio_healthcare",
        help=(
            "bio_healthcare excludes segment_bio or segment_healthcare. "
            "bio_only excludes only segment_bio."
        ),
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    df = pd.read_csv(input_path, encoding="utf-8-sig")

    mask, reasons = build_exclusion_mask(df, args.mode)
    reason = add_exclusion_reason(df, args.mode)

    general = df.loc[~mask].copy()
    excluded = df.loc[mask].copy()
    excluded.insert(
        len(excluded.columns),
        "general_exclusion_reason",
        reason.loc[mask].values,
    )

    general.to_csv(args.output, index=False, encoding="utf-8-sig")
    excluded.to_csv(args.excluded_output, index=False, encoding="utf-8-sig")

    print("=" * 60)
    print("Build general training features")
    print("=" * 60)
    print(f"Input              : {args.input}")
    print(f"Mode               : {args.mode}")
    print(f"Exclusion flags    : {', '.join(reasons) if reasons else '(none found)'}")
    print(f"Total rows         : {len(df)}")
    print(f"General rows       : {len(general)}")
    print(f"Excluded rows      : {len(excluded)}")
    print(f"Saved general      : {args.output}")
    print(f"Saved excluded     : {args.excluded_output}")

    if len(excluded) > 0:
        print()
        print("[Excluded reason counts]")
        for key, cnt in reason.loc[mask].value_counts().items():
            print(f"  {key:32s}: {cnt:4d}")

    if "discount_rate_mid_pct" in df.columns:
        print()
        print("[Target distribution]")
        for label, part in [("all", df), ("general", general), ("excluded", excluded)]:
            target = pd.to_numeric(part["discount_rate_mid_pct"], errors="coerce").dropna()
            if len(target) == 0:
                continue
            print(
                f"  {label:8s}: n={len(target):3d} "
                f"mean={target.mean():6.2f}%p "
                f"std={target.std():6.2f}%p "
                f"min={target.min():6.2f}%p "
                f"max={target.max():6.2f}%p"
            )


if __name__ == "__main__":
    main()
