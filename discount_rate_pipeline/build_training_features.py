"""
Build training_features.csv from the original Halla dataset and expanded
peer-group feature extraction output.

The current model in phase4_train_model.py only needs these source columns
before feature engineering:
  - discount_rate_mid_pct
  - price_vs_mid
  - price_band_low
  - price_band_high
  - per_stock_value     (median-imputed in phase4 when missing)
  - offering_amount     (median-imputed in phase4 when missing)
  - track_type
  - listing_date

Rows with serious discount-rate cross-validation mismatch are excluded by
default because they are likely wrong rcept_no/table matches.
"""

import argparse
from pathlib import Path

import numpy as np
import pandas as pd


DEFAULT_INPUTS = [
    "peer_group_features_enriched.csv",
    "ipo_features_raw.csv",
]

REQUIRED_BASE_COLS = [
    "company_name",
    "listing_date",
    "discount_rate_mid_pct",
    "price_band_low",
    "price_band_high",
    "final_offering_price",
    "price_vs_mid",
]


def _read_csv(path: str) -> pd.DataFrame:
    df = pd.read_csv(
        path,
        encoding="utf-8-sig",
        dtype={"corp_code": str, "stock_code": str, "rcept_no": str},
    )
    df["source_dataset"] = Path(path).name
    return df


def _to_numeric(df: pd.DataFrame, cols: list[str]) -> None:
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")


def _normalize_stock_code_series(series: pd.Series) -> pd.Series:
    stock = (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\.0$", "", regex=True)
    )
    stock = stock.where(stock.str.fullmatch(r"[0-9A-Z]{1,6}"), "")
    stock = stock.where(stock.eq(""), stock.str.zfill(6))
    return stock


DEFAULT_STOCK_CODE_MAPS = [
    "peer_group_expanded_with_corp_code.csv",
    "peer_group_with_corp_code.csv",
    "ipo_with_corp_code_v2.csv",
    "ipo_with_corp_code.csv",
]


def _load_stock_code_map(paths: list[str] | None = None) -> pd.DataFrame:
    frames = []
    for path in paths or DEFAULT_STOCK_CODE_MAPS:
        if not Path(path).exists():
            continue
        df = pd.read_csv(path, encoding="utf-8-sig", dtype={"stock_code": str})
        if {"company_name", "listing_date", "stock_code"}.issubset(df.columns):
            frames.append(df[["company_name", "listing_date", "stock_code"]])
    if not frames:
        return pd.DataFrame(columns=["company_name", "listing_date", "stock_code"])
    mapping = pd.concat(frames, ignore_index=True)
    mapping["stock_code"] = _normalize_stock_code_series(mapping["stock_code"])
    mapping.loc[mapping["stock_code"].isin(["", "0", "nan"]), "stock_code"] = ""
    mapping = mapping[mapping["stock_code"].ne("")]
    return mapping.drop_duplicates(["company_name", "listing_date"], keep="first")


def build_training_features(
    inputs: list[str],
    output: str,
    include_mismatch: bool = False,
    stock_code_maps: list[str] | None = None,
) -> pd.DataFrame:
    frames = [_read_csv(path) for path in inputs]
    df = pd.concat(frames, ignore_index=True, sort=False)

    if "track_type" not in df.columns:
        df["track_type"] = "general"
    df["track_type"] = df["track_type"].fillna("general").replace("", "general")

    numeric_cols = [
        "discount_rate_low_pct",
        "discount_rate_high_pct",
        "discount_rate_mid_pct",
        "price_band_low",
        "price_band_high",
        "band_midpoint",
        "final_offering_price",
        "price_vs_mid",
        "net_income_mn",
        "net_income_bn",
        "peer_per",
        "per_stock_value",
        "total_shares",
        "offering_shares",
        "offering_amount",
        "discount_rate_xml_low",
        "discount_rate_xml_high",
        "discount_rate_xml_mid",
        "discount_rate_diff",
        "dart_revenue_bn",
        "dart_op_income_bn",
        "dart_net_income_bn",
        "dart_assets_bn",
        "dart_equity_bn",
        "dart_debt_ratio",
        "dart_op_margin",
    ]
    _to_numeric(df, numeric_cols)

    # Current phase4 train script requires these columns to be present before
    # median-imputing per_stock_value and offering_amount.
    mask = np.ones(len(df), dtype=bool)
    for col in REQUIRED_BASE_COLS:
        mask &= df[col].notna()

    if not include_mismatch and "data_quality_flag" in df.columns:
        mask &= df["data_quality_flag"].fillna("").ne("MISMATCH_CHECK")

    train = df[mask].copy()

    # Prefer expanded enriched rows over the older original rows on duplicates.
    source_priority = {
        "peer_group_features_enriched.csv": 0,
        "ipo_features_raw.csv": 1,
    }
    train["_source_priority"] = train["source_dataset"].map(source_priority).fillna(9)
    train = train.sort_values(
        ["company_name", "listing_date", "_source_priority"]
    ).drop_duplicates(["company_name", "listing_date"], keep="first")
    train = train.drop(columns=["_source_priority"])

    if "stock_code" not in train.columns:
        train["stock_code"] = ""
    train["stock_code"] = _normalize_stock_code_series(train["stock_code"])

    mapping = _load_stock_code_map(stock_code_maps).rename(
        columns={"stock_code": "_mapped_stock_code"}
    )
    if not mapping.empty:
        train = train.merge(mapping, on=["company_name", "listing_date"], how="left")
        missing_stock = train["stock_code"].eq("")
        train.loc[missing_stock, "stock_code"] = train.loc[
            missing_stock, "_mapped_stock_code"
        ].fillna("")
        train = train.drop(columns=["_mapped_stock_code"])

    # Keep a stable, analysis-friendly column order: original shared columns
    # first, then any enriched-only columns.
    preferred = [
        "company_name",
        "listing_date",
        "corp_code",
        "stock_code",
        "rcept_no",
        "report_nm",
        "track_type",
        "discount_rate_low_pct",
        "discount_rate_high_pct",
        "discount_rate_mid_pct",
        "price_band_low",
        "price_band_high",
        "band_midpoint",
        "final_offering_price",
        "price_vs_mid",
        "net_income_mn",
        "net_income_bn",
        "peer_per",
        "valuation_model",
        "per_stock_value",
        "total_shares",
        "offering_shares",
        "offering_amount",
        "discount_rate_xml_low",
        "discount_rate_xml_high",
        "discount_rate_xml_mid",
        "listing_market",
        "is_tech_special",
        "is_growth_special",
        "filled_key_fields",
        "extraction_status",
        "xml_size",
        "discount_rate_implied",
        "discount_rate_diff",
        "data_quality_flag",
        "net_income_type",
        "institutional_ratio",
        "retail_ratio",
        "discount_rate_xml_source",
        "dart_revenue_bn",
        "dart_op_income_bn",
        "dart_net_income_bn",
        "dart_assets_bn",
        "dart_equity_bn",
        "dart_bsns_year",
        "dart_reprt_code",
        "dart_debt_ratio",
        "dart_op_margin",
        "net_income_source",
        "source_dataset",
    ]
    ordered = [c for c in preferred if c in train.columns]
    ordered += [c for c in train.columns if c not in ordered]
    train = train[ordered]

    train.to_csv(output, index=False, encoding="utf-8-sig")
    return train


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--inputs", nargs="+", default=DEFAULT_INPUTS)
    parser.add_argument("--output", default="training_features.csv")
    parser.add_argument(
        "--include-mismatch",
        action="store_true",
        help="include rows with data_quality_flag == MISMATCH_CHECK",
    )
    parser.add_argument(
        "--stock-code-maps",
        nargs="*",
        default=None,
        help="CSV files used to backfill missing stock_code values",
    )
    args = parser.parse_args()

    train = build_training_features(
        inputs=args.inputs,
        output=args.output,
        include_mismatch=args.include_mismatch,
        stock_code_maps=args.stock_code_maps,
    )

    print(f"Saved: {args.output}")
    print(f"Rows: {len(train)}")
    print("\nBy source:")
    print(train["source_dataset"].value_counts(dropna=False).to_string())
    if "data_quality_flag" in train.columns:
        print("\nQuality:")
        print(train["data_quality_flag"].value_counts(dropna=False).to_string())
    print("\nCoverage:")
    for col in [
        "price_vs_mid",
        "price_band_low",
        "price_band_high",
        "per_stock_value",
        "offering_amount",
        "net_income_bn",
        "peer_per",
        "discount_rate_xml_mid",
    ]:
        if col in train.columns:
            print(f"  {col:24s}: {train[col].notna().sum():3d}/{len(train)}")


if __name__ == "__main__":
    main()
