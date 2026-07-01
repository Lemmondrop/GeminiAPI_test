"""
Phase 4 general: train a discount-rate model after excluding bio rows.

Input:
  - training_features_general.csv

Outputs:
  - model_general.pkl
  - feature_importance_general.csv
  - predictions_cv_general.csv
  - model_metrics_general.csv

This file intentionally coexists with phase4_train_model.py and
phase4_train_model_extracted.py.
"""

import os
import warnings

import joblib
import numpy as np
import pandas as pd
import xgboost as xgb
from sklearn.model_selection import KFold, cross_val_predict, cross_val_score

warnings.filterwarnings("ignore")


INPUT_CSV = os.environ.get("TRAINING_CSV", "training_features_general.csv")
MODEL_OUT = os.environ.get("MODEL_OUT", "model_general.pkl")
IMPORTANCE_OUT = os.environ.get("IMPORTANCE_OUT", "feature_importance_general.csv")
PREDICTIONS_OUT = os.environ.get("PREDICTIONS_OUT", "predictions_cv_general.csv")
METRICS_OUT = os.environ.get("METRICS_OUT", "model_metrics_general.csv")

TARGET = "discount_rate_mid_pct"

BASE_FEATURES = [
    "price_vs_mid",
    "band_width_pct",
    "year",
    "month",
    "log_per_stock",
    "log_offering",
    "is_konex",
    "price_band_low",
    "price_band_high",
]

EXTRA_NUMERIC_FEATURES = [
    "net_income_bn",
    "dart_revenue_bn",
    "dart_op_income_bn",
    "dart_net_income_bn",
    "dart_assets_bn",
    "dart_equity_bn",
    "dart_debt_ratio",
    "dart_op_margin",
    "peer_per",
]

MISSING_FLAG_FEATURES = [
    "per_stock_value_missing",
    "offering_amount_missing",
    "net_income_missing",
    "peer_per_missing",
    "dart_financial_missing",
]


def _to_numeric(df: pd.DataFrame, cols: list[str]) -> None:
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")


def load_training_data(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, encoding="utf-8-sig")

    numeric_cols = [
        TARGET,
        "discount_rate_low_pct",
        "discount_rate_high_pct",
        "price_band_low",
        "price_band_high",
        "final_offering_price",
        "price_vs_mid",
        "per_stock_value",
        "offering_amount",
        "net_income_bn",
        "dart_revenue_bn",
        "dart_op_income_bn",
        "dart_net_income_bn",
        "dart_assets_bn",
        "dart_equity_bn",
        "dart_debt_ratio",
        "dart_op_margin",
        "peer_per",
    ]
    _to_numeric(df, numeric_cols)

    required = [
        TARGET,
        "listing_date",
        "price_vs_mid",
        "price_band_low",
        "price_band_high",
    ]
    train = df.dropna(subset=[c for c in required if c in df.columns]).copy()
    train = train[train["price_band_low"] > 0]
    train = train[train["price_band_high"] > train["price_band_low"]]

    return train


def engineer_features(train: pd.DataFrame) -> tuple[pd.DataFrame, list[str], dict, list[str]]:
    train["listing_date"] = pd.to_datetime(train["listing_date"], errors="coerce")
    train = train.dropna(subset=["listing_date"]).copy()

    train["year"] = train["listing_date"].dt.year
    train["month"] = train["listing_date"].dt.month
    train["band_width_pct"] = (
        (train["price_band_high"] - train["price_band_low"])
        / train["price_band_low"]
        * 100
    ).round(4)

    train["per_stock_value_missing"] = train["per_stock_value"].isna().astype(int)
    train["offering_amount_missing"] = train["offering_amount"].isna().astype(int)
    train["net_income_missing"] = train["net_income_bn"].isna().astype(int)
    train["peer_per_missing"] = train["peer_per"].isna().astype(int)
    train["dart_financial_missing"] = (
        train.get("dart_net_income_bn", pd.Series(index=train.index, dtype=float)).isna()
    ).astype(int)

    train["log_per_stock"] = np.log1p(
        train["per_stock_value"].fillna(train["per_stock_value"].median())
    )
    train["log_offering"] = np.log1p(
        train["offering_amount"].fillna(train["offering_amount"].median())
    )
    train["is_konex"] = (train.get("track_type", "") == "konex_transfer").astype(int)

    segment_features = sorted(
        c for c in train.columns if c.startswith("segment_") and c != "segment_unknown"
    )
    if "segment_unknown" in train.columns:
        segment_features.append("segment_unknown")

    for col in segment_features:
        train[col] = pd.to_numeric(train[col], errors="coerce").fillna(0).astype(int)

    available_extra = [c for c in EXTRA_NUMERIC_FEATURES if c in train.columns]
    raw_features = BASE_FEATURES + segment_features + available_extra + MISSING_FLAG_FEATURES

    medians = train[raw_features].median(numeric_only=True)
    X = train[raw_features].fillna(medians)

    constant_features = [
        col for col in raw_features if X[col].nunique(dropna=False) <= 1
    ]
    features = [col for col in raw_features if col not in constant_features]
    X = X[features]
    medians = medians.reindex(features)

    return X, features, medians.to_dict(), constant_features


def main() -> None:
    train = load_training_data(INPUT_CSV)
    X, features, medians, constant_features = engineer_features(train)
    y = train[TARGET].astype(float)

    print(f"Input: {INPUT_CSV}")
    print(f"Training rows: {len(train)}")
    print(f"Features: {len(features)}")
    print(f"Segment features: {len([f for f in features if f.startswith('segment_')])}")
    print(f"Dropped constant features: {len(constant_features)}")
    if constant_features:
        print("  " + ", ".join(constant_features))
    print(f"Missing-imputed features: {(X.isna().sum() > 0).sum()}")
    print()

    params = dict(
        n_estimators=150,
        max_depth=2,
        learning_rate=0.04,
        subsample=0.85,
        colsample_bytree=0.85,
        min_child_weight=8,
        reg_alpha=0.2,
        reg_lambda=1.2,
        random_state=42,
        objective="reg:squarederror",
    )

    kf = KFold(n_splits=5, shuffle=True, random_state=42)
    model_cv = xgb.XGBRegressor(**params)

    rmse_scores = np.sqrt(
        -cross_val_score(model_cv, X, y, cv=kf, scoring="neg_mean_squared_error")
    )
    mae_scores = -cross_val_score(model_cv, X, y, cv=kf, scoring="neg_mean_absolute_error")
    r2_scores = cross_val_score(model_cv, X, y, cv=kf, scoring="r2")
    y_pred_cv = cross_val_predict(model_cv, X, y, cv=kf)

    baseline_rmse = y.std()
    improvement_pct = (1 - rmse_scores.mean() / baseline_rmse) * 100

    print("=" * 50)
    print("5-Fold CV results")
    print("=" * 50)
    print(f"  RMSE : {rmse_scores.mean():.3f} +/- {rmse_scores.std():.3f} %p")
    print(f"  MAE  : {mae_scores.mean():.3f} +/- {mae_scores.std():.3f} %p")
    print(f"  R2   : {r2_scores.mean():.3f} +/- {r2_scores.std():.3f}")
    print(f"  Baseline RMSE: {baseline_rmse:.3f} %p")
    print(f"  Improvement : {improvement_pct:.1f}%")
    print()

    model_final = xgb.XGBRegressor(**params)
    model_final.fit(X, y)

    importance = pd.DataFrame(
        {
            "feature": features,
            "importance": model_final.feature_importances_,
        }
    )
    total_importance = importance["importance"].sum()
    importance["importance_pct"] = (
        importance["importance"] / total_importance * 100 if total_importance else 0
    )
    importance = importance.sort_values("importance", ascending=False).round(4)

    print("Top feature importance")
    for _, row in importance.head(20).iterrows():
        bar = "#" * int(row["importance_pct"] / 3)
        print(f"  {row['feature']:28s}: {bar:34s} {row['importance_pct']:.1f}%")
    print()

    save_info = {
        "model": model_final,
        "features": features,
        "medians": medians,
        "params": params,
        "n_train": len(train),
        "cv_rmse": round(float(rmse_scores.mean()), 4),
        "cv_mae": round(float(mae_scores.mean()), 4),
        "cv_r2": round(float(r2_scores.mean()), 4),
        "input_csv": INPUT_CSV,
        "version": "general_non_bio",
        "dropped_constant_features": constant_features,
    }
    joblib.dump(save_info, MODEL_OUT)
    print(f"Saved model: {MODEL_OUT}")

    importance.to_csv(IMPORTANCE_OUT, index=False, encoding="utf-8-sig")
    print(f"Saved feature importance: {IMPORTANCE_OUT}")

    cv_cols = [
        c
        for c in [
            "company_name",
            "listing_date",
            TARGET,
            "crawling_sector",
            "crawling_industry",
            "data_quality_flag",
        ]
        if c in train.columns
    ]
    cv_df = train[cv_cols].copy()
    cv_df["y_pred"] = y_pred_cv.round(4)
    cv_df["residual"] = (cv_df[TARGET] - cv_df["y_pred"]).round(4)
    cv_df["abs_error"] = cv_df["residual"].abs().round(4)
    cv_df = cv_df.sort_values("abs_error", ascending=False)
    cv_df.to_csv(PREDICTIONS_OUT, index=False, encoding="utf-8-sig")
    print(f"Saved CV predictions: {PREDICTIONS_OUT}")

    metrics = pd.DataFrame(
        [
            {
                "input_csv": INPUT_CSV,
                "n_train": len(train),
                "n_features": len(features),
                "n_segment_features": len([f for f in features if f.startswith("segment_")]),
                "n_dropped_constant_features": len(constant_features),
                "cv_rmse_mean": rmse_scores.mean(),
                "cv_rmse_std": rmse_scores.std(),
                "cv_mae_mean": mae_scores.mean(),
                "cv_mae_std": mae_scores.std(),
                "cv_r2_mean": r2_scores.mean(),
                "cv_r2_std": r2_scores.std(),
                "baseline_rmse": baseline_rmse,
                "improvement_pct": improvement_pct,
            }
        ]
    ).round(4)
    metrics.to_csv(METRICS_OUT, index=False, encoding="utf-8-sig")
    print(f"Saved metrics: {METRICS_OUT}")

    print()
    print("=" * 50)
    print("Top 10 residuals")
    print("=" * 50)
    print(
        cv_df.head(10)[
            ["company_name", "listing_date", TARGET, "y_pred", "residual"]
        ].to_string(index=False)
    )


if __name__ == "__main__":
    main()
