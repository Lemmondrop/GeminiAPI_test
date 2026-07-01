"""
Phase 4 input-limited: train with only platform-fillable features.

Allowed feature groups:
  - per-stock valuation amount
  - applied PER / valuation multiple
  - applied net income
  - listing market
  - special listing flags
  - industry segment flags
  - recent financials
  - missing-value flags

Intentionally excluded:
  - final offering price
  - price_vs_mid
  - price band low/high
  - offering amount / offering shares
  - discount-rate-derived XML/check columns
  - listing date-derived year/month
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
MODEL_OUT = os.environ.get("MODEL_OUT", "model_input_limited.pkl")
IMPORTANCE_OUT = os.environ.get(
    "IMPORTANCE_OUT", "feature_importance_input_limited.csv"
)
PREDICTIONS_OUT = os.environ.get("PREDICTIONS_OUT", "predictions_cv_input_limited.csv")
METRICS_OUT = os.environ.get("METRICS_OUT", "model_metrics_input_limited.csv")

TARGET = "discount_rate_mid_pct"

NUMERIC_FEATURES = [
    "per_stock_value",
    "peer_per",
    "net_income_bn",
    "dart_revenue_bn",
    "dart_op_income_bn",
    "dart_net_income_bn",
    "dart_assets_bn",
    "dart_equity_bn",
    "dart_debt_ratio",
    "dart_op_margin",
    "is_tech_special",
    "is_growth_special",
]

CATEGORICAL_FEATURES = [
    "listing_market",
    "valuation_model",
]

MISSING_FLAG_FEATURES = [
    "per_stock_value_missing",
    "peer_per_missing",
    "net_income_missing",
    "dart_revenue_missing",
    "dart_op_income_missing",
    "dart_net_income_missing",
    "dart_assets_missing",
    "dart_equity_missing",
    "dart_financial_missing",
    "listing_market_missing",
    "valuation_model_missing",
]


def _to_numeric(df: pd.DataFrame, cols: list[str]) -> None:
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")


def load_training_data(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, encoding="utf-8-sig")
    _to_numeric(df, [TARGET] + NUMERIC_FEATURES)
    train = df.dropna(subset=[TARGET]).copy()
    return train


def _missing_flag(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series(1, index=df.index, dtype=int)
    if df[col].dtype == "object":
        return df[col].astype(str).str.strip().isin(["", "nan", "None"]).astype(int)
    return df[col].isna().astype(int)


def engineer_features(
    train: pd.DataFrame,
) -> tuple[pd.DataFrame, list[str], dict, dict, list[str]]:
    for col in NUMERIC_FEATURES:
        if col not in train.columns:
            train[col] = np.nan
        train[col] = pd.to_numeric(train[col], errors="coerce")

    for col in CATEGORICAL_FEATURES:
        if col not in train.columns:
            train[col] = ""
        train[col] = train[col].fillna("").astype(str).str.strip()

    train["log_per_stock"] = np.log1p(
        train["per_stock_value"].fillna(train["per_stock_value"].median())
    )
    train["log_peer_per"] = np.log1p(
        train["peer_per"].fillna(train["peer_per"].median()).clip(lower=0)
    )
    train["is_loss_company"] = (
        (train["net_income_bn"] < 0) | (train["dart_net_income_bn"] < 0)
    ).astype(int)

    train["per_stock_value_missing"] = _missing_flag(train, "per_stock_value")
    train["peer_per_missing"] = _missing_flag(train, "peer_per")
    train["net_income_missing"] = _missing_flag(train, "net_income_bn")
    train["dart_revenue_missing"] = _missing_flag(train, "dart_revenue_bn")
    train["dart_op_income_missing"] = _missing_flag(train, "dart_op_income_bn")
    train["dart_net_income_missing"] = _missing_flag(train, "dart_net_income_bn")
    train["dart_assets_missing"] = _missing_flag(train, "dart_assets_bn")
    train["dart_equity_missing"] = _missing_flag(train, "dart_equity_bn")
    train["dart_financial_missing"] = (
        train[
            [
                "dart_revenue_missing",
                "dart_op_income_missing",
                "dart_net_income_missing",
                "dart_assets_missing",
                "dart_equity_missing",
            ]
        ]
        .min(axis=1)
        .astype(int)
    )
    train["listing_market_missing"] = _missing_flag(train, "listing_market")
    train["valuation_model_missing"] = _missing_flag(train, "valuation_model")

    segment_features = sorted(
        c
        for c in train.columns
        if c.startswith("segment_") and c not in {"segment_bio", "segment_healthcare"}
    )
    for col in segment_features:
        train[col] = pd.to_numeric(train[col], errors="coerce").fillna(0).astype(int)

    numeric_model_features = (
        NUMERIC_FEATURES
        + ["log_per_stock", "log_peer_per", "is_loss_company"]
        + MISSING_FLAG_FEATURES
        + segment_features
    )

    medians = train[numeric_model_features].median(numeric_only=True)
    X_num = train[numeric_model_features].fillna(medians)

    categorical_levels: dict[str, list[str]] = {}
    encoded_parts = [X_num]
    for col in CATEGORICAL_FEATURES:
        clean = train[col].replace("", "UNKNOWN")
        categorical_levels[col] = sorted(clean.dropna().unique().tolist())
        dummies = pd.get_dummies(clean, prefix=col, dtype=int)
        encoded_parts.append(dummies)

    X = pd.concat(encoded_parts, axis=1)
    raw_features = X.columns.tolist()

    constant_features = [
        col for col in raw_features if X[col].nunique(dropna=False) <= 1
    ]
    features = [col for col in raw_features if col not in constant_features]
    X = X[features]
    medians = medians.reindex([c for c in features if c in medians.index])

    return X, features, medians.to_dict(), categorical_levels, constant_features


def main() -> None:
    train = load_training_data(INPUT_CSV)
    X, features, medians, categorical_levels, constant_features = engineer_features(train)
    y = train[TARGET].astype(float)

    print(f"Input: {INPUT_CSV}")
    print(f"Training rows: {len(train)}")
    print(f"Features: {len(features)}")
    print(f"Segment features: {len([f for f in features if f.startswith('segment_')])}")
    print(
        f"Categorical dummy features: "
        f"{len([f for f in features if f.startswith('listing_market_') or f.startswith('valuation_model_')])}"
    )
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
        print(f"  {row['feature']:32s}: {bar:34s} {row['importance_pct']:.1f}%")
    print()

    save_info = {
        "model": model_final,
        "features": features,
        "medians": medians,
        "categorical_features": CATEGORICAL_FEATURES,
        "categorical_levels": categorical_levels,
        "params": params,
        "n_train": len(train),
        "cv_rmse": round(float(rmse_scores.mean()), 4),
        "cv_mae": round(float(mae_scores.mean()), 4),
        "cv_r2": round(float(r2_scores.mean()), 4),
        "input_csv": INPUT_CSV,
        "version": "input_limited_general",
        "dropped_constant_features": constant_features,
        "excluded_feature_note": (
            "No final offering price, price_vs_mid, price band, offering amount, "
            "or listing-date-derived year/month features."
        ),
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
            "listing_market",
            "valuation_model",
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
                "n_categorical_dummy_features": len(
                    [
                        f
                        for f in features
                        if f.startswith("listing_market_")
                        or f.startswith("valuation_model_")
                    ]
                ),
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
