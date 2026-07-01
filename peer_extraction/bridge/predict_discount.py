"""평가 → 공모할인율 브리지 — 피어 추출/평가 결과를 discount_rate 모델(input_limited)에 연결.

input_limited 모델만 사전평가에 사용 가능(나머지 v1/general/v2는 price_vs_mid·price_band 등
'공모가를 이미 알아야 하는' 순환 피처에 의존 → 비상장 사전평가 불가).

흐름:
  피어추출 → 대표PER·평가시총 →  [이 브리지: 피처벡터 재구성 → model.predict 할인율]
  → 예상 공모시총 = 평가시총 × (1 − 할인율)

추론 시 피처 벡터를 학습(engineer_features)과 동일하게 재구성한다:
  - 수치 결측은 저장된 medians로 대체 + *_missing 플래그 1
  - log_per_stock=log1p(per_stock_value), log_peer_per=log1p(peer_per), is_loss_company
  - listing_market/valuation_model 범주 더미는 저장된 features 컬럼에 정렬(없으면 0)
  - 최종 features 순서로 reindex
저장된 모델 dict(joblib): {model, features, medians, categorical_features, categorical_levels, ...}
"""
from __future__ import annotations

import numpy as np
import pandas as pd

NUMERIC = ["per_stock_value", "peer_per", "net_income_bn",
           "dart_revenue_bn", "dart_op_income_bn", "dart_net_income_bn",
           "dart_assets_bn", "dart_equity_bn", "dart_debt_ratio", "dart_op_margin",
           "is_tech_special", "is_growth_special"]
DART_KEYS = ["dart_revenue_bn", "dart_op_income_bn", "dart_net_income_bn",
             "dart_assets_bn", "dart_equity_bn"]


def build_feature_row(model_info: dict, *, peer_per=None, per_stock_value=None,
                      net_income_bn=None, dart: dict | None = None,
                      is_tech_special=0, is_growth_special=0,
                      listing_market="코스닥", valuation_model="PER",
                      segments: dict | None = None) -> pd.DataFrame:
    """단일 발행사 입력 → input_limited 모델의 features에 정렬된 1행 X."""
    features = model_info["features"]
    medians = model_info.get("medians", {})
    dart = dart or {}
    seg = segments or {}

    raw = {k: None for k in NUMERIC}
    raw["per_stock_value"] = per_stock_value
    raw["peer_per"] = peer_per
    raw["net_income_bn"] = net_income_bn
    raw["is_tech_special"] = is_tech_special
    raw["is_growth_special"] = is_growth_special
    for k in ["dart_revenue_bn", "dart_op_income_bn", "dart_net_income_bn",
              "dart_assets_bn", "dart_equity_bn", "dart_debt_ratio", "dart_op_margin"]:
        raw[k] = dart.get(k)

    row: dict[str, float] = {}
    # 1) 결측 플래그 (대체 전에 계산)
    row["per_stock_value_missing"] = int(per_stock_value is None)
    row["peer_per_missing"] = int(peer_per is None)
    row["net_income_missing"] = int(net_income_bn is None)
    row["dart_revenue_missing"] = int(dart.get("dart_revenue_bn") is None)
    row["dart_op_income_missing"] = int(dart.get("dart_op_income_bn") is None)
    row["dart_net_income_missing"] = int(dart.get("dart_net_income_bn") is None)
    row["dart_assets_missing"] = int(dart.get("dart_assets_bn") is None)
    row["dart_equity_missing"] = int(dart.get("dart_equity_bn") is None)
    row["dart_financial_missing"] = int(min(
        row["dart_revenue_missing"], row["dart_op_income_missing"],
        row["dart_net_income_missing"], row["dart_assets_missing"],
        row["dart_equity_missing"]))
    row["listing_market_missing"] = int(not listing_market)
    row["valuation_model_missing"] = int(not valuation_model)

    # 2) 수치 피처 (결측 → 저장 median)
    for k in NUMERIC:
        v = raw[k]
        row[k] = float(v) if v is not None else float(medians.get(k, 0.0) or 0.0)

    # 3) 파생
    psv = per_stock_value if per_stock_value is not None else medians.get("per_stock_value", 0.0)
    pp = peer_per if peer_per is not None else medians.get("peer_per", 0.0)
    row["log_per_stock"] = float(np.log1p(max(psv or 0.0, 0.0)))
    row["log_peer_per"] = float(np.log1p(max(pp or 0.0, 0.0)))
    ni = net_income_bn if net_income_bn is not None else 0.0
    dni = dart.get("dart_net_income_bn")
    row["is_loss_company"] = int((ni < 0) or (dni is not None and dni < 0))

    # 4) 섹터 플래그 (features에 있는 segment_*만)
    for f in features:
        if f.startswith("segment_"):
            row[f] = int(seg.get(f, 0))

    # 5) 범주 더미 (features 컬럼명에 맞춰 1/0)
    for f in features:
        if f.startswith("listing_market_"):
            row[f] = int(f == f"listing_market_{listing_market or 'UNKNOWN'}")
        elif f.startswith("valuation_model_"):
            row[f] = int(f == f"valuation_model_{valuation_model or 'UNKNOWN'}")

    # 6) features 순서로 정렬 (누락은 median 또는 0)
    out = {}
    for f in features:
        if f in row:
            out[f] = row[f]
        elif f in medians:
            out[f] = float(medians[f])
        else:
            out[f] = 0.0
    return pd.DataFrame([out])[features]


def predict_discount(model_info: dict, **kwargs) -> float:
    """할인율(discount_rate_mid_pct) 예측값(%) 반환."""
    X = build_feature_row(model_info, **kwargs)
    return float(model_info["model"].predict(X)[0])


def apply_discount_to_valuation(equity_value: float, discount_pct: float) -> float:
    """평가시총 × (1 − 할인율) = 예상 공모시총. 단위는 equity_value와 동일."""
    return equity_value * (1.0 - discount_pct / 100.0)
