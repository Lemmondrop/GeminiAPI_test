"""
Phase 4: EDA + XGBoost 모델 학습 및 저장

현재 79건 기준 베이스라인 모델 수립.
데이터 250건 달성 시 재학습 예정.

출력:
  - model_v1.pkl          : 학습된 XGBoost 모델
  - feature_importance.csv: 피처 중요도
  - eda_summary.csv       : EDA 요약
  - predictions_cv.csv    : CV 예측값 (잔차 분석용)
"""
import os, warnings, joblib
import pandas as pd
import numpy as np
from sklearn.model_selection import cross_val_predict, KFold, cross_val_score
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
import xgboost as xgb
warnings.filterwarnings('ignore')

# ── 1. 데이터 로드 및 필터링 ───────────────────────────────
INPUT_CSV = os.environ.get("TRAINING_CSV", "training_features.csv")
df = pd.read_csv(INPUT_CSV, encoding='utf-8-sig')

train = df[
    (df['extraction_status'] != 'download_failed') &
    (df['company_name'] != '대진첨단소재') &   # 한라오기재
    df['price_band_low'].notna() &
    df['final_offering_price'].notna()
].copy()

print(f"학습셋: {len(train)}건")

# ── 2. 피처 엔지니어링 ────────────────────────────────────
train['year']  = pd.to_datetime(train['listing_date']).dt.year
train['month'] = pd.to_datetime(train['listing_date']).dt.month

train['band_width_pct'] = (
    (train['price_band_high'] - train['price_band_low'])
    / train['price_band_low'] * 100
).round(4)

train['log_per_stock'] = np.log1p(
    train['per_stock_value'].fillna(train['per_stock_value'].median())
)
train['log_offering']  = np.log1p(
    train['offering_amount'].fillna(train['offering_amount'].median())
)
train['is_konex']  = (train['track_type']    == 'konex_transfer').astype(int)
train['is_kosdaq'] = (train['listing_market'] == '코스닥').astype(int)

# ── 3. 피처 정의 ──────────────────────────────────────────
FEATURES = [
    'price_vs_mid',        # 확정가의 밴드 내 위치 (핵심)
    'band_width_pct',      # 밴드 폭 (시장 불확실성)
    'year',                # 상장 연도 (시장 환경)
    'month',               # 상장 월 (계절성)
    'log_per_stock',       # 주당 평가가액 로그 (기업 규모)
    'log_offering',        # 공모 총액 로그 (딜 규모)
    'is_konex',            # KONEX 이전 상장 여부
    'price_band_low',      # 희망밴드 하단
    'price_band_high',     # 희망밴드 상단
]
TARGET = 'discount_rate_mid_pct'

X = train[FEATURES].copy()
y = train[TARGET]

# 결측치: 중앙값 대체 (학습/추론 동일하게)
medians = X.median()
X = X.fillna(medians)

print(f"피처: {len(FEATURES)}개")
print(f"결측치 보완: {(train[FEATURES].isna().sum() > 0).sum()}개 피처")
print()

# ── 4. 최적 파라미터 (탐색 결과 반영) ────────────────────
params = dict(
    n_estimators=100,
    max_depth=2,
    learning_rate=0.05,
    subsample=0.8,
    colsample_bytree=0.8,
    min_child_weight=10,
    reg_alpha=0.1,
    reg_lambda=1.0,
    random_state=42,
    objective='reg:squarederror',
)

# ── 5. CV 평가 ────────────────────────────────────────────
kf = KFold(n_splits=5, shuffle=True, random_state=42)
model_cv = xgb.XGBRegressor(**params)

rmse_scores = np.sqrt(-cross_val_score(model_cv, X, y, cv=kf, scoring='neg_mean_squared_error'))
mae_scores  = -cross_val_score(model_cv, X, y, cv=kf, scoring='neg_mean_absolute_error')
r2_scores   =  cross_val_score(model_cv, X, y, cv=kf, scoring='r2')
y_pred_cv   = cross_val_predict(model_cv, X, y, cv=kf)

print("=" * 50)
print("5-Fold CV 결과")
print("=" * 50)
print(f"  RMSE : {rmse_scores.mean():.3f} ± {rmse_scores.std():.3f} %p")
print(f"  MAE  : {mae_scores.mean():.3f} ± {mae_scores.std():.3f} %p")
print(f"  R²   : {r2_scores.mean():.3f} ± {r2_scores.std():.3f}")
print(f"  베이스라인 RMSE: {y.std():.3f} %p (평균예측)")
print(f"  개선율: {(1 - rmse_scores.mean()/y.std())*100:.1f}%")
print()

# ── 6. 전체 데이터로 최종 모델 학습 ─────────────────────
model_final = xgb.XGBRegressor(**params)
model_final.fit(X, y)

# 피처 중요도
importance = pd.DataFrame({
    'feature': FEATURES,
    'importance': model_final.feature_importances_,
    'importance_pct': model_final.feature_importances_ / model_final.feature_importances_.sum() * 100
}).sort_values('importance', ascending=False).round(4)

print("피처 중요도:")
for _, row in importance.iterrows():
    bar = '█' * int(row['importance_pct'] / 3) + '░' * (34 - int(row['importance_pct'] / 3))
    print(f"  {row['feature']:20s}: {bar} {row['importance_pct']:.1f}%")
print()

# ── 7. 저장 ──────────────────────────────────────────────
save_info = {
    'model':    model_final,
    'features': FEATURES,
    'medians':  medians.to_dict(),
    'params':   params,
    'n_train':  len(train),
    'cv_rmse':  round(rmse_scores.mean(), 4),
    'cv_mae':   round(mae_scores.mean(), 4),
    'cv_r2':    round(r2_scores.mean(), 4),
}
joblib.dump(save_info, 'model_v1.pkl')
print("✅ 모델 저장: model_v1.pkl")

importance.to_csv('feature_importance.csv', index=False, encoding='utf-8-sig')
print("✅ 피처 중요도: feature_importance.csv")

# CV 예측 저장 (잔차 분석)
cv_df = train[['company_name','listing_date','discount_rate_mid_pct']].copy()
cv_df['y_pred'] = y_pred_cv.round(4)
cv_df['residual'] = (cv_df['discount_rate_mid_pct'] - cv_df['y_pred']).round(4)
cv_df['abs_error'] = cv_df['residual'].abs().round(4)
cv_df = cv_df.sort_values('abs_error', ascending=False)
cv_df.to_csv('predictions_cv.csv', index=False, encoding='utf-8-sig')
print("✅ CV 예측값: predictions_cv.csv")
print()

print("=" * 50)
print("상위 잔차 10건 (오차 큰 순)")
print("=" * 50)
print(cv_df.head(10)[['company_name','listing_date',
                        'discount_rate_mid_pct','y_pred','residual']].to_string(index=False))
print()
print("📌 Note: 79건 기준 베이스라인 모델.")
print("   250건 달성 및 DART 재무 API 피처 추가 후 v2 재학습 예정.")