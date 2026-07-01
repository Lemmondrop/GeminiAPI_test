"""
V-Wheel ML API — Value-Band 할인율 예측 서버
모델 v1 (79건 학습, 완전커버 피처)

엔드포인트:
  GET  /health              헬스체크
  GET  /model/info          모델 메타데이터
  POST /predict/discount    공모할인율 예측
  POST /predict/batch       배치 예측
"""
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field, field_validator
from typing import Optional
import numpy as np
import joblib
import os
from datetime import datetime

app = FastAPI(
    title="V-Wheel ML API",
    description="IPO 공모할인율 예측 모델 (Value-Band Engine v1)",
    version="1.0.0",
)

# ── 모델 로드 ──────────────────────────────────────────────
MODEL_PATH = os.environ.get("MODEL_PATH", "model_v1.pkl")
_bundle = None

def get_bundle():
    global _bundle
    if _bundle is None:
        _bundle = joblib.load(MODEL_PATH)
    return _bundle


# ══════════════════════════════════════════════════════════
# 스키마
# ══════════════════════════════════════════════════════════

class DiscountRequest(BaseModel):
    company_name:       str             = Field(...,  description="회사명")
    listing_date:       str             = Field(...,  description="상장예정일 YYYY-MM-DD")
    price_band_low:     float           = Field(...,  gt=0, description="희망공모가 하단 (원)")
    price_band_high:    float           = Field(...,  gt=0, description="희망공모가 상단 (원)")
    final_offering_price: Optional[float] = Field(None, gt=0, description="확정공모가 (원). 미확정 시 밴드 중간값 사용")
    per_stock_value:    Optional[float] = Field(None, gt=0, description="주당 평가가액 (원)")
    offering_amount:    Optional[float] = Field(None, gt=0, description="공모총액 (원)")
    track_type:         Optional[str]   = Field("general", description="general | konex_transfer")

    @field_validator('listing_date')
    @classmethod
    def validate_date(cls, v):
        try:
            datetime.strptime(v, '%Y-%m-%d')
        except ValueError:
            raise ValueError("listing_date 형식: YYYY-MM-DD")
        return v

    @field_validator('price_band_high')
    @classmethod
    def validate_band(cls, v, info):
        if 'price_band_low' in info.data and v <= info.data['price_band_low']:
            raise ValueError("price_band_high > price_band_low 이어야 합니다")
        return v


class DiscountResponse(BaseModel):
    company_name:           str
    discount_rate_mid_pct:  float   = Field(..., description="예측 할인율 mid (%)")
    discount_rate_low_pct:  float   = Field(..., description="예측 할인율 low (mid - 1σ)")
    discount_rate_high_pct: float   = Field(..., description="예측 할인율 high (mid + 1σ)")
    implied_offering_price: float   = Field(..., description="역산 추정 공모가 (원)")
    price_vs_mid_used:      float   = Field(..., description="사용된 price_vs_mid 값")
    model_version:          str
    n_train:                int


class BatchRequest(BaseModel):
    items: list[DiscountRequest]


class ModelInfo(BaseModel):
    version:    str
    n_train:    int
    features:   list[str]
    cv_rmse:    float
    cv_mae:     float
    cv_r2:      float
    loaded_at:  str


# ══════════════════════════════════════════════════════════
# 예측 로직
# ══════════════════════════════════════════════════════════

CV_RMSE_SIGMA = 7.72   # 모델 CV RMSE → 불확실성 범위로 사용

def _build_features(req: DiscountRequest, bundle: dict) -> np.ndarray:
    dt       = datetime.strptime(req.listing_date, '%Y-%m-%d')
    band_mid = (req.price_band_low + req.price_band_high) / 2
    final    = req.final_offering_price or band_mid

    price_vs_mid    = (final / band_mid) - 1
    band_width_pct  = (req.price_band_high - req.price_band_low) / req.price_band_low * 100
    medians         = bundle['medians']

    per_stock_val   = req.per_stock_value  or medians.get('log_per_stock', 0)
    log_per_stock   = np.log1p(per_stock_val) if req.per_stock_value else medians.get('log_per_stock', 0)
    offering_amt    = req.offering_amount  or 0
    log_offering    = np.log1p(offering_amt) if req.offering_amount else medians.get('log_offering', 0)
    is_konex        = 1 if req.track_type == 'konex_transfer' else 0

    feat_order = bundle['features']
    feat_map = {
        'price_vs_mid':    price_vs_mid,
        'band_width_pct':  band_width_pct,
        'year':            dt.year,
        'month':           dt.month,
        'log_per_stock':   log_per_stock,
        'log_offering':    log_offering,
        'is_konex':        is_konex,
        'price_band_low':  req.price_band_low,
        'price_band_high': req.price_band_high,
    }
    return np.array([feat_map[f] for f in feat_order]).reshape(1, -1), price_vs_mid


def _predict_one(req: DiscountRequest, bundle: dict) -> DiscountResponse:
    X, pvs = _build_features(req, bundle)
    mid    = float(bundle['model'].predict(X)[0])
    mid    = round(max(5.0, min(70.0, mid)), 2)   # 클리핑

    per_stock = req.per_stock_value or (
        (req.price_band_low + req.price_band_high) / 2 / (1 - mid / 100)
    )
    implied_price = round(per_stock * (1 - mid / 100), 0)

    return DiscountResponse(
        company_name=req.company_name,
        discount_rate_mid_pct=mid,
        discount_rate_low_pct=round(max(0, mid - CV_RMSE_SIGMA), 2),
        discount_rate_high_pct=round(min(70, mid + CV_RMSE_SIGMA), 2),
        implied_offering_price=implied_price,
        price_vs_mid_used=round(pvs, 4),
        model_version=bundle.get('version', 'v1'),
        n_train=bundle.get('n_train', 0),
    )


# ══════════════════════════════════════════════════════════
# 라우터
# ══════════════════════════════════════════════════════════

@app.get("/health")
def health():
    return {"status": "ok", "timestamp": datetime.utcnow().isoformat()}


@app.get("/model/info", response_model=ModelInfo)
def model_info():
    b = get_bundle()
    return ModelInfo(
        version=b.get('version', 'v1'),
        n_train=b.get('n_train', 0),
        features=b.get('features', []),
        cv_rmse=b.get('cv_rmse', 0),
        cv_mae=b.get('cv_mae', 0),
        cv_r2=b.get('cv_r2', 0),
        loaded_at=datetime.utcnow().isoformat(),
    )


@app.post("/predict/discount", response_model=DiscountResponse)
def predict_discount(req: DiscountRequest):
    try:
        bundle = get_bundle()
        return _predict_one(req, bundle)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/predict/batch", response_model=list[DiscountResponse])
def predict_batch(req: BatchRequest):
    if len(req.items) > 100:
        raise HTTPException(status_code=400, detail="배치 최대 100건")
    bundle = get_bundle()
    return [_predict_one(item, bundle) for item in req.items]


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
