"""PER 기반 평가가액 산출 — 신고서 표준 방법(미래 추정순이익 현가화 × 피어 PER).

적자/초기기업은 현재 PER 평가 불가 → 미래 추정순이익을 현재가치로 할인한
'적용 순이익'에 피어 대표 PER을 곱해 평가 시가총액(equity value) 산출.
(이지트로닉스 신고서 방식: 적용순이익 C = 추정순이익 A / (1+할인율)^n, 평가액 = C × 적용PER)

대표 PER 산출 방식을 여러 가지 제공 — 유사도 반영 여부가 평가액을 크게 좌우.
"""
from __future__ import annotations

from statistics import median, mean


def present_value(future_ni: float, discount_rate: float, years: float) -> float:
    """미래 추정순이익 → 현재가치(적용 순이익). C = A / (1+r)^n."""
    return future_ni / ((1.0 + discount_rate) ** years)


def representative_per(peers: list[dict], method: str = "median") -> float:
    """피어 PER들의 대표값. peers: [{name, per, similarity(0~100), tier}].

    method:
      median        - 전체 중앙값 (현행 기본; 이상치에 강건)
      mean          - 전체 평균
      sim_weighted  - 유사도 가중평균 (유사할수록 PER 비중↑)
      high_only     - '높음' tier만 평균 (가장 유사한 피어만)
      high_median   - '높음' tier만 중앙값
    """
    pers = [p["per"] for p in peers if p.get("per") is not None]
    if not pers:
        return 0.0
    if method == "median":
        return median(pers)
    if method == "mean":
        return mean(pers)
    if method == "sim_weighted":
        num = sum(p["per"] * p.get("similarity", 50) for p in peers if p.get("per") is not None)
        den = sum(p.get("similarity", 50) for p in peers if p.get("per") is not None)
        return num / den if den else median(pers)
    if method in ("high_only", "high_median"):
        highs = [p["per"] for p in peers
                 if p.get("per") is not None and str(p.get("tier", "")).strip() == "높음"]
        if not highs:
            return median(pers)
        return mean(highs) if method == "high_only" else median(highs)
    return median(pers)


def equity_value(applied_ni: float, per: float) -> float:
    """평가 시가총액 = 적용 순이익 × 대표 PER."""
    return applied_ni * per


def valuation_scenarios(peers: list[dict], future_ni: float,
                        discount_rate: float, years: float) -> dict:
    """여러 대표 PER 방식별 평가액 시나리오. 금액 단위는 future_ni와 동일."""
    applied = present_value(future_ni, discount_rate, years)
    out = {"applied_ni": applied, "future_ni": future_ni,
           "discount_rate": discount_rate, "years": years, "scenarios": {}}
    for m in ("median", "sim_weighted", "high_only", "mean"):
        per = representative_per(peers, m)
        out["scenarios"][m] = {"per": per, "equity_value": equity_value(applied, per)}
    return out


# -------------------- 주당 평가가액 → 희망 공모가액 밴드 (신고서 형식) --------------------
def per_share_value(equity_value: float, shares: float) -> float:
    """주당 평가가액 = 평가시총 / 공모후 발행주식수. 단위: equity_value와 동일/주."""
    return equity_value / shares if shares else 0.0


def post_offering_shares(capital: float, par_value: float = 500.0,
                         dilution: float = 0.25) -> float:
    """공모후 발행주식수 추정.

    공모전 = 자본금 / 액면가.  공모후 = 공모전 / (1 − dilution).
    dilution = 신주(공모주식)가 공모후 총주식에서 차지하는 비율(기본 25%).
    dilution=0이면 공모전 주식수 그대로(희석 무시).
    """
    pre = capital / par_value
    return pre / (1.0 - dilution) if dilution < 1.0 else pre


def offering_price_band(per_share: float, discount_mid_pct: float,
                        halfwidth_pct: float = 5.0, round_to: int = 100) -> dict:
    """주당 평가가액 + 할인율(점추정) → 희망 공모가액 밴드(신고서 형식).

    신고서 관행: 할인율 낮음(상단가) ~ 할인율 높음(하단가).
      공모가 상단 = 주당평가 × (1 − 낮은할인율)
      공모가 하단 = 주당평가 × (1 − 높은할인율)
    discount_mid를 중심으로 ±halfwidth 밴드. round_to원 단위 반올림.
    """
    d_low = max(discount_mid_pct - halfwidth_pct, 0.0)   # 낮은 할인 → 높은 가격
    d_high = discount_mid_pct + halfwidth_pct            # 높은 할인 → 낮은 가격
    price_high = per_share * (1.0 - d_low / 100.0)
    price_low = per_share * (1.0 - d_high / 100.0)
    if round_to:
        price_high = round(price_high / round_to) * round_to
        price_low = round(price_low / round_to) * round_to
    return {
        "discount_low": round(d_low, 2), "discount_high": round(d_high, 2),
        "price_low": price_low, "price_high": price_high,
        "price_mid": round((price_low + price_high) / 2),
    }