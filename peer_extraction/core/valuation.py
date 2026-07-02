"""PER 기반 평가가액 산출 — 신고서 표준 방법(미래 추정순이익 현가화 × 피어 PER).
(실제 라이브 core/valuation.py 그대로 복사 — 검증용)
"""
from __future__ import annotations
from statistics import median, mean


def present_value(future_ni: float, discount_rate: float, years: float) -> float:
    return future_ni / ((1.0 + discount_rate) ** years)


def representative_per(peers: list[dict], method: str = "median") -> float:
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
    return applied_ni * per


def valuation_scenarios(peers: list[dict], future_ni: float,
                        discount_rate: float, years: float) -> dict:
    applied = present_value(future_ni, discount_rate, years)
    out = {"applied_ni": applied, "future_ni": future_ni,
           "discount_rate": discount_rate, "years": years, "scenarios": {}}
    for m in ("median", "sim_weighted", "high_only", "mean"):
        per = representative_per(peers, m)
        out["scenarios"][m] = {"per": per, "equity_value": equity_value(applied, per)}
    return out


# ══════════════════════════════════════════════════════════════
# 다년도 추정치 기반 현재가치 밴드 (2026-07 설계)
#
# 배경: 스타트업은 상장사와 달리 "지금 안정적인 실적"이 가치의 근거가 아니라,
# "매출/이익이 성장하는 궤적"이 가치의 근거다. "3년 뒤 목표를 달성하면 지금
# 얼마짜리인가"를 보여주는 게 핵심이라, 단일 성장률로 미래를 외삽하는 대신
# 플랫폼의 "추정 재무제표 입력" 화면(연도별 매출/영업이익/당기순이익)을 그대로
# 받아 연도별로 각각 현재가치화해서 밴드로 제시한다.
#
# 할인율 스케줄: 다음해 15% / 2년뒤 20% / 3년뒤 30% (2026-07 확정).
# 연차가 멀수록 요율이 커지는 건 예측 신뢰도가 떨어진다는 뜻 — 초기 연차는
# 예측이 비교적 믿을만하지만 멀어질수록 불확실성이 커진다는 철학을 반영.
# 복리 방식은 "해당 연차 요율로 단순 거듭제곱"(연차마다 다른 요율을 순차
# 복리하는 방식이 아님) — 2026-07 확정. 예: 3년 뒤 값은 (1+30%)^3로 나눔.
# ══════════════════════════════════════════════════════════════

DISCOUNT_RATE_SCHEDULE = {1: 0.15, 2: 0.20, 3: 0.30}
_MAX_SCHEDULED_YEAR = max(DISCOUNT_RATE_SCHEDULE)


def tiered_discount_rate(years_out: int) -> float:
    """years_out(정수, 당해년도=0)에 해당하는 할인율.
    스케줄 밖(4년 이상)은 가장 먼 연차 요율을 그대로 유지한다 — 더 먼 미래일수록
    불확실성이 줄어들 이유가 없으므로 보수적으로 마지막 요율을 캡으로 쓴다."""
    if years_out <= 0:
        return 0.0
    return DISCOUNT_RATE_SCHEDULE.get(years_out, DISCOUNT_RATE_SCHEDULE[_MAX_SCHEDULED_YEAR])


def present_value_tiered(future_value: float, years_out: int) -> float:
    """미래값을 연차별 요율로 단순 거듭제곱 할인 (present_value()의 다년도 스케줄 버전).
    당해년도(years_out<=0)는 할인 없이 그대로 반환."""
    if years_out <= 0:
        return future_value
    rate = tiered_discount_rate(years_out)
    return future_value / ((1.0 + rate) ** years_out)


def value_band_from_projections(projections: dict, representative_multiple: float,
                                base_year: int) -> dict:
    """연도별 추정치(매출 또는 순이익) × 대표 배수(PER 또는 PSR) → 연도별 현재가치 밴드.

    Args:
        projections: {연도(int): 추정치(억원)} — 예: {2027: 32.5, 2028: 54.0}.
                     "추정 재무제표 입력" 화면이나 core/ir_parser.py에서 온 값.
        representative_multiple: 대표 PER(순이익 기준) 또는 PSR(매출 기준).
        base_year: 오늘 기준 연도 — 할인연수(years_out = 연도 - base_year) 계산 기준.

    Returns:
        {연도: {"years_out", "raw_value_억원"(할인 전 평가액),
                "discount_rate", "present_value_억원"(할인 후, 오늘 기준 가치)}}
        연도 오름차순 정렬은 호출측 책임(dict는 삽입 순서 유지되므로 정렬해서 넣으면 됨).
    """
    band = {}
    for year in sorted(projections.keys()):
        proj_value = projections[year]
        years_out = year - base_year
        raw_value = proj_value * representative_multiple
        rate = tiered_discount_rate(years_out)
        pv = present_value_tiered(raw_value, years_out)
        band[year] = {
            "years_out": years_out,
            "raw_value_억원": raw_value,
            "discount_rate": rate,
            "present_value_억원": pv,
        }
    return band