"""Stage 4 — 최종 비교군 압축 (하드 컷 + 유사도 밴드 + PER 대표 멀티플).

흐름(v3):
  1) PER = 시가총액 / 순이익(LTM). 적자/PER≤0/시총결측 제거
  2) 하드 컷 — 스타트업 비교 부적합 상장사 제외:
       시총 ≥ mcap_max(1000억) / PER < per_min(10) / PER ≥ per_max(100)
  3) 유사도(tier·score) 상위 max_peers개 후보 밴드
  4) (선택) 통계적 이상치 제거 outlier_method(iqr/trim) — 기본 none
  5) 대표 멀티플 = 중앙값(median) 기본

하드 컷은 도메인 규칙이라 min_peers로 완화하지 않음 — 부족하면 '피어 부족'으로 드러냄.
발행사 본인은 파이프라인 상단에서 배제됨.
"""
from __future__ import annotations

import pandas as pd

from config import Stage4Config

_TIER_RANK = {"높음": 2, "중간": 1, "낮음": 0}


def _outlier_drops(pers: pd.Series, cfg: Stage4Config) -> pd.Index:
    """제거할 인덱스. min_peers 바닥 보장(과제거 방지). 하드컷 이후 보조."""
    n = len(pers)
    if cfg.outlier_method == "none" or n <= cfg.min_peers:
        return pd.Index([])
    if cfg.outlier_method == "trim":
        k = cfg.trim_each_side
        if n - 2 * k < cfg.min_peers:
            k = max(0, (n - cfg.min_peers) // 2)
        if k == 0:
            return pd.Index([])
        s = pers.sort_values().index
        return pd.Index(list(s[:k]) + list(s[-k:]))
    q1, q3 = pers.quantile(0.25), pers.quantile(0.75)
    iqr = q3 - q1
    lo, hi = q1 - cfg.iqr_k * iqr, q3 + cfg.iqr_k * iqr
    out = pers[(pers < lo) | (pers > hi)].index
    if n - len(out) < cfg.min_peers:
        return pd.Index([])
    return out


def apply_stage4(stage3_peers: pd.DataFrame,
                 cfg: Stage4Config = Stage4Config()) -> pd.DataFrame:
    """Stage3 통과 → 하드컷 → 유사도 밴드 → (이상치) → 최종 비교군."""
    df = stage3_peers.copy()
    if df.empty:
        for c in ("per", "passed_stage4", "drop_reason"):
            df[c] = []
        return df

    df["per"] = df["market_cap"] / df[cfg.per_basis]

    if df["market_cap"].notna().sum() == 0:           # 시총 전부 결측 → 보류
        df["passed_stage4"] = False
        df["drop_reason"] = "시총없음(KRX 인가 필요)"
        return df.reset_index(drop=True)

    reason = pd.Series("", index=df.index)
    eligible = pd.Series(True, index=df.index)

    # 1. 시총 결측 / PER ≤ 0
    miss = df["market_cap"].isna()
    reason[miss], eligible[miss] = "시총없음", False
    if cfg.drop_nonpositive_per:
        bad = eligible & ~(df["per"] > 0)
        reason[bad], eligible[bad] = "PER≤0", False

    # 2. 하드 컷 (스타트업 비교 부적합)
    #    고유사도 면제: 사업이 매우 유사한 피어는 PER 상·하한 컷에서 제외(핵심 피어 보존).
    #    단 시총과대·PER≤0는 면제하지 않음(시총은 규모 비교, PER≤0는 멀티플 자체가 무의미).
    if cfg.high_sim_exempt:
        hs = (df["tier"].astype(str).str.strip() == cfg.high_sim_tier)
        if "sim_score" in df:
            hs = hs | (df["sim_score"] >= cfg.high_sim_min_score)
    else:
        hs = pd.Series(False, index=df.index)

    if cfg.mcap_max is not None:
        big = eligible & (df["market_cap"] >= cfg.mcap_max)
        reason[big], eligible[big] = f"시총과대(≥{cfg.mcap_max/1e8:.0f}억)", False
    if cfg.per_min is not None:
        low = eligible & (df["per"] < cfg.per_min) & ~hs
        reason[low], eligible[low] = f"PER과소(<{cfg.per_min:g})", False
    if cfg.per_max is not None:
        high = eligible & (df["per"] >= cfg.per_max) & ~hs
        reason[high], eligible[high] = f"PER과대(≥{cfg.per_max:g})", False

    # 3. 유사도 상위 max_peers개 후보 밴드
    df["_tr"] = df["tier"].map(_TIER_RANK).fillna(0)
    cand = df[eligible].sort_values(["_tr", "sim_score"], ascending=[False, False])
    beyond_idx = cand.index[cfg.max_peers:]
    for i in beyond_idx:
        reason[i], eligible[i] = f"max_peers({cfg.max_peers})초과", False
    band_idx = cand.index[:cfg.max_peers]

    # 4. (선택) 밴드 내 통계적 이상치
    if cfg.outlier_method != "none":
        drops = _outlier_drops(df.loc[band_idx, "per"], cfg)
        tag = "PER절사" if cfg.outlier_method == "trim" else "PER이상치"
        for i in drops:
            reason[i], eligible[i] = tag, False

    df["passed_stage4"] = eligible
    df["drop_reason"] = reason
    df = df.sort_values(["passed_stage4", "_tr", "sim_score"],
                        ascending=[False, False, False]).drop(columns="_tr")
    return df.reset_index(drop=True)


def final_comps(stage4_df: pd.DataFrame) -> pd.DataFrame:
    return stage4_df[stage4_df["passed_stage4"]].reset_index(drop=True)


def peer_per_summary(stage4_df: pd.DataFrame,
                     cfg: Stage4Config = Stage4Config()) -> dict:
    """최종 비교군 PER 통계 — 대표 멀티플(headline)은 중앙값 기본. 부족 시 thin=True."""
    fc = final_comps(stage4_df)
    if fc.empty:
        return {}
    pers = fc["per"]
    median = round(float(pers.median()), 2)
    mean = round(float(pers.mean()), 2)
    # 유사도 가중평균 — sim_score 가중(없으면 균등). 유사한 피어 PER에 더 큰 비중.
    if "sim_score" in fc and fc["sim_score"].notna().any():
        w = fc["sim_score"].clip(lower=1.0)
        sim_weighted = round(float((pers * w).sum() / w.sum()), 2)
    else:
        sim_weighted = mean
    headline_per = {"median": median, "mean": mean,
                    "sim_weighted": sim_weighted}.get(cfg.headline, median)
    return {
        "n": int(len(fc)),
        "thin": bool(len(fc) < cfg.min_peers),
        "headline": cfg.headline,
        "headline_per": headline_per,
        "per_median": median,
        "per_mean": mean,
        "per_sim_weighted": sim_weighted,
        "per_min": round(float(pers.min()), 2),
        "per_max": round(float(pers.max()), 2),
    }