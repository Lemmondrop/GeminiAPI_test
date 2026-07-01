"""Stage 3 — 사업 유사성 (하이브리드: 임베딩 선별 + LLM 근거).

흐름:
  1) Stage 2 통과를 embed_score 내림차순 정렬 → 상위 K개 (결정론 백본)
  2) 각 후보를 Gemini LLM이 사업모델 유사성 판정 → {tier, score, rationale} (캐싱, temp=0)
  3) tier(높음/중간)로 최종 선별

후보 사업텍스트:
  1순위 = DART 보강 텍스트 (stage3_business_text_final.json, 평균 ~450자)
  2순위 = BUSINESS_TEXT_COLS(주요제품 + 업종) 원본 컬럼 — 보강 캐시 미존재/저품질 시 fallback

발행사 텍스트는 호출측이 제공.
재현성은 LlmJudge 캐시가 보장 — 같은 쌍은 항상 같은 판정.
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from config import BUSINESS_TEXT_COLS, NAME_COL, STOCK_COL, Stage3Config

_TIER_RANK = {"높음": 2, "중간": 1, "낮음": 0}

# ── DART 보강 텍스트 캐시 ────────────────────────────────────
# dart_biz_merge_final.py 산출물. peer_extraction/data/cache/ 에 위치
# (이 파일은 peer_extraction/stages/stage3_business.py 이므로 두 단계 상위 참조).
_BIZ_CACHE_PATH = Path(__file__).parent.parent / "data" / "cache" / "stage3_business_text_final.json"
_biz_cache: dict | None = None
_MIN_ENRICHED_LEN = 10   # 이보다 짧으면 보강 실패로 간주 → fallback


def _load_biz_cache() -> dict:
    """stage3_business_text_final.json을 1회 로드 후 메모리에 캐싱."""
    global _biz_cache
    if _biz_cache is None:
        if _BIZ_CACHE_PATH.exists():
            with open(_BIZ_CACHE_PATH, encoding="utf-8") as f:
                _biz_cache = json.load(f)
        else:
            _biz_cache = {}
    return _biz_cache


def business_text(row, cols=BUSINESS_TEXT_COLS) -> str:
    """후보 1행 → LLM 입력 사업설명. 결측 컬럼은 건너뜀.

    DART 보강 텍스트가 있으면 우선 사용하고, 없거나 너무 짧으면
    기존 컬럼 결합 방식(주요제품 + 업종)으로 폴백한다.
    """
    cache = _load_biz_cache()
    stock_code = str(row.get(STOCK_COL, "")).strip().zfill(6)
    entry = cache.get(stock_code)
    if entry:
        enriched = str(entry.get("business_text", "")).strip()
        if len(enriched) >= _MIN_ENRICHED_LEN:
            return enriched

    # ── fallback: 기존 컬럼 결합 방식 ──────────────────────
    parts = []
    for c in cols:
        v = row.get(c)
        if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip():
            parts.append(str(v).strip())
    return " | ".join(parts) if parts else "(사업설명 없음)"


def apply_stage3(survivors: pd.DataFrame, issuer_text: str, judge,
                 cfg: Stage3Config = Stage3Config(),
                 progress: bool = True) -> pd.DataFrame:
    """Stage2 통과 → 상위K LLM 판정 → tier 선별. 판정·사유 컬럼이 붙은 DataFrame 반환."""
    if survivors.empty:
        return survivors.assign(tier=[], sim_score=[], rationale=[], passed_stage3=[])

    ranked = survivors.sort_values("embed_score", ascending=False).head(cfg.top_k_judge).copy()

    cands = [(str(row[STOCK_COL]), business_text(row)) for _, row in ranked.iterrows()]
    judged = judge.judge_many(issuer_text, cands, progress=progress)   # {종목코드: {tier,score,rationale}}

    codes = ranked[STOCK_COL].astype(str)
    ranked["tier"] = [judged[c]["tier"] for c in codes]
    ranked["sim_score"] = [judged[c]["score"] for c in codes]
    ranked["rationale"] = [judged[c]["rationale"] for c in codes]
    ranked["passed_stage3"] = ranked["tier"].isin(cfg.keep_tiers)

    ranked["_tier_rank"] = ranked["tier"].map(_TIER_RANK).fillna(0)
    ranked = ranked.sort_values(["passed_stage3", "_tier_rank", "sim_score"],
                                ascending=[False, False, False]).drop(columns="_tier_rank")
    return ranked.reset_index(drop=True)


def final_peers(stage3_df: pd.DataFrame) -> pd.DataFrame:
    """최종 피어(통과)만, 유사도순."""
    return stage3_df[stage3_df["passed_stage3"]].reset_index(drop=True)