"""Stage 1 — 업종 유사성 모집단(Recall Pool) 생성기.

설계 원칙: 1단계는 precision이 아니라 RECALL 문제. (모집단 100~250개 → 2~4단계에서 정제)
3개 레인을 union 하고, 각 후보에 '어느 레인이 잡았는지' provenance를 태깅한다.

  Lane A  코드 확장   : KSIC 소분류 exact + 중분류 fallback   (결정론, 재현가능, 저비용)
  Lane B  임베딩 검색 : 캐시된 코퍼스 ↔ 쿼리 코사인 top-K     (ML, 결정론적 검색 / 생성형 아님)
  Lane C  동시출현 prior : 과거 신고서 (발행사 소분류→모집단 소분류)  (데이터 앵커, 선택)

생성형 LLM은 이 단계에 없음. (재랭킹/근거 서술은 3단계 또는 옵션)

사용
  vc = VectorCache(GeminiEmbedder()); vc.build_or_load()        # 코퍼스 1회 캐시
  q  = StartupQuery(seed_code="C00303", business_text="...")
  pool = build_population(q, vc, cooccur_prior=PRIOR)
  save_pool(pool, "data/stage1_pool.csv")
"""
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping

import numpy as np
import pandas as pd

from config import CODE_COL, MARKET_COL, NAME_COL, Stage1Config
from core.ksic import normalize_seed, parse_ksic
from core.vector_cache import VectorCache
from config import STOCK_COL, BUSINESS_TEXT_COLS

PROVENANCE_COLS = ["lane_code_exact", "lane_code_mid", "lane_cooccur", "lane_embed"]
OUTPUT_COLS = [NAME_COL, STOCK_COL, MARKET_COL, CODE_COL, *BUSINESS_TEXT_COLS,
               "embed_score", *PROVENANCE_COLS, "in_pool"]


@dataclass
class StartupQuery:
    seed_code: str       # 자가입력 KSIC (검증 권장 / 임베딩은 코드 오류에 강건)
    business_text: str   # IR·사업계획에서 추출한 사업설명 (임베딩 쿼리)


def build_population(
    startup: StartupQuery,
    vcache: VectorCache,
    cooccur_prior: Mapping[str, Mapping[str, float]] | None = None,
    cfg: Stage1Config = Stage1Config(),
) -> pd.DataFrame:
    """vcache.build_or_load()가 선행되어 vcache.vectors가 채워져 있다고 가정."""
    if vcache.vectors is None:
        raise RuntimeError("VectorCache.build_or_load()를 먼저 호출하세요.")

    df = vcache.df.copy().reset_index(drop=True)
    maj, mid, sub = zip(*df[CODE_COL].map(parse_ksic))
    df["_maj"], df["_mid"], df["_sub"] = maj, mid, sub

    # ----- Lane B: 임베딩 (캐시 코퍼스 행렬 · 쿼리 1건) -----
    # 정규화 벡터라 내적 = 코사인. 전 행 스코어가 필요하므로 직접 matmul.
    # (FAISS 인덱스는 유니버스가 수십만으로 커질 때를 위한 가속기로 vcache에 보존)
    qvec = vcache.embedder.encode_query([startup.business_text])[0]
    df["embed_score"] = (vcache.vectors @ qvec).astype(np.float32)
    topk = set(df["embed_score"].nlargest(cfg.top_k_embed).index)
    df["lane_embed"] = df.index.isin(topk)

    # ----- Lane A: 코드 -----
    s_maj, s_mid, s_sub = normalize_seed(startup.seed_code)
    df["lane_code_exact"] = df["_sub"] == s_sub
    if cfg.use_mid_fallback:
        same_mid = (df["_mid"] == s_mid) & ((s_maj is None) | (df["_maj"] == s_maj))
        df["lane_code_mid"] = same_mid & (~df["lane_code_exact"])  # exact과 배타 → provenance 명확
    else:
        df["lane_code_mid"] = False

    # ----- Lane C: 동시출현 prior -----
    prior_subs = set((cooccur_prior or {}).get(s_sub, {}).keys())
    df["lane_cooccur"] = df["_sub"].isin(prior_subs)

    # ----- Union + recall_cap -----
    det = df["lane_code_exact"] | df["lane_code_mid"] | df["lane_cooccur"]  # 결정론 히트
    df["in_pool"] = det | df["lane_embed"]
    pool = df[df["in_pool"]].copy()

    # 결정론 히트는 무조건 유지, 임베딩-only 후보는 점수순으로 cap까지만
    if len(pool) > cfg.recall_cap:
        keep_det = pool[det.loc[pool.index]]
        embed_only = pool[~det.loc[pool.index]].sort_values("embed_score", ascending=False)
        room = max(cfg.recall_cap - len(keep_det), 0)
        pool = pd.concat([keep_det, embed_only.head(room)])

    pool = pool.sort_values("embed_score", ascending=False)
    return pool[[c for c in OUTPUT_COLS if c in pool.columns]].reset_index(drop=True)


def save_pool(pool: pd.DataFrame, path: str | Path) -> Path:
    path = Path(path)
    pool.to_csv(path, index=False, encoding="utf-8-sig")  # 엑셀 한글 호환
    return path