"""피어 추출 파이프라인 오케스트레이션 (track_type 1파라미터).

stage1(모집단) → stage2(재무) → stage3(사업유사성·LLM) → stage4(PER 압축) → 최종 비교군.
트랙별 스크립트 복제 없이 track_type 하나로 분기. 각 단계 산출물(provenance 포함)을
함께 반환해 감사 가능. 의존성(vcache/judge/financials)은 주입 가능(테스트·재사용).
"""
from __future__ import annotations
from dataclasses import dataclass

import pandas as pd

import config
from stages.stage1_population import StartupQuery, build_population
from stages.stage2_financial import apply_stage2, load_financials, survivors as s2_survivors
from stages.stage3_business import apply_stage3, final_peers as s3_final
from stages.stage4_general import apply_stage4, final_comps, peer_per_summary


@dataclass
class PipelineResult:
    pool: pd.DataFrame        # stage1 모집단
    stage2: pd.DataFrame      # 재무 필터 (passed_stage2/drop_reason)
    stage3: pd.DataFrame      # 사업유사성 (tier/sim_score/rationale)
    stage4: pd.DataFrame      # PER 압축 (per/passed_stage4/drop_reason)
    final: pd.DataFrame       # 최종 비교군
    per_summary: dict         # 최종 PER 통계

    def counts(self) -> dict:
        return {
            "모집단": int(len(self.pool)),
            "재무통과": int(self.stage2["passed_stage2"].sum()) if "passed_stage2" in self.stage2 else 0,
            "사업유사통과": int(self.stage3["passed_stage3"].sum()) if "passed_stage3" in self.stage3 else 0,
            "최종비교군": int(len(self.final)),
        }


def run_pipeline(
    issuer_name: str,
    issuer_seed_code: str,
    issuer_business_text: str,
    track_type: str = "general",
    *,
    cooccur_prior: dict | None = None,
    vcache=None,
    judge=None,
    financials: pd.DataFrame | None = None,
    s1cfg: config.Stage1Config | None = None,
    s3cfg: config.Stage3Config | None = None,
    s4cfg: config.Stage4Config | None = None,
    verbose: bool = False,
) -> PipelineResult:
    if track_type not in config.VALID_TRACKS:
        raise ValueError(f"track_type must be one of {config.VALID_TRACKS}")

    def log(msg):
        if verbose:
            print(msg, flush=True)

    # 의존성 (미주입 시 운영 기본: Gemini 임베더/판정기 + 캐시)
    if vcache is None:
        from core.embedder import GeminiEmbedder
        from core.vector_cache import VectorCache
        vcache = VectorCache(GeminiEmbedder())
    log("  [1/4] 코퍼스 로드 + 모집단 구성...")
    vcache.build_or_load()
    if judge is None:
        from core.llm_judge import LlmJudge
        judge = LlmJudge()
    if financials is None:
        financials = load_financials()

    # Stage 1 — 모집단(업종 유사성)
    if cooccur_prior is None:                         # 미주입 시 데이터 산출 prior 자동 로드
        from core.cooccur import load_prior
        loaded = load_prior()
        cooccur_prior = loaded or None                # 없으면 None(Lane C 비활성)
    q = StartupQuery(seed_code=issuer_seed_code, business_text=issuer_business_text)
    pool = build_population(q, vcache, cooccur_prior=cooccur_prior,
                            cfg=s1cfg or config.Stage1Config())
    # 발행사 본인 배제 (상장 동명이 유니버스에 있을 때 자기참조 방지; 비상장이면 무영향)
    pool = pool[pool[config.NAME_COL] != issuer_name].reset_index(drop=True)

    # Stage 2 — 재무 필터(track_type별)
    log(f"  [2/4] 재무 필터 (track={track_type})...")
    s2 = apply_stage2(pool, financials, track_type)
    surv = s2_survivors(s2)

    # Stage 3 — 사업 유사성(임베딩 선별 + LLM 근거, 캐싱)
    cfg3 = s3cfg or config.Stage3Config()
    log(f"  [3/4] 사업 유사성 LLM 판정 (상위 {cfg3.top_k_judge}, 신규는 건당 수초)...")
    s3 = apply_stage3(surv, issuer_business_text, judge, cfg3, progress=verbose)

    # Stage 4 — 유사도 밴드 → PER 이상치 제거 → 최종 비교군
    log("  [4/4] 유사도 밴드 → PER 이상치 제거 → 최종 비교군...")
    cfg4 = s4cfg or config.Stage4Config()
    s4 = apply_stage4(s3_final(s3), cfg4)
    final = final_comps(s4)

    return PipelineResult(pool=pool, stage2=s2, stage3=s3, stage4=s4,
                          final=final, per_summary=peer_per_summary(s4, cfg4))