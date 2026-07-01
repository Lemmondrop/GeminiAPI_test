"""Stage 3 라이브 스모크 — 한라캐스트 stage1→2→3 연결, 실제 Gemini 사업유사성 판정.

stage2 통과 피어를 embed_score 상위 K개 → Gemini가 tier/score/근거 판정 → 최종 피어.
1회 실행 후 판정은 캐시됨 → 재실행 시 추가 API 호출 0, 동일 결과.

선행: 코퍼스 캐시 + universe_financials.csv 존재. .env GEMINI_API_KEY.
실행: peer_extraction 루트에서  python tests/smoke_stage3.py

※ GEMINI_LLM_MODEL(config) 접근 안 되면 에러 → 본인 접근 가능한 모델로 조정.
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from core.embedder import GeminiEmbedder
from core.llm_judge import LlmJudge
from core.vector_cache import VectorCache
from stages.stage1_population import StartupQuery, build_population
from stages.stage2_financial import apply_stage2, load_financials, survivors
from stages.stage3_business import apply_stage3, business_text, final_peers

TRACK = "general"
PRIOR = {"303": {"242": 1.0, "281": 1.0}}
FIRM = "한라캐스트"


def main():
    if not config.UNIVERSE_FIN_CSV.exists():
        sys.exit(f"[중단] 재무 CSV 없음: {config.UNIVERSE_FIN_CSV}")

    vc = VectorCache(GeminiEmbedder())
    vc.build_or_load()
    row = vc.df[vc.df[config.NAME_COL] == FIRM].iloc[0]
    issuer_text = business_text(row)
    q = StartupQuery(seed_code=str(row[config.CODE_COL]), business_text=str(row["주요제품"]))

    pool = build_population(q, vc, cooccur_prior=PRIOR,
                            cfg=config.Stage1Config(top_k_embed=30))
    s2 = apply_stage2(pool, load_financials(), TRACK)
    surv = survivors(s2)

    cfg3 = config.Stage3Config()
    print(f"stage2 통과 {len(surv)} → Stage3 LLM 판정 (embed 상위 {cfg3.top_k_judge})")
    print(f"발행사 사업: {issuer_text}")
    print(f"모델: {config.GEMINI_LLM_MODEL}\n")

    judge = LlmJudge()
    s3 = apply_stage3(surv, issuer_text, judge, cfg3)
    fp = final_peers(s3)

    print(f"판정 결과 (최종 피어 {len(fp)}개):\n")
    for _, r in s3.iterrows():
        mark = "✓" if r["passed_stage3"] else "✗"
        print(f"  [{mark}] {str(r['회사명'])[:12]:<12} {r['tier']:<2} {int(r['sim_score']):>3}  {r['rationale']}")

    print(f"\n캐시 저장: {config.LLM_JUDGE_CACHE}")
    print("→ 재실행하면 추가 API 호출 없이 동일 결과(재현성). 프롬프트 개정 시 LLM_PROMPT_VERSION ↑.")


if __name__ == "__main__":
    main()
