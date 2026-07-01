"""파이프라인 라이브 스모크 — 한라캐스트 4단계 엔드투엔드.

실제 GeminiEmbedder + LlmJudge + universe_financials로 stage1→2→3→4 전 구간 실행.
코퍼스/LLM 캐시를 재사용하므로 저렴(쿼리 1콜 + 미캐시 후보만 LLM).
Stage4는 시총(KRX) 필요 — 미수집이면 보류 표시되고 Stage3 통과까지 출력된다.

실행: peer_extraction 루트에서  python tests/smoke_pipeline.py
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from core.embedder import GeminiEmbedder
from core.llm_judge import LlmJudge
from core.vector_cache import VectorCache
from pipeline import run_pipeline
from stages.stage3_business import business_text

FIRM = "한라캐스트"
TRACK = "general"
PRIOR = {"303": {"242": 1.0, "281": 1.0}}


def main():
    if not config.UNIVERSE_FIN_CSV.exists():
        sys.exit(f"[중단] 재무 CSV 없음: {config.UNIVERSE_FIN_CSV}")

    vc = VectorCache(GeminiEmbedder())
    vc.build_or_load()
    row = vc.df[vc.df[config.NAME_COL] == FIRM].iloc[0]
    issuer_text = business_text(row)
    seed = str(row[config.CODE_COL])

    print(f"=== {FIRM} 파이프라인 (track={TRACK}) ===")
    print(f"발행사 사업: {issuer_text}\n")

    res = run_pipeline(FIRM, seed, issuer_text, TRACK,
                       cooccur_prior=PRIOR, vcache=vc, judge=LlmJudge(), verbose=True)

    c = res.counts()
    print(f"깔때기: 모집단 {c['모집단']} → 재무통과 {c['재무통과']} "
          f"→ 사업유사 {c['사업유사통과']} → 최종 {c['최종비교군']}\n")

    if res.final.empty:
        s3p = res.stage3[res.stage3["passed_stage3"]]
        print(f"Stage4 보류(시총 미수집 가능) — Stage3 통과 {len(s3p)}개 상위:")
        for _, r in s3p.head(12).iterrows():
            print(f"  {str(r['회사명'])[:14]:<14} {r['tier']:<2} 유사도 {int(r['sim_score']):>3}")
        if len(res.stage4) and any("KRX" in x for x in set(res.stage4["drop_reason"])):
            print("\n※ 시총 미수집(KRX 인가 필요). 인가 후 --mcap backfill하면 Stage4 작동.")
    else:
        from collections import Counter
        cuts = Counter(r for r in res.stage4["drop_reason"] if r and "max_peers" not in r)
        if cuts:
            print("Stage4 컷 내역:", "  ".join(f"{k}×{v}" for k, v in cuts.most_common()))
        print(f"\n최종 비교군 {len(res.final)}개:")
        for _, r in res.final.iterrows():
            print(f"  {str(r['회사명'])[:14]:<14} {r['tier']:<2} 유사도 {int(r['sim_score']):>3}  PER {r['per']:.1f}")
        s = res.per_summary
        print(f"\n대표 멀티플(PER, {s['headline']}): {s['headline_per']}   "
              f"[중앙값 {s['per_median']} / 평균 {s['per_mean']} / 범위 {s['per_min']}~{s['per_max']}]")
        if s.get("thin"):
            print(f"⚠ 피어 부족(n={s['n']} < 권장 {3}) — 하드컷이 빡빡하거나 동종 상장사가 적음. 기준 완화 검토.")


if __name__ == "__main__":
    main()