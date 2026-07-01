"""Stage 2 실데이터 스모크 — 한라캐스트 모집단에 재무 필터 적용.

stage1(한라캐스트 모집단) → universe_financials 조인 → 트랙별 재무 필터.
통과 수 · drop 사유 분포 · 알려진 피어 4개의 통과/탈락을 출력.

선행: smoke_real로 코퍼스 캐시 생성됨 + collect_financials로 universe_financials.csv 생성됨.
실행: peer_extraction 루트에서  python tests/smoke_stage2.py

※ 주의: 재무는 현재시점(2025 온기 / 2026 1Q LTM) 기준이다. 한라캐스트의 과거 신고서
   피어선정과 1:1 백테스트는 아니다(그건 신고서 시점 재무가 필요). 여기선 필터의
   '현재 데이터에 대한 작동'을 확인한다 — 라이브 신규 IPO에선 현재시점이 맞다.
"""
from __future__ import annotations
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from core.embedder import GeminiEmbedder
from core.vector_cache import VectorCache
from stages.stage1_population import StartupQuery, build_population
from stages.stage2_financial import apply_stage2, load_financials, survivors

TRACK = "general"                                   # 한라캐스트 가정 트랙
PRIOR = {"303": {"242": 1.0, "281": 1.0}}
CALIBRATION_FIRM = "한라캐스트"
KNOWN_PEERS = ["알루코", "조일알미늄", "파워로직스", "삼현"]


def main():
    if not config.UNIVERSE_FIN_CSV.exists():
        sys.exit(f"[중단] 재무 CSV 없음: {config.UNIVERSE_FIN_CSV} "
                 f"(먼저 scripts/collect_financials.py 실행)")

    vc = VectorCache(GeminiEmbedder())
    vc.build_or_load()                              # 캐시 적중(즉시) + 쿼리 1콜
    row = vc.df[vc.df[config.NAME_COL] == CALIBRATION_FIRM].iloc[0]
    q = StartupQuery(seed_code=str(row[config.CODE_COL]), business_text=str(row["주요제품"]))
    pool = build_population(q, vc, cooccur_prior=PRIOR,
                            cfg=config.Stage1Config(top_k_embed=30))
    print(f"stage1 모집단: {len(pool)}행")

    fin = load_financials()
    s2 = apply_stage2(pool, fin, TRACK)
    surv = survivors(s2)
    print(f"stage2 통과: {len(surv)}/{len(s2)}  (track={TRACK})\n")

    reasons = Counter()
    for r in s2.loc[~s2["passed_stage2"], "drop_reason"]:
        for part in str(r).split(";"):
            if part:
                reasons[part] += 1
    print("drop 사유 분포:")
    for k, v in reasons.most_common():
        print(f"  {v:>3}  {k}")

    print("\n알려진 피어 4개 상태:")
    for nm in KNOWN_PEERS:
        r = s2[s2[config.NAME_COL] == nm]
        if r.empty:
            print(f"  {nm}: 모집단에 없음")
            continue
        r = r.iloc[0]
        status = "통과 ✓" if r["passed_stage2"] else f"탈락 ({r['drop_reason']})"
        ni = r.get("net_income_ltm")
        ni_s = f"{ni/1e8:,.0f}억" if isinstance(ni, (int, float)) and ni == ni else "—"
        print(f"  {nm}: {status}  | net_LTM={ni_s}")

    print("\n참고: 파워로직스는 현재 LTM 순이익이 적자라 general(net_ltm>0)에선 탈락이 정상이야.")
    print("      과거 신고서 선정과의 1:1 검증은 신고서 시점 재무가 필요(별도).")


if __name__ == "__main__":
    main()
