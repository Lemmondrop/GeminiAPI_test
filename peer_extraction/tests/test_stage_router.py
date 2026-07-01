"""test_stage_router.py — 이전 세션에서 사람이 수동으로 분류했던 4개 실제
케이스(센스톤/스쿼드엑스/NUCLUX.AI/케이더블유씨)를 자동 판별기에 다시 넣어
같은 결론이 나오는지 검증. + 엣지케이스(신호 없음, 신호 상충) 확인.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from stage_router import StageSignals, determine_company_stage


def show(decision):
    print(f"\n=== {decision.company_name} ===")
    print(f"  판정: {decision.track}"
          + (f" / {decision.sub_track}" if decision.sub_track else ""))
    print(f"  신뢰도: {decision.confidence:.2f}  "
          f"(IPO {decision.score_ipo:.1f} vs 다음라운드 {decision.score_next_round:.1f})")
    print(f"  review_needed: {decision.review_needed}")
    for r in decision.reasoning:
        print(f"    - {r}")


results = []

# ── 1. 센스톤(SSenStone) — 실제로 IPO 트랙(PER, 공모시총 경로)으로 처리했던 케이스 ──
s1 = StageSignals(
    company_name="센스톤 (SSenStone)",
    founding_year=2015,
    latest_round="Pre-IPO",
    ipo_mentioned=True,
    ipo_target_year=2027,
    net_income=None,
    future_net_income=152.3,  # 2027E 15,230백만원 = 152.3억
    revenue=None,
    current_year=2026,
)
d1 = determine_company_stage(s1)
show(d1)
assert d1.track == "ipo_track", "센스톤은 ipo_track이어야 함"
assert d1.sub_track == "tech_special", "센스톤은 미래 추정치 기반이므로 기술특례여야 함"
results.append(("센스톤", d1.track, d1.sub_track))

# ── 2. 스쿼드엑스(SQUADX) — 실제로 다음라운드 트랙(PSR)으로 처리했던 케이스 ──
s2 = StageSignals(
    company_name="스쿼드엑스 (SQUADX)",
    founding_year=2022,
    latest_round="Series A",
    ipo_mentioned=False,
    ipo_target_year=None,
    net_income=None,
    revenue=46.3,
    current_year=2026,
)
d2 = determine_company_stage(s2)
show(d2)
assert d2.track == "next_round_track", "스쿼드엑스는 next_round_track이어야 함"
results.append(("스쿼드엑스", d2.track, d2.sub_track))

# ── 3. NUCLUX.AI — Seed+Angel, 2025년 5월 설립 → 다음라운드 트랙 ──
s3 = StageSignals(
    company_name="NUCLUX.AI",
    founding_year=2025,
    latest_round="Seed+Angel",
    ipo_mentioned=False,
    ipo_target_year=None,
    current_year=2026,
)
d3 = determine_company_stage(s3)
show(d3)
assert d3.track == "next_round_track", "NUCLUX.AI는 next_round_track이어야 함"
results.append(("NUCLUX.AI", d3.track, d3.sub_track))

# ── 4. 케이더블유씨(KWC) — "26년 IPO 준비 착수", 순이익 적자 → IPO/기술특례 ──
s4 = StageSignals(
    company_name="주식회사 케이더블유씨 (KWC)",
    founding_year=None,
    latest_round=None,
    ipo_mentioned=True,
    ipo_target_year=2028,
    net_income=-5.0,
    future_net_income=25.0,
    current_year=2026,
)
d4 = determine_company_stage(s4)
show(d4)
assert d4.track == "ipo_track", "케이더블유씨는 ipo_track이어야 함"
assert d4.sub_track == "tech_special", "케이더블유씨는 적자+미래추정 있으므로 기술특례여야 함"
results.append(("케이더블유씨", d4.track, d4.sub_track))

# ── 엣지케이스 A: 신호가 전혀 없음 ──
s5 = StageSignals(company_name="신호없음테스트")
d5 = determine_company_stage(s5)
show(d5)
assert d5.track == "next_round_track", "신호 없으면 안전한 기본값(다음라운드)이어야 함"
assert d5.review_needed is True

# ── 엣지케이스 B: 상충 신호 (Seed 단계인데 IPO 언급) ──
s6 = StageSignals(
    company_name="상충신호테스트",
    latest_round="Seed",
    ipo_mentioned=True,
    ipo_target_year=2027,
    current_year=2026,
)
d6 = determine_company_stage(s6)
show(d6)
assert d6.review_needed is True, "상충 신호는 review_needed=True여야 함"

# ── 엣지케이스 C: manual override ──
s7 = StageSignals(
    company_name="override테스트",
    latest_round="Series A",
    ipo_mentioned=False,
    current_year=2026,
)
d7 = determine_company_stage(s7, manual_override="ipo_track")
show(d7)
assert d7.track == "ipo_track"
assert d7.source == "manual_override"

print("\n" + "=" * 60)
print("전체 결과 요약 (이전 세션 수동 판단과 비교)")
print("=" * 60)
expected = {
    "센스톤": "ipo_track",
    "스쿼드엑스": "next_round_track",
    "NUCLUX.AI": "next_round_track",
    "케이더블유씨": "ipo_track",
}
for name, track, sub in results:
    match = "✅" if expected[name] == track else "❌"
    print(f"  {match} {name}: {track} {f'({sub})' if sub else ''}  [기대값: {expected[name]}]")

print("\n모든 assertion 통과 — 4개 실사례 + 3개 엣지케이스 검증 완료.")

# ── 2026-07 실사례로 발견된 케이스 2개 — 영구 회귀 테스트로 추가 ──

# 케이스 A: "Series Pre-A" 표기 인식 (기존 정규식이 놓쳤던 실제 IR 표기)
from stage_router import normalize_round
assert normalize_round("Series Pre-A").value == "seed", \
    "Series Pre-A가 인식 안 됨 — 2026-07 스쿼드엑스 PDF 실사례에서 발견된 회귀"
print("\n✅ Series Pre-A 라운드 인식 (회귀 테스트)")

# 케이스 B: 신호 1개뿐인데 confidence가 부풀려지던 문제 (피엔피 실사례)
# + 절대점수 최소선 미달로 track 자체도 next_round_track으로 안전 처리되는지
s_thin = StageSignals(company_name="피엔피_회귀테스트", founding_year=2011, current_year=2026)
d_thin = determine_company_stage(s_thin)
assert d_thin.confidence < 0.3, \
    f"신호 1개짜리(총점 1.0)인데 confidence가 여전히 높음: {d_thin.confidence}"
assert d_thin.review_needed is True
assert d_thin.track == "next_round_track", \
    (f"약한 신호 1개(+1.0)만으로 ipo_track이 되면 안 됨(실사례 버그) — "
     f"실제: {d_thin.track}")
print(f"✅ 얇은 증거 → next_round_track 기본값 + confidence 보정 "
      f"(피엔피 실사례, track={d_thin.track}, confidence={d_thin.confidence:.2f})")

print("\n전체 회귀 테스트 통과.")