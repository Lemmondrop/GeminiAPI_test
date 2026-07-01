"""CLI 진입점 — 신규 IPO 발행사 정보로 피어 추출 4단계 실행.

예:
  python run.py --name "나라스페이스" --seed C00303 ^
      --text "위성 시스템 및 지구관측 데이터 서비스" --track tech_special

  python run.py --name "한라캐스트" --seed C00303 ^
      --text "마그네슘 알루미늄 다이캐스팅 차량 경량화 부품" --track general --max-peers 6
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import config
from pipeline import run_pipeline


def main():
    ap = argparse.ArgumentParser(description="IPO 피어그룹 추출")
    ap.add_argument("--name", required=True, help="발행사명")
    ap.add_argument("--seed", required=True, help="자가입력 KSIC 시드코드 (예: C00303)")
    ap.add_argument("--text", required=True, help="발행사 사업설명")
    ap.add_argument("--track", default="general", choices=list(config.VALID_TRACKS))
    ap.add_argument("--top-k", type=int, default=None, help="Stage3 LLM 판정 후보 수")
    ap.add_argument("--max-peers", type=int, default=None, help="유사도 상위 후보 밴드 크기")
    ap.add_argument("--mcap-max-억", type=float, default=None, dest="mcap_max_eok",
                    help="시총 상한(억). 이상 제외. 0이면 비활성")
    ap.add_argument("--per-min", type=float, default=None, help="PER 하한. 미만 제외. 0이면 비활성")
    ap.add_argument("--per-max", type=float, default=None, help="PER 상한. 이상 제외. 0이면 비활성")
    ap.add_argument("--outlier", choices=["iqr", "trim", "none"], default=None,
                    help="통계적 이상치 처리 (기본 none, 하드컷이 주 필터)")
    ap.add_argument("--headline", choices=["median", "mean", "sim_weighted"], default=None,
                    help="대표 멀티플 (기본 median, sim_weighted=유사도 가중평균)")
    ap.add_argument("--high-sim-exempt", action="store_true", dest="high_sim_exempt",
                    help="고유사도(높음/sim≥85) 피어를 PER 상·하한 컷에서 면제(핵심 피어 보존)")
    # 평가가액 산출 (적자/초기기업: 미래 추정순이익 현가화 × 피어 PER)
    ap.add_argument("--future-ni", type=float, default=None, dest="future_ni",
                    help="평가 기준 미래 추정순이익(백만원). 입력 시 피어 PER로 평가가액 산출")
    ap.add_argument("--ni-year", default=None, dest="ni_year",
                    help="추정순이익 기준연도 라벨(표시용, 예: 2027E)")
    ap.add_argument("--discount-rate", type=float, default=0.20, dest="discount_rate",
                    help="현가 할인율 (기본 0.20)")
    ap.add_argument("--discount-years", type=float, default=2.0, dest="discount_years",
                    help="현가 할인 기간(년, 기본 2)")
    a = ap.parse_args()

    s3cfg = config.Stage3Config(top_k_judge=a.top_k) if a.top_k else None
    s4kw = {}
    if a.max_peers:
        s4kw["max_peers"] = a.max_peers
    if a.mcap_max_eok is not None:
        s4kw["mcap_max"] = None if a.mcap_max_eok == 0 else a.mcap_max_eok * 1e8
    if a.per_min is not None:
        s4kw["per_min"] = None if a.per_min == 0 else a.per_min
    if a.per_max is not None:
        s4kw["per_max"] = None if a.per_max == 0 else a.per_max
    if a.outlier:
        s4kw["outlier_method"] = a.outlier
    if a.headline:
        s4kw["headline"] = a.headline
    if a.high_sim_exempt:
        s4kw["high_sim_exempt"] = True
    s4cfg = config.Stage4Config(**s4kw) if s4kw else None

    print(f"=== {a.name} 피어 추출 (track={a.track}) ===")
    res = run_pipeline(a.name, a.seed, a.text, a.track, s3cfg=s3cfg, s4cfg=s4cfg, verbose=True)

    c = res.counts()
    print(f"\n깔때기: 모집단 {c['모집단']} → 재무통과 {c['재무통과']} "
          f"→ 사업유사 {c['사업유사통과']} → 최종 {c['최종비교군']}\n")

    if res.final.empty:
        # 시총 미수집(KRX 인가 전)이면 Stage3 통과까지를 안내
        s3p = res.stage3[res.stage3["passed_stage3"]] if "passed_stage3" in res.stage3 else res.stage3
        print("최종 비교군 없음 — Stage4(PER) 보류 가능성. Stage3 통과 후보:")
        for _, r in s3p.head(15).iterrows():
            print(f"  {str(r['회사명'])[:14]:<14} {str(r.get('종목코드',''))[:8]:<8} {r['tier']:<2} 유사도 {int(r['sim_score']):>3}")
        if "drop_reason" in res.stage4 and len(res.stage4):
            reasons = set(res.stage4["drop_reason"]) - {""}
            if any("KRX" in x for x in reasons):
                print("\n※ 시총 미수집(KRX 인가 필요). 인가 후 시총 backfill하면 Stage4 작동.")
        return

    print("최종 비교군:")
    for _, r in res.final.iterrows():
        per = f"PER {r['per']:.1f}" if r.get("per") == r.get("per") else "PER —"
        print(f"  {str(r['회사명'])[:14]:<14} {str(r.get('종목코드',''))[:8]:<8} {r['tier']:<2} 유사도 {int(r['sim_score']):>3}  {per}")
        print(f"      └ {r.get('rationale', '')}")
    from collections import Counter
    cuts = Counter(r for r in res.stage4["drop_reason"] if r and "max_peers" not in r)
    if cuts:
        print("\nStage4 컷 내역:", "  ".join(f"{k}×{v}" for k, v in cuts.most_common()))
    if res.per_summary:
        s = res.per_summary
        sw = s.get("per_sim_weighted")
        print(f"\n대표 멀티플(PER, {s['headline']}): {s['headline_per']}   "
              f"[중앙값 {s['per_median']} / 평균 {s['per_mean']}"
              + (f" / 유사도가중 {sw}" if sw is not None else "")
              + f" / 범위 {s['per_min']}~{s['per_max']}, n={s['n']}]")
        if s.get("thin"):
            print(f"⚠ 피어 부족(n={s['n']}) — 하드컷이 빡빡하거나 동종 상장사가 적음. 기준 완화 검토.")

    # 평가가액 산출 — 미래 추정순이익 현가화 × 피어 PER (적자/초기기업 평가)
    if a.future_ni is not None and not res.final.empty:
        from core.valuation import valuation_scenarios
        peers = [{"per": r["per"], "similarity": r.get("sim_score", 50), "tier": r.get("tier", "")}
                 for _, r in res.final.iterrows() if r.get("per") == r.get("per")]
        if peers:
            v = valuation_scenarios(peers, a.future_ni, a.discount_rate, a.discount_years)
            yr = f"({a.ni_year}) " if a.ni_year else ""
            print(f"\n=== 평가가액 ===")
            print(f"추정순이익 {yr}{a.future_ni:,.0f}백만원 → 현가화"
                  f"(할인율 {a.discount_rate*100:.0f}%, {a.discount_years:g}년)"
                  f" = 적용순이익 {v['applied_ni']:,.0f}백만원 ({v['applied_ni']/100:,.0f}억)")
            labels = {"median": "중앙값(보수)", "sim_weighted": "유사도가중",
                      "high_only": "높음tier만", "mean": "평균"}
            for m in ("median", "sim_weighted", "high_only", "mean"):
                sc = v["scenarios"][m]
                print(f"  {labels[m]:<12} PER {sc['per']:5.1f} → 평가시총 "
                      f"{sc['equity_value']/100:>8,.0f}억")
            print("  ※ 적자/초기기업이라 현재 PER 평가 불가 → 미래 추정순이익 현가화 방식(주관사 표준)")


if __name__ == "__main__":
    main()