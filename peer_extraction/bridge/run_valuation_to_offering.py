r"""End-to-end 브리지 CLI — 피어추출 평가시총 + discount 모델 → 예상 공모시총(value-band).

전제: discount_rate_pipeline의 input_limited 모델 .pkl 필요
      (v1/general/v2는 price_vs_mid 등 순환 피처라 사전평가 불가).

흐름:
  피어추출 → 대표PER·평가시총 → [이 브리지: 할인율 예측] → 예상 공모시총 = 평가시총 ×(1−할인율)

예:
  python bridge\run_valuation_to_offering.py ^
    --model ..\discount_rate_pipeline\model_input_limited.pkl ^
    --scenarios "median:12.6:1332,sim_weighted:29.0:3070,high_only:47.8:5061" ^
    --net-income-bn 152.3 --track general --market 코스닥 ^
    --segments segment_security,segment_software
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import joblib

from bridge.predict_discount import predict_discount, apply_discount_to_valuation


def _segments(s):
    return {k.strip(): 1 for k in s.split(",")} if s else {}


def main():
    ap = argparse.ArgumentParser(description="평가시총 → 예상 공모시총 (discount 모델 연결)")
    ap.add_argument("--model", required=True, help="input_limited 모델 .pkl")
    ap.add_argument("--peer-per", type=float, default=None)
    ap.add_argument("--equity-value-eok", type=float, default=None)
    ap.add_argument("--scenarios", default=None,
                    help='다중: "라벨:PER:평가시총억,..."')
    ap.add_argument("--net-income-bn", type=float, default=None)
    ap.add_argument("--per-stock-value", type=float, default=None)
    ap.add_argument("--track", default="general",
                    choices=["general", "tech_special", "growth_special", "konex_transfer"])
    ap.add_argument("--market", default="코스닥")
    ap.add_argument("--valuation-model", default="PER")
    ap.add_argument("--segments", default=None)
    a = ap.parse_args()

    mi = joblib.load(a.model)
    ver = mi.get("version", "?")
    if "input_limited" not in str(ver):
        print(f"⚠ 경고: version='{ver}'. 사전평가엔 input_limited 모델만 유효"
              "(v1/general/v2는 price_vs_mid 등 순환 피처 의존).")

    common = dict(
        net_income_bn=a.net_income_bn, per_stock_value=a.per_stock_value,
        is_tech_special=int(a.track == "tech_special"),
        is_growth_special=int(a.track == "growth_special"),
        listing_market=a.market, valuation_model=a.valuation_model,
        segments=_segments(a.segments),
    )

    scen = []
    if a.scenarios:
        for part in a.scenarios.split(","):
            label, per, ev = part.split(":")
            scen.append((label, float(per), float(ev)))
    elif a.peer_per is not None and a.equity_value_eok is not None:
        scen.append(("single", a.peer_per, a.equity_value_eok))
    else:
        ap.error("--scenarios 또는 (--peer-per + --equity-value-eok) 필요")

    print(f"=== 평가시총 → 예상 공모시총 (모델: {ver}) ===")
    print(f"{'시나리오':<14}{'대표PER':>8}{'예측할인율':>10}{'평가시총':>10}{'예상공모시총':>12}")
    for label, per, ev in scen:
        d = predict_discount(mi, peer_per=per, **common)
        pub = apply_discount_to_valuation(ev, d)
        print(f"{label:<14}{per:>8.1f}{d:>9.1f}%{ev:>9,.0f}억{pub:>11,.0f}억")
    print("\n※ 예상 공모시총 = 평가시총 × (1 − 예측할인율). 회사 가치(value-band) 레벨 산출.")
    print("  적자/비상장이라 dart 재무는 결측(모델 missing-flag 처리). 모델은 베이스라인(데이터 보강 시 재학습).")


if __name__ == "__main__":
    main()