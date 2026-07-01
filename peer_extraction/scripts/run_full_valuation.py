"""
scripts/run_full_valuation.py
────────────────────────────────────────────────────────────
poc_results/{회사명}.json(Stage1~4 피어 추출 결과)을 읽어, IR 신호로
IPO트랙/다음라운드트랙을 자동 판별한 뒤(core.stage_router), 실제
universe_financials.csv 조인을 거쳐 최종 가치(밴드)까지 한 번에
산출하는 통합 스크립트.

이전 버전과의 차이: determine_valuation_method()를 라우터로 쓰던 방식을
완전히 제거하고, stage_decision.track / stage_decision.sub_track 하나로
PSR/PER 경로 및 general/tech_special 서브트랙까지 전부 결정한다.
────────────────────────────────────────────────────────────
"""
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime

SCRIPT_DIR = Path(__file__).parent
ROOT_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(ROOT_DIR))

from core.peer_join import enrich_peers_with_financials, load_financials_lookup
from core.stage_router import StageSignals, determine_company_stage
from core.valuation_psr import calculate_representative_psr, calculate_equity_value_psr_band
# equity_value는 core.valuation의 '함수'명이자 이 파일에서 계산 결과를 담던
# 지역변수명이기도 했다 — 예전 'decision' 충돌과 같은 패턴이라 별칭으로 회피.
from core.valuation import present_value, representative_per, equity_value as calc_equity_value
from bridge.predict_discount import predict_discount, apply_discount_to_valuation

POC_RESULTS_DIR = ROOT_DIR / "poc_results"
# discount_rate_pipeline은 peer_extraction의 형제 폴더에 있음.
DISCOUNT_MODEL_PATH = ROOT_DIR.parent / "discount_rate_pipeline" / "model_input_limited.pkl"

_discount_model_cache = None


def load_discount_model():
    """공모할인율 예측 모델(input_limited)을 1회 로드해 캐시.

    v1/general/v2 모델은 공모가를 이미 알아야 하는 피처(price_vs_mid 등)에
    의존해 비상장 사전평가에 못 쓴다 — input_limited만 유효(memory 기준 확정).
    """
    global _discount_model_cache
    if _discount_model_cache is None:
        import joblib
        if not DISCOUNT_MODEL_PATH.exists():
            raise FileNotFoundError(
                f"공모할인율 모델을 찾을 수 없습니다: {DISCOUNT_MODEL_PATH}")
        _discount_model_cache = joblib.load(DISCOUNT_MODEL_PATH)
    return _discount_model_cache


def load_poc_result(case_name: str) -> dict:
    path = POC_RESULTS_DIR / f"{case_name}.json"
    if not path.exists():
        available = [p.stem for p in POC_RESULTS_DIR.glob("*.json") if p.stem != "_summary"]
        sys.exit(f"[ERROR] {path} 없음.\n  사용 가능한 케이스: {available}")
    return json.loads(path.read_text(encoding="utf-8"))


def run_one(case: str, signals: dict, verbose: bool = True) -> dict:
    """한 케이스에 대해 [1]PoC로드 → [2]단계판별 → [3]가치산출까지 전부 수행.

    signals: StageSignals 필드와 동일한 키(founding_year, latest_round,
    ipo_mentioned, ipo_target_year, net_income, future_net_income, revenue,
    yoy_growth, decay_factor, psr_method)를 담은 dict. 없는 키는 기본값 사용.
    CLI(main())와 배치 스크립트(run_all_cases.py)가 이 함수 하나를 공유한다
    — 로직이 두 곳에 따로 있으면 예전 'decision' 이중 라우팅 문제가 다시
    생기므로, 계산 로직의 단일 진실 공급원은 이 함수 하나로 고정한다.

    반환값에는 stage_decision, 산출 결과(가능한 경우), 산출 불가 사유를 담아
    배치 스크립트가 표로 요약할 수 있게 한다.
    """
    yoy_growth = signals.get("yoy_growth", 0.0) or 0.0
    decay_factor = signals.get("decay_factor", 0.6) or 0.6
    psr_method = signals.get("psr_method", "median") or "median"
    per_method = signals.get("per_method", "median") or "median"
    time_discount_rate = signals.get("time_discount_rate", 0.20)
    time_discount_rate = 0.20 if time_discount_rate is None else time_discount_rate
    skip_offering_discount = bool(signals.get("skip_offering_discount", False))
    current_year = signals.get("current_year") or datetime.now().year

    def log(msg=""):
        if verbose:
            print(msg)

    # ── 1. PoC 결과 로드 ─────────────────────────────────────
    log(f"[1] PoC 결과 로드: {case}")
    poc = load_poc_result(case)
    # Stage1~4 파이프라인의 실제 출력 키는 'peers'가 아니라 'final_peers'다.
    # (poc_results/*.json 최상위 키: case/elapsed_sec/funnel/final_peers/
    #  headline/no_peers/stage3_candidates_only)
    peers = poc.get("final_peers", [])
    log(f"    피어 {len(peers)}개")

    lookup = load_financials_lookup()
    enriched = enrich_peers_with_financials(peers, lookup)

    # ── 2. 사업 단계 자동판별 (IPO 트랙 vs 다음라운드 트랙) — 유일한 라우팅 지점 ──
    log(f"\n[2] 사업 단계 판별 (IR 신호 기준)...")
    stage_decision = determine_company_stage(StageSignals(
        company_name=case,
        founding_year=signals.get("founding_year"),
        latest_round=signals.get("latest_round"),
        ipo_mentioned=bool(signals.get("ipo_mentioned", False)),
        ipo_target_year=signals.get("ipo_target_year"),
        net_income=signals.get("net_income"),
        future_net_income=signals.get("future_net_income"),
        revenue=signals.get("revenue"),
        current_year=current_year,
    ))
    log(f"    → {stage_decision.track}"
        + (f" / {stage_decision.sub_track}" if stage_decision.sub_track else ""))
    log(f"    신뢰도 {stage_decision.confidence:.2f}"
        + ("  ⚠ 확인 필요" if stage_decision.review_needed else ""))
    for r in stage_decision.reasoning:
        log(f"      - {r}")

    # ── 3. 트랙에 따른 기업가치 산출 ─────────────────────────
    log(f"\n[3] 기업가치 산출...")
    result = {"case": case, "stage_decision": stage_decision,
              "equity_value_억원": None, "value_range_억원": None, "warning": None}

    try:
        if stage_decision.track == "next_round_track":
            revenue = signals.get("revenue")
            if revenue is None:
                msg = "next_round_track인데 매출 데이터 없음 — PSR 산출 불가."
                log(f"    [경고] {msg}")
                result["warning"] = msg
            else:
                rep_psr = calculate_representative_psr(enriched, method=psr_method)
                n_psr = sum(1 for p in enriched if p.get('psr') is not None)
                log(f"    대표 PSR({psr_method}): {rep_psr:.3f}배  (조인 성공 피어 {n_psr}개 기준)")

                band = calculate_equity_value_psr_band(
                    revenue_억원=revenue,
                    representative_psr=rep_psr,
                    yoy_growth_rate=yoy_growth,
                    decay_factor=decay_factor,
                )
                log(f"\n    [시나리오별 기업가치]")
                tr = band["scenarios"]["conservative_trailing"]
                log(f"      보수(trailing, 성장 미반영)        : {tr['equity_value_억원']:>8,.1f}억원")
                if "base_decayed" in band["scenarios"]:
                    bd = band["scenarios"]["base_decayed"]
                    log(f"      기준(YoY {yoy_growth*100:.0f}% × 둔화{decay_factor:.1f}"
                        f" = {bd['decayed_growth_rate']*100:.0f}% 성장 반영) : "
                        f"{bd['equity_value_억원']:>8,.1f}억원")
                    ag = band["scenarios"]["aggressive_naive"]
                    log(f"      공격(YoY {yoy_growth*100:.0f}% 둔화없이 그대로 외삽)  : "
                        f"{ag['equity_value_억원']:>8,.1f}억원  (참고용 — 비현실적 상한)")
                log(f"\n    ▶ 추정 기업가치: 약 {band['low_억원']:.0f}억"
                    + (f" ~ {band['mid_억원']:.0f}억원"
                       f"  (둔화 미반영 외삽 시 상한 {band['high_억원']:.0f}억원)"
                       if band['mid_억원'] else f" ~ {band['high_억원']:.0f}억원"))
                result["value_range_억원"] = (band["low_억원"], band.get("mid_억원") or band["high_억원"])

        elif stage_decision.track == "ipo_track":
            net_income = signals.get("net_income")
            future_net_income = signals.get("future_net_income")
            ni_for_calc = net_income if stage_decision.sub_track == "general" else future_net_income
            if ni_for_calc is None:
                msg = "ipo_track이지만 서브트랙에 필요한 순이익 데이터 없음 — PER 산출 불가."
                log(f"    [경고] {msg}")
                result["warning"] = msg
            else:
                # 할인연수: general은 트레일링 실적이라 항상 0년(할인 없음).
                # tech_special은 --discount-years 명시가 있으면 그걸 쓰고,
                # 없으면 ni_year(추정 순이익 기준연도) - 현재연도로 자동 유도한다.
                if stage_decision.sub_track == "general":
                    years = 0.0
                else:
                    years = signals.get("discount_years")
                    if years is None:
                        ni_year = signals.get("ni_year")
                        if ni_year is not None:
                            years = ni_year - current_year
                            log(f"    할인연수 자동 계산: {ni_year} - {current_year} = {years}년")

                if stage_decision.sub_track == "tech_special" and years is None:
                    msg = ("tech_special인데 할인연수를 결정할 수 없음"
                           "(ni_year 또는 discount_years 필요) — PER 산출 불가.")
                    log(f"    [경고] {msg}")
                    result["warning"] = msg
                else:
                    # representative_per()는 피어 dict에 'similarity' 키를 기대하는데
                    # enrich_peers_with_financials()의 출력은 'sim_score'다 — 이름이
                    # 달라 조용히 기본값(50)으로 깔리는 걸 막기 위해 명시적으로 매핑.
                    peer_input = [
                        {"name": p.get("name"), "per": p.get("per"),
                         "similarity": p.get("sim_score"), "tier": p.get("tier")}
                        for p in enriched
                    ]
                    per_values = [p["per"] for p in peer_input if p.get("per") is not None]
                    if not per_values:
                        msg = "재조인된 피어 중 PER 산출 가능한 곳이 없음 — 재계산 PER 없음."
                        log(f"    [경고] {msg}")
                        result["warning"] = msg
                    else:
                        rep_per = representative_per(peer_input, method=per_method)
                        applied_ni = present_value(ni_for_calc, time_discount_rate, years)
                        val = calc_equity_value(applied_ni, rep_per)
                        log(f"    서브트랙: {stage_decision.sub_track}  "
                            f"(순이익 {ni_for_calc}억원, 시간가치할인율 {time_discount_rate*100:.0f}% × {years:g}년"
                            f" → 적용순이익 {applied_ni:,.1f}억원)")
                        log(f"    대표 PER({per_method}): {rep_per:.2f}배  (n={len(per_values)})")
                        log(f"    ▶ 평가시총: {val:,.1f}억원")
                        result["equity_value_억원"] = val
                        result["offering_market_cap_억원"] = None

                        # ── 공모할인율 예측 (ML, input_limited) → 예상 공모시총 ──
                        if skip_offering_discount:
                            log(f"    (공모할인율 예측 생략 — skip_offering_discount=True)")
                        else:
                            try:
                                model_info = load_discount_model()
                                seg_key = signals.get("segment")
                                segments = {f"segment_{seg_key}": 1} if seg_key else {"segment_unknown": 1}
                                offering_discount_pct = predict_discount(
                                    model_info,
                                    peer_per=rep_per,
                                    per_stock_value=None,       # 이 파이프라인은 주당가치를 만들지 않음 → median 대체
                                    net_income_bn=applied_ni / 10.0,   # 억원 → 10억원(bn) 단위 환산
                                    dart=None,                  # 비상장사 DART 재무 구조적 부재 → median 대체
                                    is_tech_special=int(stage_decision.sub_track == "tech_special"),
                                    is_growth_special=int(bool(signals.get("is_growth_special", False))),
                                    listing_market="코스닥",
                                    valuation_model="PER",
                                    segments=segments,
                                )
                                offering_val = apply_discount_to_valuation(val, offering_discount_pct)
                                cv_r2 = model_info.get("cv_r2")
                                cv_rmse = model_info.get("cv_rmse")
                                n_train = model_info.get("n_train")
                                if cv_r2 is not None and cv_r2 < 0:
                                    quality = ("⚠⚠ 예측력 없음 — R²<0은 단순 평균값 예측보다도 "
                                               "못하다는 뜻. 이 조정값은 참고용으로도 신뢰하지 말 것")
                                elif cv_r2 is not None and cv_r2 < 0.3:
                                    quality = "⚠ 프로토타입 수준 — 참고용, 의사결정 근거로 쓰지 말 것"
                                else:
                                    quality = "참고용"
                                log(f"    공모할인율(예측, input_limited): {offering_discount_pct:.1f}%"
                                    f"  [모델 n={n_train}, cv_r2={cv_r2}, cv_rmse={cv_rmse}]")
                                log(f"    {quality}")
                                log(f"    ▶ (참고) ML조정 공모시총: {offering_val:,.1f}억원"
                                    f"  ※ 메인 지표는 위 평가시총({val:,.1f}억원)")
                                result["offering_market_cap_억원"] = offering_val
                                result["offering_discount_pct"] = offering_discount_pct
                            except Exception as e:
                                log(f"    [경고] 공모할인율 예측 실패 — {type(e).__name__}: {e}"
                                    f"  (평가시총 {val:,.1f}억원은 그대로 유효)")
    except Exception as e:
        # 실제 core/valuation_psr.py, core/valuation.py 쪽 계산 함수가 예외를
        # 던지는 경우(예: 산출 가능한 피어 0개) 이 케이스만 실패 처리하고
        # 배치 전체는 계속 진행한다. 원인은 warning에 그대로 담아 표에 노출.
        msg = f"{type(e).__name__}: {e}"
        log(f"    [오류] 가치 산출 중 예외 발생 — {msg}")
        result["warning"] = msg

    log(f"\n{'='*65}")
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", required=True)
    ap.add_argument("--latest-round", default=None,
                     help='예: "Seed+Angel", "Series A", "Pre-IPO"')
    ap.add_argument("--ipo-mentioned", action="store_true",
                     help="IR에 IPO 계획이 명시되어 있으면 지정")
    ap.add_argument("--ipo-target-year", type=int, default=None)
    ap.add_argument("--founding-year", type=int, default=None)
    ap.add_argument("--net-income", type=float, default=None, help="발행사 순이익(억원, 트레일링)")
    ap.add_argument("--future-net-income", type=float, default=None, help="발행사 추정 순이익(억원)")
    ap.add_argument("--ni-year", type=int, default=None)
    ap.add_argument("--revenue", type=float, default=None, help="발행사 매출(억원)")
    ap.add_argument("--yoy-growth", type=float, default=0.0)
    ap.add_argument("--decay-factor", type=float, default=0.6)
    ap.add_argument("--psr-method", default="median", choices=["median", "mean"])
    ap.add_argument("--per-method", default="median",
                     choices=["median", "mean", "sim_weighted", "high_only", "high_median"])
    ap.add_argument("--time-discount-rate", type=float, default=0.20,
                     help="tech_special 서브트랙 시간가치 현가할인율 (기본 20%%) — "
                          "ML이 예측하는 공모할인율과는 다른 값이니 혼동 주의")
    ap.add_argument("--discount-years", type=float, default=None,
                     help="명시하면 ni_year 자동계산보다 우선 적용")
    ap.add_argument("--segment", default=None,
                     choices=["ai", "consumer", "deeptech", "fintech", "manufacturing",
                              "materials", "mobility", "platform", "robotics",
                              "semiconductor", "software", "unknown"],
                     help="공모할인율 모델의 업종 세그먼트 (미지정 시 unknown)")
    ap.add_argument("--is-growth-special", action="store_true",
                     help="성장특례 트랙이면 지정")
    ap.add_argument("--skip-offering-discount", action="store_true",
                     help="ML 공모할인율 예측 단계 생략 (평가시총까지만 산출)")
    args = ap.parse_args()

    run_one(args.case, {
        "founding_year": args.founding_year,
        "latest_round": args.latest_round,
        "ipo_mentioned": args.ipo_mentioned,
        "ipo_target_year": args.ipo_target_year,
        "net_income": args.net_income,
        "future_net_income": args.future_net_income,
        "ni_year": args.ni_year,
        "revenue": args.revenue,
        "yoy_growth": args.yoy_growth,
        "decay_factor": args.decay_factor,
        "psr_method": args.psr_method,
        "per_method": args.per_method,
        "time_discount_rate": args.time_discount_rate,
        "discount_years": args.discount_years,
        "segment": args.segment,
        "is_growth_special": args.is_growth_special,
        "skip_offering_discount": args.skip_offering_discount,
    })


if __name__ == "__main__":
    main()