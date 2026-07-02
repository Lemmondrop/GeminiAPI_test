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
from core.valuation_psr import calculate_representative_psr
from core.valuation import representative_per, value_band_from_projections
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

    signals: StageSignals 필드와 동일한 키 + 다년도 추정치(net_income_projections:
    {연도:순이익}, revenue_projections: {연도:매출})를 담은 dict. 구버전 단일값
    (future_net_income+ni_year, revenue)만 와도 1개짜리 밴드로 자동 변환되어
    같은 코드 경로를 탄다. 없는 키는 기본값 사용.
    CLI(main())와 배치 스크립트(run_all_cases.py)가 이 함수 하나를 공유한다
    — 로직이 두 곳에 따로 있으면 예전 'decision' 이중 라우팅 문제가 다시
    생기므로, 계산 로직의 단일 진실 공급원은 이 함수 하나로 고정한다.

    반환값에는 stage_decision, 산출 결과(가능한 경우), 산출 불가 사유를 담아
    배치 스크립트가 표로 요약할 수 있게 한다.
    """
    psr_method = signals.get("psr_method", "median") or "median"
    per_method = signals.get("per_method", "median") or "median"
    skip_offering_discount = bool(signals.get("skip_offering_discount", False))
    current_year = signals.get("current_year") or datetime.now().year

    def log(msg=""):
        if verbose:
            print(msg)

    # JSON(ir_signals.json)에서 읽어온 dict는 키가 항상 문자열이다(JSON 스펙상
    # 정수 키가 없음) — "2026" vs 2026. CLI 경로(_parse_year_dict)는 이미
    # int로 변환해서 넘기지만, run_all_cases.py가 ir_signals.json을 그대로
    # json.loads()해서 넘기는 경로는 변환이 없어서 여기서 문자열 키가 들어오면
    # `year - current_year`에서 TypeError가 난다(실사례로 발견한 버그).
    # 호출 경로가 뭐든 여기 한 곳에서 정규화해 원천 차단한다.
    def _int_keys(d):
        if not d:
            return d
        return {int(k): v for k, v in d.items()}

    signals = dict(signals)  # 원본 dict 변형 방지
    signals["net_income_projections"] = _int_keys(signals.get("net_income_projections"))
    signals["revenue_projections"] = _int_keys(signals.get("revenue_projections"))

    # 과거 연도(현재연도보다 이전) 추정치는 밴드에서 제외한다. IR 작성 시점엔
    # "당해년도"였던 연도라도, 계산을 실행하는 지금 시점 기준으론 이미 지난
    # 연도라 "미래 목표 달성 시 현재가치"라는 이 기능의 취지에 안 맞는다
    # (실사례로 발견: 2025년 9월 작성 IR의 "2025E"가, 2026년에 계산을 돌리니
    # 이미 지난 연도인데도 할인율 0%로 "당해년도"인 것처럼 처리되던 버그).
    def _drop_past_years(proj, label):
        if not proj:
            return proj
        past = {y: v for y, v in proj.items() if y < current_year}
        if past:
            log(f"    (과거 연도 추정치 제외 — {label}: "
                + ", ".join(f"{y}E({v:+.1f})" for y, v in sorted(past.items()))
                + f"  ※ 문서 작성 당시엔 미래였으나 현재({current_year}년) 기준 이미 지난 연도)")
        return {y: v for y, v in proj.items() if y >= current_year} or None

    signals["net_income_projections"] = _drop_past_years(
        signals.get("net_income_projections"), "순이익")
    signals["revenue_projections"] = _drop_past_years(
        signals.get("revenue_projections"), "매출")

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
    # determine_ipo_subtrack()은 net_income/future_net_income 단일값 기준으로
    # general/tech_special을 가른다. 다년도 스키마(net_income_projections)만
    # 주어졌을 때도 판별이 되도록, 여기서 대표값을 유도해 보강한다
    # (stage_router.py 자체는 안 건드리고 입력을 풍부하게 만드는 어댑터 방식).
    sub_net_income = signals.get("net_income")
    sub_future_net_income = signals.get("future_net_income")
    ni_proj_for_subtrack = signals.get("net_income_projections")
    if ni_proj_for_subtrack:
        if sub_net_income is None and current_year in ni_proj_for_subtrack:
            sub_net_income = ni_proj_for_subtrack[current_year]
        if sub_future_net_income is None:
            sub_future_net_income = ni_proj_for_subtrack[max(ni_proj_for_subtrack)]

    stage_decision = determine_company_stage(StageSignals(
        company_name=case,
        founding_year=signals.get("founding_year"),
        latest_round=signals.get("latest_round"),
        ipo_mentioned=bool(signals.get("ipo_mentioned", False)),
        ipo_target_year=signals.get("ipo_target_year"),
        net_income=sub_net_income,
        future_net_income=sub_future_net_income,
        revenue=signals.get("revenue") or (
            max(signals["revenue_projections"].values())
            if signals.get("revenue_projections") else None
        ),
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
            # 새 스키마(revenue_projections: {연도: 매출억원}) 우선 — 플랫폼
            # "추정 재무제표 입력" 화면과 1:1 대응. 구버전(단일 revenue)만
            # 있으면 당해년도 1개짜리 밴드로 변환해 같은 코드 경로를 탄다
            # (yoy_growth 외삽 방식은 폐기 — 실제 연도별 목표치가 있는데
            # 성장률 하나로 추정하는 게 더 약한 방법이었음, 2026-07 결정).
            revenue_projections = signals.get("revenue_projections")
            if not revenue_projections:
                revenue = signals.get("revenue")
                if revenue is not None:
                    revenue_projections = {current_year: revenue}
                    log(f"    (구버전 호환: revenue 단일값 → "
                        f"{{{current_year}: {revenue}}} 단일 밴드로 변환)")

            if not revenue_projections:
                msg = ("next_round_track인데 매출 데이터 없음"
                       "(revenue_projections 또는 revenue 필요) — PSR 산출 불가.")
                log(f"    [경고] {msg}")
                result["warning"] = msg
            else:
                rep_psr = calculate_representative_psr(enriched, method=psr_method)
                n_psr = sum(1 for p in enriched if p.get('psr') is not None)
                log(f"    대표 PSR({psr_method}): {rep_psr:.3f}배  (조인 성공 피어 {n_psr}개 기준)")

                band = value_band_from_projections(revenue_projections, rep_psr, current_year)
                log(f"\n    [연도별 현재가치 밴드 — 매출 목표 달성 시나리오]")
                for year, d in band.items():
                    label = "당해년도" if d["years_out"] <= 0 else f"{d['years_out']}년 후"
                    log(f"      {year}E ({label}, 할인율 {d['discount_rate']*100:.0f}%): "
                        f"매출 {revenue_projections[year]:.1f}억 × PSR {rep_psr:.3f} = "
                        f"평가액 {d['raw_value_억원']:,.1f}억 → 현재가치 {d['present_value_억원']:,.1f}억원")

                low_year, high_year = min(band), max(band)
                low_v = band[low_year]["present_value_억원"]
                high_v = band[high_year]["present_value_억원"]
                avg_v = sum(d["present_value_억원"] for d in band.values()) / len(band)
                log(f"\n    ▶ 추정 기업가치 밴드: {low_v:,.0f}억원({low_year}E 목표 기준)"
                    + (f" ~ {high_v:,.0f}억원({high_year}E 목표 기준)" if high_year != low_year else ""))
                log(f"    ▶ 추정 기업가치({len(band)}개년 평균): {avg_v:,.0f}억원")
                result["value_range_억원"] = (low_v, high_v)
                result["average_value_억원"] = avg_v
                result["value_band"] = band

        elif stage_decision.track == "ipo_track":
            # general: 트레일링 실적이라 당해년도 1개짜리 밴드(할인 없음, years_out=0).
            # tech_special: 새 스키마(net_income_projections: {연도: 순이익억원})
            # 우선 — 구버전(future_net_income+ni_year 단일값)만 있으면 1개짜리
            # 밴드로 변환해 같은 코드 경로를 탄다. general/tech_special이 이제
            # "밴드가 1개 연도냐 여러 연도냐"의 차이일 뿐, 계산 로직은 하나로 통일.
            if stage_decision.sub_track == "general":
                net_income = signals.get("net_income")
                if net_income is None and signals.get("net_income_projections"):
                    net_income = signals["net_income_projections"].get(current_year)
                ni_projections = {current_year: net_income} if net_income is not None else None
            else:
                ni_projections = signals.get("net_income_projections")
                if not ni_projections:
                    fni = signals.get("future_net_income")
                    ni_year = signals.get("ni_year")
                    if fni is not None and ni_year is not None:
                        ni_projections = {ni_year: fni}
                        log(f"    (구버전 호환: future_net_income+ni_year → "
                            f"{{{ni_year}: {fni}}} 단일 밴드로 변환)")

            if not ni_projections:
                msg = "ipo_track이지만 서브트랙에 필요한 순이익 데이터 없음 — PER 산출 불가."
                log(f"    [경고] {msg}")
                result["warning"] = msg
            else:
                # PER은 순이익이 양수일 때만 경제적으로 의미가 있다 — 적자 연도에
                # PER을 곱하면 "마이너스 기업가치"라는 무의미한 숫자가 나온다.
                # tech_special은 정의상 당해년도가 적자일 수 있으므로(그래서
                # 일반트랙이 아니라 기술특례트랙), 밴드에서 적자 연도는 제외하고
                # 어느 연도를 왜 뺐는지 로그로 남긴다.
                excluded_years = {y: v for y, v in ni_projections.items() if v <= 0}
                ni_projections = {y: v for y, v in ni_projections.items() if v > 0}
                if excluded_years:
                    log(f"    적자 연도 제외(PER 적용 무의미): "
                        + ", ".join(f"{y}E({v:+.1f}억)" for y, v in sorted(excluded_years.items())))

            if not ni_projections:
                msg = "순이익 추정치가 전부 적자라 PER 밴드를 산출할 연도가 없음."
                log(f"    [경고] {msg}")
                result["warning"] = msg
            else:
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
                    log(f"    대표 PER({per_method}): {rep_per:.2f}배  (n={len(per_values)})")

                    band = value_band_from_projections(ni_projections, rep_per, current_year)
                    log(f"\n    [연도별 현재가치 밴드 — 순이익 목표 달성 시나리오]")
                    for year, d in band.items():
                        label = "당해년도" if d["years_out"] <= 0 else f"{d['years_out']}년 후"
                        log(f"      {year}E ({label}, 할인율 {d['discount_rate']*100:.0f}%): "
                            f"순이익 {ni_projections[year]:.1f}억 × PER {rep_per:.2f} = "
                            f"평가시총 {d['raw_value_억원']:,.1f}억 → 현재가치 {d['present_value_억원']:,.1f}억원")

                    # 헤드라인: ipo_target_year와 맞는 연도 우선, 없으면 밴드 중 가장 먼 연도
                    headline_year = signals.get("ipo_target_year")
                    if headline_year not in band:
                        headline_year = max(band.keys())
                    val = band[headline_year]["present_value_억원"]
                    applied_ni = val / rep_per if rep_per else 0.0

                    # 밴드 연도들의 현재가치 단순평균 — 헤드라인(목표연도 단일값)을
                    # 대체하는 게 아니라 "여러 시나리오를 종합한 참고치"로 추가 제공.
                    avg_val = sum(d["present_value_억원"] for d in band.values()) / len(band)

                    log(f"\n    ▶ 평가시총(헤드라인, {headline_year}E 기준): {val:,.1f}억원")
                    log(f"    ▶ 평가시총({len(band)}개년 평균): {avg_val:,.1f}억원")
                    result["equity_value_억원"] = val
                    result["average_value_억원"] = avg_val
                    result["value_band"] = band
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
    ap.add_argument("--net-income", type=float, default=None, help="발행사 순이익(억원, 트레일링/general 전용)")
    ap.add_argument("--net-income-projections", type=str, default=None,
                     help='다년도 추정 순이익 JSON, 예: \'{"2027": 4.1, "2028": 8.5}\' (tech_special용, 신규 스키마) '
                          '— PowerShell 인용 문제 잦으면 --net-income-projections-file 사용 권장')
    ap.add_argument("--net-income-projections-file", type=str, default=None,
                     help='다년도 추정 순이익 JSON을 담은 파일 경로 (셸 따옴표 문제 우회용)')
    ap.add_argument("--future-net-income", type=float, default=None,
                     help="[구버전 호환] 단일 연도 추정 순이익 — --net-income-projections 권장")
    ap.add_argument("--ni-year", type=int, default=None, help="[구버전 호환] --future-net-income의 기준연도")
    ap.add_argument("--revenue", type=float, default=None,
                     help="[구버전 호환] 단일 연도 매출 — --revenue-projections 권장")
    ap.add_argument("--revenue-projections", type=str, default=None,
                     help='다년도 추정 매출 JSON, 예: \'{"2027": 32.5, "2028": 54.0}\' (신규 스키마) '
                          '— PowerShell 인용 문제 잦으면 --revenue-projections-file 사용 권장')
    ap.add_argument("--revenue-projections-file", type=str, default=None,
                     help='다년도 추정 매출 JSON을 담은 파일 경로 (셸 따옴표 문제 우회용)')
    ap.add_argument("--yoy-growth", type=float, default=0.0,
                     help="[더 이상 계산에 안 쓰임 — revenue_projections로 대체됨, 하위호환용 인자만 유지]")
    ap.add_argument("--decay-factor", type=float, default=0.6,
                     help="[더 이상 계산에 안 쓰임 — revenue_projections로 대체됨, 하위호환용 인자만 유지]")
    ap.add_argument("--psr-method", default="median", choices=["median", "mean"])
    ap.add_argument("--per-method", default="median",
                     choices=["median", "mean", "sim_weighted", "high_only", "high_median"])
    ap.add_argument("--time-discount-rate", type=float, default=0.20,
                     help="[더 이상 계산에 안 쓰임 — 연차별 할인율 스케줄(15%%/20%%/30%%)로 대체됨]")
    ap.add_argument("--discount-years", type=float, default=None,
                     help="[더 이상 계산에 안 쓰임]")
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

    def _parse_year_dict(raw: str | None, file_path: str | None) -> dict | None:
        """CLI 문자열 또는 파일 경로에서 {연도: 값} dict 파싱.
        둘 다 주어지면 파일을 우선한다 — 셸 따옴표 문제로 raw가 깨져 있을
        가능성이 파일보다 높기 때문."""
        text = None
        if file_path:
            p = Path(file_path)
            if not p.exists():
                sys.exit(f"[ERROR] 파일을 찾을 수 없습니다: {p}")
            # utf-8-sig: BOM이 있으면 자동으로 벗기고, 없어도 그대로 읽는다.
            # PowerShell의 Out-File -Encoding utf8은 기본적으로 BOM을 붙이므로
            # 이걸 안 하면 JSONDecodeError(Unexpected UTF-8 BOM)가 난다.
            text = p.read_text(encoding="utf-8-sig")
        elif raw:
            text = raw
        if not text:
            return None
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError as e:
            sys.exit(
                f"[ERROR] JSON 파싱 실패: {e}\n"
                f"  받은 값: {text!r}\n"
                f"  PowerShell에서 따옴표가 깨지는 문제라면 --net-income-projections-file "
                f"또는 --revenue-projections-file로 파일 경로를 넘겨보세요."
            )
        return {int(k): float(v) for k, v in parsed.items()}

    run_one(args.case, {
        "founding_year": args.founding_year,
        "latest_round": args.latest_round,
        "ipo_mentioned": args.ipo_mentioned,
        "ipo_target_year": args.ipo_target_year,
        "net_income": args.net_income,
        "net_income_projections": _parse_year_dict(
            args.net_income_projections, args.net_income_projections_file),
        "future_net_income": args.future_net_income,
        "ni_year": args.ni_year,
        "revenue": args.revenue,
        "revenue_projections": _parse_year_dict(
            args.revenue_projections, args.revenue_projections_file),
        "psr_method": args.psr_method,
        "per_method": args.per_method,
        "segment": args.segment,
        "is_growth_special": args.is_growth_special,
        "skip_offering_discount": args.skip_offering_discount,
    })


if __name__ == "__main__":
    main()