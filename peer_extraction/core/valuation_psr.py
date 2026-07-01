"""valuation.py 확장 — PSR(매출배수) 기반 밸류에이션 + 자동 판별 로직.

배경:
  기존 calculate_equity_value()는 PER × 순이익 구조라서,
  '현재 적자, 그러나 신뢰 가능한 미래 순이익 추정치 보유'(예: 센스톤,
  future_net_income_억원 + 현가할인) 케이스는 이미 커버한다.

  그러나 Seed~Series A 단계 기업 중 다음 두 부류는 기존 구조로 처리 불가:
    A) 매출은 있으나 순이익 추정이 신뢰 불가 (예: 스쿼드엑스)
       → PER 적용 불가. PSR(매출배수) 필요.
    B) 매출조차 미미/전무 (예: 뉴클럭스, 설립 1년 미만)
       → PER도 PSR도 무의미. 본 파이프라인 범위 밖
         (동단계 비상장 거래 데이터 — The VC 등 — 별도 모듈 필요,
         Module 5 Exit Probability와 연계 검토 대상).

  본 파일은 위 A) 케이스를 처리하는 PSR 계산 함수와,
  세 갈래(PER 적용 / PSR 적용 / 데이터부족) 중 어느 경로를 탈지
  입력값만으로 자동 판별하는 라우터를 제공한다.

기존 valuation.py와의 통합:
  determine_valuation_method() → calculate_representative_per()
                                  또는 calculate_representative_psr()
                                  중 선택해 호출하는 식으로 사용.
  기존 함수(calculate_representative_per, calculate_equity_value,
  calculate_valuation_band)는 변경하지 않음 — 순수 추가.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

ValuationMethod = Literal["PER", "PSR", "insufficient_data"]


# ══════════════════════════════════════════════════════════
# 0. 단위 변환 헬퍼 — UNIVERSE_FIN_CSV는 원(₩) 단위로 저장됨
# ══════════════════════════════════════════════════════════
# config.UNIVERSE_FIN_CSV(market_cap, revenue_annual, revenue_ltm,
# net_income_annual, net_income_ltm 등)는 전부 '원' 단위로 저장되어
# 있음(실측 확인: 엔비알모션 market_cap=92,703,810,000원).
# 본 모듈의 모든 함수는 '억원' 단위를 전제로 설계되었으므로,
# CSV에서 읽은 raw 값은 반드시 아래 변환을 거쳐야 한다.
#
#   잘못된 사용 예 (1억 배 과대 산출 위험):
#     calculate_equity_value_via_psr(revenue_억원=row["revenue_annual"], ...)
#   올바른 사용:
#     calculate_equity_value_via_psr(
#         revenue_억원=won_to_억원(row["revenue_annual"]), ...)
def won_to_억원(value_won: float) -> float:
    """원(₩) 단위 값을 억원 단위로 변환. None 입력 시 None 반환."""
    if value_won is None:
        return None
    return value_won / 1e8


def peer_row_to_psr_dict(row: dict) -> dict:
    """UNIVERSE_FIN_CSV의 1행(dict)을 calculate_representative_psr()이
    기대하는 형식({market_cap_억원, revenue_억원, ...})으로 변환.

    row 예상 키: market_cap, revenue_annual(또는 revenue_ltm) — 둘 다 원 단위.
    revenue_ltm이 있으면 우선 사용(최신 실적 반영), 없으면 revenue_annual.
    """
    revenue_won = row.get("revenue_ltm") or row.get("revenue_annual")
    return {
        "market_cap_억원": won_to_억원(row.get("market_cap")),
        "revenue_억원":    won_to_억원(revenue_won),
        "similarity_score": row.get("similarity_score", 0.0),
        "similarity_tier":  row.get("similarity_tier", ""),
    }



@dataclass
class MethodDecision:
    method: ValuationMethod
    reason: str


def determine_valuation_method(
    net_income_억원: float | None,
    future_net_income_억원: float | None = None,
    ni_year: int | None = None,
    revenue_억원: float | None = None,
    min_revenue_억원: float = 1.0,   # 이 미만이면 '매출 미미'로 간주
) -> MethodDecision:
    """입력 재무 데이터만으로 PER / PSR / 데이터부족 중 적용 경로를 결정.

    우선순위:
      1) 현재 순이익이 흑자 → PER (할인 불필요, 가장 신뢰도 높음)
      2) 현재 적자이나 신뢰 가능한 미래 순이익 추정치 보유
         (future_net_income_억원 + ni_year 둘 다 제공) → PER (현가할인)
      3) 미래 순이익 추정 없음(또는 신뢰 불가) + 매출 존재 → PSR
      4) 매출조차 min_revenue_억원 미만 → insufficient_data
         (본 파이프라인 범위 밖, 동단계 비상장 거래 데이터 필요)
    """
    if net_income_억원 is not None and net_income_억원 > 0:
        return MethodDecision(
            "PER", "현재 순이익 흑자 — 할인 없이 직접 PER 적용 가능")

    if (future_net_income_억원 is not None and future_net_income_억원 > 0
            and ni_year is not None):
        return MethodDecision(
            "PER",
            f"현재 적자이나 {ni_year}년 추정 순이익({future_net_income_억원}억원) "
            "보유 — 현가할인 PER 적용")

    if revenue_억원 is not None and revenue_억원 >= min_revenue_억원:
        return MethodDecision(
            "PSR",
            f"순이익 추정 불가(미래 추정치 미제공) + 매출 {revenue_억원}억원 "
            "보유 — PSR(매출배수) 적용")

    return MethodDecision(
        "insufficient_data",
        "순이익 추정치도 유의미한 매출도 없음 — 본 파이프라인(PER/PSR) "
        "적용 불가. 동단계 비상장 거래 데이터(The VC 등) 기반 별도 "
        "평가 필요 (Module 5 Exit Probability 연계 검토 대상)"
    )


# ══════════════════════════════════════════════════════════
# 2. PSR 대표값 산출 — calculate_representative_per()의 매출배수 버전
# ══════════════════════════════════════════════════════════
def calculate_representative_psr(
    peers: list[dict],
    method: str = "median",
) -> float:
    """피어 리스트에서 대표 PSR(시총/매출 배수)을 계산한다.

    peers의 각 원소는 'market_cap_억원'과 'revenue_억원' 키를 가져야 한다.
    revenue_억원이 없거나 0 이하인 피어는 자동 제외된다(매출 없는 피어로
    배수를 계산하면 분모가 0이 되어 무의미).

    method: "median" | "sim_weighted" | "high_only" | "mean"
            (calculate_representative_per()과 동일 시맨틱)
    """
    valid = [
        p for p in peers
        if p.get("revenue_억원") not in (None, 0)
        and p.get("market_cap_억원") not in (None, 0)
    ]
    if not valid:
        raise ValueError(
            "PSR 산출 가능한 피어가 0개입니다. peers 원소에 "
            "'revenue_억원' 키가 있는지 확인하세요 — 현재 파이프라인의 "
            "PeerItem에 매출 필드가 없다면, 재무 수집 단계(DART "
            "fnlttSinglAcntAll)에서 매출액(account_nm='매출액' 또는 "
            "'수익(매출액)')도 함께 저장하도록 확장이 필요합니다."
        )

    # ── 단위 실수 감지 가드 ──────────────────────────────────
    # UNIVERSE_FIN_CSV는 원(₩) 단위 — 변환 없이 그대로 넘기면
    # market_cap_억원에 '927억원'이 아니라 '927조원'스러운 raw 값이
    # 들어와 즉시 비정상적으로 큰 수치가 됨. 삼성전자급(약 500만억원)을
    # 훨씬 상회하는 1000만억원을 절대 상한으로 두어 조기 차단한다.
    SANITY_MAX_억원 = 10_000_000  # = 1경원, 사실상 도달 불가능한 상한
    for p in valid:
        if p["market_cap_억원"] > SANITY_MAX_억원 or p["revenue_억원"] > SANITY_MAX_억원:
            raise ValueError(
                f"비정상적으로 큰 값 감지: market_cap_억원={p['market_cap_억원']:,.0f}, "
                f"revenue_억원={p['revenue_억원']:,.0f}. "
                "단위 변환(원→억원, won_to_억원()) 누락 가능성이 높습니다. "
                "UNIVERSE_FIN_CSV에서 직접 읽었다면 peer_row_to_psr_dict()를 "
                "거쳤는지 확인하세요."
            )

    psrs = [p["market_cap_억원"] / p["revenue_억원"] for p in valid]

    if method == "median":
        psrs_sorted = sorted(psrs)
        n = len(psrs_sorted)
        mid = n // 2
        return (psrs_sorted[mid] if n % 2 == 1
                else (psrs_sorted[mid - 1] + psrs_sorted[mid]) / 2)

    if method == "mean":
        return sum(psrs) / len(psrs)

    if method == "high_only":
        high = [p for p in valid if p.get("similarity_tier") == "높음"]
        pool = high if high else valid
        vals = [p["market_cap_억원"] / p["revenue_억원"] for p in pool]
        return sum(vals) / len(vals)

    if method == "sim_weighted":
        weights = [p.get("similarity_score", 0.0) for p in valid]
        total_w = sum(weights)
        if total_w <= 0:
            return sum(psrs) / len(psrs)   # 가중치 전무 시 단순평균 폴백
        return sum(psr * w for psr, w in zip(psrs, weights)) / total_w

    raise ValueError(f"알 수 없는 method: {method!r}")


# ══════════════════════════════════════════════════════════
# 3. PSR 기반 기업가치 산출 — calculate_equity_value()의 매출배수 버전
# ══════════════════════════════════════════════════════════
def calculate_equity_value_via_psr(
    revenue_억원: float,
    representative_psr: float,
    growth_adjustment: float | None = None,
) -> dict:
    """매출액과 대표 PSR을 곱해 기업가치를 산출한다.

    growth_adjustment: 선택적 성장률 보정 (예: YoY 성장률이 매우 높은
    스쿼드엑스(YoY 317%) 같은 케이스는 trailing 매출보다 forward 매출이
    더 적합할 수 있음. 0.2 = 향후 매출 20% 추가 성장 가정 시
    revenue_억원 * (1 + 0.2)로 보정. None이면 trailing 매출 그대로 사용
    (가장 보수적 — 기본값).
    """
    if revenue_억원 <= 0:
        raise ValueError("revenue_억원은 0보다 커야 합니다 (PSR은 매출 기반)")
    if representative_psr <= 0:
        raise ValueError("representative_psr은 0보다 커야 합니다")

    SANITY_MAX_억원 = 10_000_000
    if revenue_억원 > SANITY_MAX_억원:
        raise ValueError(
            f"revenue_억원={revenue_억원:,.0f} 비정상적으로 큽니다. "
            "원 단위 값을 억원으로 변환(won_to_억원()) 했는지 확인하세요."
        )

    applied_revenue = revenue_억원
    if growth_adjustment is not None:
        applied_revenue = revenue_억원 * (1 + growth_adjustment)

    equity_value = applied_revenue * representative_psr

    return {
        "applied_revenue_억원": round(applied_revenue, 1),
        "representative_psr": round(representative_psr, 2),
        "equity_value_억원": round(equity_value, 1),
        "growth_adjustment_applied": growth_adjustment is not None,
        "valuation_method": (
            f"PSR 기반: {applied_revenue:.1f}억원(매출"
            + (f", {growth_adjustment:+.0%} 성장보정" if growth_adjustment else "")
            + f") × {representative_psr:.2f}배(대표 PSR)"
        ),
    }

# ══════════════════════════════════════════════════════════
# 4. PSR 기반 밴드 산출 — trailing(보수) / decayed(기준) / naive(공격)
# ══════════════════════════════════════════════════════════
def calculate_equity_value_psr_band(
    revenue_억원: float,
    representative_psr: float,
    yoy_growth_rate: float | None = None,
    decay_factor: float = 0.5,
) -> dict:
    """PSR 기반 기업가치를 3개 시나리오로 산출해 밴드를 만든다.

    단일 trailing 매출 기준 PSR만 적용하면 고성장 기업의 미래 가치를
    과소평가한다(예: YoY 317% 성장 기업에 성장 미반영 PSR을 그대로
    적용 시 결과가 절반 이하로 나옴). 그렇다고 현재 성장률을 그대로
    1년 외삽하는 것도 비현실적(하이퍼그로스는 통상 지속되지 않음).
    셋을 함께 제시해 범위로 판단하게 한다.

    시나리오:
      conservative_trailing : 성장 미반영. 가장 보수적인 하한.
      base_decayed          : 성장률에 decay_factor(기본 0.5, 즉 "성장률
                               절반으로 둔화") 적용. 가장 현실적인 중심값.
                               (근거: 하이퍼그로스(YoY 200%+) 스타트업은
                               매출 기반이 커질수록 분모효과로 다음 해
                               성장률이 대략 절반 수준으로 꺾이는 경향이
                               업계에서 경험적으로 자주 관찰됨 — 절대
                               법칙은 아니므로 decay_factor는 업종·단계별
                               조정 가능하게 파라미터화함.)
      aggressive_naive      : 성장률 둔화 없이 그대로 외삽. 가장 공격적인
                               상한 — 실제로는 거의 도달하지 않는
                               시나리오임을 사용 시 명확히 인지할 것.

    yoy_growth_rate: 최근 YoY 매출 성장률 (예: 3.17 = +317%). None이거나
                      0 이하면 성장 시나리오를 생략하고 trailing만 반환.
    decay_factor: 성장률 둔화 계수. 1.0=둔화 없음(aggressive와 동일),
                  0.5=기본값(절반 둔화), 0.0=즉시 성장 정지(매우 보수적).
    """
    scenarios = {
        "conservative_trailing": calculate_equity_value_via_psr(
            revenue_억원, representative_psr, growth_adjustment=None),
    }

    if yoy_growth_rate is not None and yoy_growth_rate > 0:
        decayed_rate = yoy_growth_rate * decay_factor
        base = calculate_equity_value_via_psr(
            revenue_억원, representative_psr, growth_adjustment=decayed_rate)
        base["decay_factor_applied"] = decay_factor
        base["decayed_growth_rate"] = round(decayed_rate, 3)
        scenarios["base_decayed"] = base

        aggressive = calculate_equity_value_via_psr(
            revenue_억원, representative_psr, growth_adjustment=yoy_growth_rate)
        scenarios["aggressive_naive"] = aggressive

    low = scenarios["conservative_trailing"]["equity_value_억원"]
    high_scenario = scenarios.get("aggressive_naive", scenarios["conservative_trailing"])
    high = high_scenario["equity_value_억원"]
    mid = scenarios.get("base_decayed", {}).get("equity_value_억원")

    return {
        "low_억원": low,
        "mid_억원": mid,
        "high_억원": high,
        "representative_psr": round(representative_psr, 3),
        "revenue_억원": revenue_억원,
        "yoy_growth_rate": yoy_growth_rate,
        "decay_factor": decay_factor if yoy_growth_rate else None,
        "scenarios": scenarios,
    }



if __name__ == "__main__":
    print("=== 판별 로직 테스트 (5개 PoC 케이스) ===\n")

    test_cases = [
        # (회사명, net_income, future_ni, ni_year, revenue)
        ("센스톤",       -32.5, 152.3, 2027, None),
        ("피엔피",        None, None,  None, None),   # 재무 데이터 미수집 상태
        ("케이더블유씨",   None, None,  None, None),
        ("뉴클럭스",       None, None,  None, 0.0),    # 설립 1년 미만, 매출 거의 없음 가정
        ("스쿼드엑스",     None, None,  None, 46.3),   # IR 기재 매출
    ]

    for name, ni, fni, year, rev in test_cases:
        decision = determine_valuation_method(
            net_income_억원=ni, future_net_income_억원=fni,
            ni_year=year, revenue_억원=rev,
        )
        print(f"  {name:12s} → {decision.method:18s} | {decision.reason}")

    print("\n=== PSR 계산 시뮬레이션 (스쿼드엑스, 가상 피어 데이터) ===\n")
    # 실제로는 Module 1 peers에 revenue_억원이 없으면 이 단계에서 ValueError 발생
    # (의도된 동작 — 데이터 누락을 조용히 넘기지 않음). 여기선 시연을 위해
    # 가상의 피어 재무를 주입.
    demo_peers = [
        {"corp_name": "아티스트컴퍼니", "market_cap_억원": 1800, "revenue_억원": 530,
         "similarity_score": 0.65, "similarity_tier": "중간"},
        {"corp_name": "브리지텍",       "market_cap_억원": 950,  "revenue_억원": 410,
         "similarity_score": 0.45, "similarity_tier": "중간"},
    ]
    rep_psr = calculate_representative_psr(demo_peers, method="sim_weighted")
    print(f"  대표 PSR(sim_weighted): {rep_psr:.2f}배")

    val = calculate_equity_value_via_psr(
        revenue_억원=46.3, representative_psr=rep_psr, growth_adjustment=None)
    print(f"  → {val['valuation_method']}")
    print(f"  → 기업가치: {val['equity_value_억원']}억원  (trailing 매출 기준, 보수적)")

    val_growth = calculate_equity_value_via_psr(
        revenue_억원=46.3, representative_psr=rep_psr, growth_adjustment=3.17)
    print(f"\n  [참고] YoY 317% 성장 반영 시(forward 추정):")
    print(f"  → {val_growth['valuation_method']}")
    print(f"  → 기업가치: {val_growth['equity_value_억원']}억원  "
          f"(공격적 — 단순 시연용, 실사용 시 신중한 보정 필요)")