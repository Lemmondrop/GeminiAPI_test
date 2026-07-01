"""
Growth Index Calculator - V-Wheel (브이휠)
==========================================
스타트업의 성장 지수(0~100)를 산출하는 핵심 모델.

[구조]
- 카테고리 1~5: Rule-Based 가중합 (90점)
- 카테고리 6 (핵심인력): Gemini AI 평가 (10점)
- Gemini: 전체 Growth Index 기반 보완 조언 생성

[카테고리 배점]
1. 매출 성장성      25점
2. 재무 안정성      20점
3. 런웨이           20점
4. 투자 현황        15점
5. 예상가치         10점
6. 핵심인력 (AI)    10점
"""

import os
import json
from dataclasses import dataclass, field
from typing import Optional
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()

# ─────────────────────────────────────────────
# 1. 입력 데이터 구조 정의
# ─────────────────────────────────────────────

@dataclass
class FinancialData:
    """재무 데이터 (재무제표 기반)"""
    revenue_current_year: float          # 당해 매출 (만원)
    revenue_previous_year: float         # 전년 매출 (만원). 창업 1년차면 0
    operating_profit_rate: float         # 영업이익률 (%). 예: -15.0, 5.2
    debt_ratio: float                    # 부채비율 (%). 예: 150.0
    current_ratio: float                 # 유동비율 (%). 예: 120.0
    estimated_revenue_target: float = 0  # 추정 매출 목표 (만원). 없으면 0


@dataclass
class RunwayData:
    """런웨이(번아웃 기간) 데이터"""
    cash_balance: float          # 현재 현금 잔고 (만원)
    monthly_burn_rate: float     # 월 순지출 (만원). 0이면 런웨이 무한대로 처리


@dataclass
class InvestmentData:
    """투자 현황 데이터"""
    investment_stage: str                    # "미투자" / "시드" / "Pre-A" / "시리즈A" / "시리즈B" / "시리즈C이상"
    total_investment_amount: float           # 누적 투자금 (만원)
    months_since_last_investment: int        # 마지막 투자 후 경과 개월 수. 미투자면 0


@dataclass
class ValuationData:
    """예상가치 데이터"""
    current_valuation: float        # 현재 밸류에이션 (만원)
    industry_avg_per: float         # 업종 평균 PER. 없으면 0 (Gemini 추정 사용)
    net_profit: float               # 순이익 (만원). 적자면 음수


@dataclass
class KeyPersonData:
    """핵심인력 개인 데이터"""
    role: str                 # 직책. 예: "CEO", "CTO", "CFO", "COO"
    name: str                 # 이름 (선택)
    education: list[str]      # 학력 목록. 예: ["KAIST 전산학과 석사", "연세대 경영학과 학사"]
    career: list[str]         # 경력 목록. 예: ["카카오 ML팀 5년", "스타트업 창업 2회"]


@dataclass
class KeyPeopleData:
    """핵심인력 전체"""
    solution_description: str           # 기업이 개발하는 솔루션 설명 (1~3문장)
    members: list[KeyPersonData]        # 핵심인력 리스트


@dataclass
class StartupInput:
    """Growth Index 계산을 위한 전체 입력"""
    company_name: str
    industry: str                        # 업종. 예: "핀테크", "헬스케어", "SaaS", "제조"
    founded_year: int                    # 창업 연도
    financial: FinancialData
    runway: RunwayData
    investment: InvestmentData
    valuation: ValuationData
    key_people: KeyPeopleData


# ─────────────────────────────────────────────
# 2. 카테고리별 점수 결과 구조
# ─────────────────────────────────────────────

@dataclass
class CategoryScore:
    name: str
    score: float
    max_score: float
    breakdown: dict = field(default_factory=dict)   # 세부 항목별 점수
    note: str = ""                                   # 특이사항


@dataclass
class GrowthIndexResult:
    company_name: str
    total_score: float                       # 최종 Growth Index (0~100)
    grade: str                               # S / A / B / C / D
    categories: list[CategoryScore]
    ai_advice: str                           # Gemini 보완 조언
    ai_personnel_reason: str                 # Gemini 핵심인력 평가 근거
    calculation_notes: list[str]             # 계산 중 특이사항


# ─────────────────────────────────────────────
# 3. Rule-Based 채점 함수
# ─────────────────────────────────────────────

def score_revenue_growth(data: StartupInput) -> CategoryScore:
    """
    매출 성장성 (25점)
    - YoY 매출 성장률 기반 채점
    - 창업 1년차(전년 매출 0)는 절대 매출 기준으로 보정
    """
    MAX = 25
    notes = []
    breakdown = {}

    prev = data.financial.revenue_previous_year
    curr = data.financial.revenue_current_year

    # 창업 1년차 예외 처리
    if prev == 0:
        # 절대 매출로 점수 산정 (억원 기준)
        curr_eok = curr / 10000  # 만원 → 억원
        if curr_eok >= 10:
            growth_score = 25
        elif curr_eok >= 5:
            growth_score = 20
        elif curr_eok >= 1:
            growth_score = 14
        elif curr_eok >= 0.5:
            growth_score = 8
        else:
            growth_score = 3
        notes.append(f"창업 1년차 → 절대 매출 기준 적용 ({curr_eok:.1f}억원)")
        growth_rate_str = "N/A (창업 1년차)"
    else:
        growth_rate = ((curr - prev) / prev) * 100
        growth_rate_str = f"{growth_rate:.1f}%"

        if growth_rate >= 100:
            growth_score = 25
        elif growth_rate >= 50:
            growth_score = 21
        elif growth_rate >= 30:
            growth_score = 17
        elif growth_rate >= 10:
            growth_score = 11
        elif growth_rate >= 0:
            growth_score = 5
        else:
            growth_score = 0
            notes.append("매출 역성장 감지")

    breakdown["YoY 매출 성장률"] = {"value": growth_rate_str, "score": growth_score}

    # 추정 매출 달성률 보조 점수 (최대 +2, 초과 시 기본 점수에서 차감 없음)
    bonus = 0
    if data.financial.estimated_revenue_target > 0 and curr > 0:
        achievement = (curr / data.financial.estimated_revenue_target) * 100
        if achievement >= 100:
            bonus = 2
        elif achievement >= 80:
            bonus = 1
        breakdown["추정 매출 달성률"] = {"value": f"{achievement:.1f}%", "score": bonus}
        notes.append(f"추정 매출 달성률 {achievement:.1f}% → 보너스 +{bonus}점")

    final_score = min(MAX, growth_score + bonus)
    return CategoryScore(
        name="매출 성장성",
        score=final_score,
        max_score=MAX,
        breakdown=breakdown,
        note=" | ".join(notes)
    )


def score_financial_stability(data: StartupInput) -> CategoryScore:
    """
    재무 안정성 (20점)
    - 영업이익률 10점
    - 부채비율   5점
    - 유동비율   5점
    """
    MAX = 20
    notes = []
    breakdown = {}

    # 영업이익률 (10점)
    opr = data.financial.operating_profit_rate
    if opr >= 15:
        op_score = 10
    elif opr >= 5:
        op_score = 8
    elif opr >= 0:
        op_score = 5
    elif opr >= -20:
        op_score = 2
    elif opr >= -50:
        op_score = 1
    else:
        op_score = 0
        notes.append("심각한 적자 상태")
    breakdown["영업이익률"] = {"value": f"{opr:.1f}%", "score": op_score}

    # 부채비율 (5점) — 낮을수록 안정
    dr = data.financial.debt_ratio
    if dr <= 50:
        debt_score = 5
    elif dr <= 100:
        debt_score = 4
    elif dr <= 200:
        debt_score = 2
    elif dr <= 400:
        debt_score = 1
    else:
        debt_score = 0
        notes.append("부채비율 위험 수준")
    breakdown["부채비율"] = {"value": f"{dr:.1f}%", "score": debt_score}

    # 유동비율 (5점) — 높을수록 단기 안전
    cr = data.financial.current_ratio
    if cr >= 200:
        cur_score = 5
    elif cr >= 150:
        cur_score = 4
    elif cr >= 100:
        cur_score = 3
    elif cr >= 50:
        cur_score = 1
    else:
        cur_score = 0
        notes.append("유동성 부족 경고")
    breakdown["유동비율"] = {"value": f"{cr:.1f}%", "score": cur_score}

    final_score = op_score + debt_score + cur_score
    return CategoryScore(
        name="재무 안정성",
        score=final_score,
        max_score=MAX,
        breakdown=breakdown,
        note=" | ".join(notes)
    )


def score_runway(data: StartupInput) -> CategoryScore:
    """
    런웨이 / 번아웃 기간 (20점)
    - 현금 잔고 ÷ 월 순지출 = 남은 개월 수
    - 월 순지출이 0(흑자)이면 만점 처리
    """
    MAX = 20
    notes = []
    breakdown = {}

    burn = data.runway.monthly_burn_rate
    cash = data.runway.cash_balance

    if burn <= 0:
        # 흑자 or 지출 없음 → 런웨이 무한대
        runway_months = 999
        notes.append("월 순지출 ≤ 0 → 흑자 운영 (런웨이 만점)")
    else:
        runway_months = cash / burn

    breakdown["현금 잔고"] = {"value": f"{cash:,.0f}만원", "score": None}
    breakdown["월 순지출"] = {"value": f"{burn:,.0f}만원", "score": None}
    breakdown["예상 런웨이"] = {
        "value": f"{runway_months:.0f}개월" if runway_months < 999 else "흑자",
        "score": None
    }

    if runway_months >= 24:
        rw_score = 20
    elif runway_months >= 18:
        rw_score = 16
    elif runway_months >= 12:
        rw_score = 11
    elif runway_months >= 6:
        rw_score = 6
    elif runway_months >= 3:
        rw_score = 2
    else:
        rw_score = 0
        notes.append("🚨 런웨이 3개월 미만 — 긴급 자금 조달 필요")

    breakdown["런웨이 점수"] = {"value": f"{runway_months:.0f}개월", "score": rw_score}

    return CategoryScore(
        name="런웨이 (번아웃 기간)",
        score=rw_score,
        max_score=MAX,
        breakdown=breakdown,
        note=" | ".join(notes)
    )


def score_investment(data: StartupInput) -> CategoryScore:
    """
    투자 현황 (15점)
    - 투자 라운드 기반 기본 점수
    - 마지막 투자 후 18개월 초과 시 패널티 -2점
    """
    MAX = 15
    notes = []
    breakdown = {}

    stage_map = {
        "미투자":    2,
        "시드":      5,
        "Pre-A":     8,
        "시리즈A":   11,
        "시리즈B":   14,
        "시리즈C이상": 15,
    }

    stage = data.investment.investment_stage
    base_score = stage_map.get(stage, 2)
    breakdown["투자 라운드"] = {"value": stage, "score": base_score}

    # 마지막 투자 후 경과 패널티
    months_passed = data.investment.months_since_last_investment
    penalty = 0
    if stage != "미투자" and months_passed > 18:
        penalty = 2
        notes.append(f"마지막 투자 후 {months_passed}개월 경과 → -2점 패널티")
    breakdown["마지막 투자 경과"] = {"value": f"{months_passed}개월", "score": -penalty}

    # 누적 투자금 보너스 (최대 +1)
    bonus = 0
    total_inv_eok = data.investment.total_investment_amount / 10000
    if total_inv_eok >= 50:
        bonus = 1
        notes.append(f"누적 투자금 {total_inv_eok:.0f}억원 → +1점 보너스")
    breakdown["누적 투자금"] = {"value": f"{total_inv_eok:.1f}억원", "score": bonus}

    final_score = max(0, min(MAX, base_score - penalty + bonus))
    return CategoryScore(
        name="투자 현황",
        score=final_score,
        max_score=MAX,
        breakdown=breakdown,
        note=" | ".join(notes)
    )


def score_valuation(data: StartupInput) -> CategoryScore:
    """
    예상가치 (10점)
    - 현재 밸류에이션 vs 업종 평균 PER 기반 적정가치 비교
    - 업종 평균 PER이 0이면 밸류에이션 절대 규모로 대체 채점
    """
    MAX = 10
    notes = []
    breakdown = {}

    val = data.valuation.current_valuation       # 만원
    per = data.valuation.industry_avg_per
    net_profit = data.valuation.net_profit       # 만원

    if per > 0 and net_profit > 0:
        # 적정가치 = 순이익 × 업종 평균 PER
        fair_value = net_profit * per
        ratio = val / fair_value  # 1.0 = 정확히 적정가치
        breakdown["업종 평균 PER"] = {"value": f"{per:.1f}x", "score": None}
        breakdown["PER 기반 적정가치"] = {"value": f"{fair_value:,.0f}만원", "score": None}
        breakdown["밸류/적정가치 비율"] = {"value": f"{ratio:.2f}x", "score": None}

        # 밸류가 적정가치에 가까울수록 고점수
        # 너무 낮으면 저평가(성장 여지), 너무 높으면 고평가(리스크)
        if 0.8 <= ratio <= 1.3:
            val_score = 10   # 적정 구간
        elif 0.5 <= ratio < 0.8:
            val_score = 8    # 약간 저평가 (투자 매력)
        elif 1.3 < ratio <= 2.0:
            val_score = 6    # 약간 고평가
        elif ratio < 0.5:
            val_score = 5    # 심한 저평가 → 투자 유치에 불리
        else:
            val_score = 3    # 심한 고평가
            notes.append("밸류에이션 고평가 주의")

    elif per > 0 and net_profit <= 0:
        # 적자 기업 → PER 계산 불가, 밸류 절대 규모 기준
        notes.append("적자 기업 → PER 계산 불가, 밸류 절대 규모 기준 적용")
        val_eok = val / 10000
        if val_eok >= 500:
            val_score = 8
        elif val_eok >= 100:
            val_score = 6
        elif val_eok >= 30:
            val_score = 4
        else:
            val_score = 2
        breakdown["현재 밸류에이션"] = {"value": f"{val_eok:.0f}억원", "score": val_score}

    else:
        # PER 데이터 없음 → 절대 밸류만 참고 (Value-Band 모델 연동 전 임시)
        notes.append("업종 PER 미입력 → 밸류 절대 규모 기준 (임시 채점, Value-Band 모델 연동 후 개선 예정)")
        val_eok = val / 10000
        if val_eok >= 300:
            val_score = 7
        elif val_eok >= 100:
            val_score = 5
        elif val_eok >= 30:
            val_score = 3
        else:
            val_score = 2
        breakdown["현재 밸류에이션"] = {"value": f"{val_eok:.0f}억원", "score": val_score}

    breakdown["예상가치 점수"] = {"value": "-", "score": val_score}

    return CategoryScore(
        name="예상가치 (밸류에이션)",
        score=val_score,
        max_score=MAX,
        breakdown=breakdown,
        note=" | ".join(notes)
    )


# ─────────────────────────────────────────────
# 4. Gemini 기반 핵심인력 평가 (10점)
# ─────────────────────────────────────────────

def score_key_people_with_gemini(data: StartupInput) -> tuple[CategoryScore, str]:
    """
    핵심인력 적절성 평가 (10점) — Gemini 사용
    반환: (CategoryScore, ai_advice_str)
    """
    MAX = 10

    # Gemini 클라이언트 초기화
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise EnvironmentError(
            ".env 파일에 GEMINI_API_KEY가 설정되지 않았습니다.\n"
            "예) GEMINI_API_KEY=AIza..."
        )
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")

    # 핵심인력 정보를 텍스트로 변환
    people_text = ""
    for person in data.key_people.members:
        edu_str = " / ".join(person.education) if person.education else "정보 없음"
        career_str = " / ".join(person.career) if person.career else "정보 없음"
        people_text += (
            f"\n- [{person.role}] {person.name}\n"
            f"  학력: {edu_str}\n"
            f"  경력: {career_str}"
        )

    # Growth Index 보완 조언을 위한 Rule-Based 점수 요약
    # (이 함수 호출 시점에는 아직 전체 점수 미산출 — 조언은 별도 호출에서 처리)

    prompt = f"""
당신은 스타트업 투자 전문가입니다.
아래 스타트업의 핵심인력이 해당 솔루션을 기획하고 구현하기에 얼마나 적합한지 평가해주세요.

[스타트업 정보]
- 회사명: {data.company_name}
- 업종: {data.industry}
- 개발 솔루션: {data.key_people.solution_description}

[핵심인력]
{people_text}

[평가 기준]
1. 각 인력의 학력이 솔루션 개발에 관련성이 있는가
2. 각 인력의 경력이 솔루션의 기술적/비즈니스적 요구사항에 적합한가
3. 팀 전체적으로 솔루션 구현에 필요한 역량(기술, 도메인, 비즈니스)이 균형있게 갖춰져 있는가
4. 부족한 역량이 있다면 무엇인가

[출력 형식 — 반드시 아래 JSON만 출력, 다른 텍스트 없이]
{{
  "score": <0~10 사이 정수>,
  "reason": "<평가 근거 2~4문장, 한국어>",
  "weakness": "<가장 부족한 역량 1~2가지, 없으면 '없음'>"
}}
"""

    response = model.generate_content(prompt)
    raw = response.text.strip()

    # JSON 파싱 (마크다운 코드블록 제거)
    if "```" in raw:
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    parsed = json.loads(raw.strip())

    ai_score = max(0, min(MAX, int(parsed.get("score", 5))))
    reason = parsed.get("reason", "")
    weakness = parsed.get("weakness", "")

    category = CategoryScore(
        name="핵심인력 적절성 (AI 평가)",
        score=ai_score,
        max_score=MAX,
        breakdown={
            "AI 평가 점수": {"value": f"{ai_score}/{MAX}", "score": ai_score},
            "부족 역량": {"value": weakness, "score": None},
        },
        note=f"Gemini 평가 | {reason[:80]}..."
    )

    return category, reason


# ─────────────────────────────────────────────
# 5. Gemini 보완 조언 생성
# ─────────────────────────────────────────────

def generate_advice_with_gemini(
    data: StartupInput,
    categories: list[CategoryScore],
    total_score: float,
    personnel_reason: str
) -> str:
    """
    전체 Growth Index 결과를 바탕으로 개선 조언 생성
    """
    api_key = os.getenv("GEMINI_API_KEY")
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")

    # 카테고리 점수 요약
    score_summary = "\n".join(
        f"- {c.name}: {c.score:.1f}/{c.max_score}점"
        + (f" ({c.note})" if c.note else "")
        for c in categories
    )

    prompt = f"""
당신은 스타트업 성장 전략 전문가입니다.
아래 스타트업의 Growth Index 분석 결과를 바탕으로, 성장 지수를 높이기 위한 실질적인 조언을 작성해주세요.

[기업 정보]
- 회사명: {data.company_name}
- 업종: {data.industry}
- 솔루션: {data.key_people.solution_description}

[Growth Index 분석 결과]
- 총점: {total_score:.1f}/100점
{score_summary}

[핵심인력 평가 요약]
{personnel_reason}

[작성 지침]
1. 점수가 낮은 카테고리 2~3개를 우선 지목
2. 각 카테고리별로 구체적인 개선 방안 1~2가지 제시
3. 단기(3개월) / 중기(6~12개월) 액션 플랜 구분
4. 전체 300자 이내, 항목별 bullet 형식, 한국어
"""

    response = model.generate_content(prompt)
    return response.text.strip()


# ─────────────────────────────────────────────
# 6. 등급 산정
# ─────────────────────────────────────────────

def get_grade(score: float) -> str:
    if score >= 85:
        return "S"
    elif score >= 70:
        return "A"
    elif score >= 55:
        return "B"
    elif score >= 40:
        return "C"
    else:
        return "D"


# ─────────────────────────────────────────────
# 7. 메인 계산 함수
# ─────────────────────────────────────────────

def calculate_growth_index(startup: StartupInput) -> GrowthIndexResult:
    """
    Growth Index 최종 산출 함수
    - Rule-Based 5개 카테고리 계산
    - Gemini 핵심인력 평가
    - Gemini 보완 조언 생성
    """
    calc_notes = []

    # Rule-Based 채점
    cat1 = score_revenue_growth(startup)
    cat2 = score_financial_stability(startup)
    cat3 = score_runway(startup)
    cat4 = score_investment(startup)
    cat5 = score_valuation(startup)

    # Gemini 핵심인력 채점
    cat6, personnel_reason = score_key_people_with_gemini(startup)

    categories = [cat1, cat2, cat3, cat4, cat5, cat6]

    # 총점 합산
    total = sum(c.score for c in categories)
    total = round(min(100.0, max(0.0, total)), 1)

    grade = get_grade(total)

    # Gemini 보완 조언 (총점 기반)
    advice = generate_advice_with_gemini(startup, categories, total, personnel_reason)

    if total < 40:
        calc_notes.append("Growth Index D등급 — 투자 유치 전 핵심 지표 개선 필요")
    if cat3.score == 0:
        calc_notes.append("🚨 런웨이 위험 — 긴급 자금 조달 권고")

    return GrowthIndexResult(
        company_name=startup.company_name,
        total_score=total,
        grade=grade,
        categories=categories,
        ai_advice=advice,
        ai_personnel_reason=personnel_reason,
        calculation_notes=calc_notes
    )


# ─────────────────────────────────────────────
# 8. 결과 출력 유틸리티
# ─────────────────────────────────────────────

def print_result(result: GrowthIndexResult):
    """터미널 출력용 포맷"""
    bar_len = 30
    filled = int((result.total_score / 100) * bar_len)
    bar = "█" * filled + "░" * (bar_len - filled)

    print("\n" + "=" * 60)
    print(f"  📊 Growth Index — {result.company_name}")
    print("=" * 60)
    print(f"  [{bar}]  {result.total_score:.1f} / 100   등급: {result.grade}")
    print("-" * 60)

    for cat in result.categories:
        cat_bar_filled = int((cat.score / cat.max_score) * 20)
        cat_bar = "■" * cat_bar_filled + "□" * (20 - cat_bar_filled)
        print(f"\n  {cat.name}")
        print(f"  [{cat_bar}]  {cat.score:.1f} / {cat.max_score}점")
        for k, v in cat.breakdown.items():
            score_str = f"  → {v['score']}점" if v['score'] is not None else ""
            print(f"    · {k}: {v['value']}{score_str}")
        if cat.note:
            print(f"    ※ {cat.note}")

    if result.calculation_notes:
        print("\n" + "-" * 60)
        for note in result.calculation_notes:
            print(f"  ⚠  {note}")

    print("\n" + "-" * 60)
    print("  🤖 AI 핵심인력 평가")
    print(f"  {result.ai_personnel_reason}")

    print("\n" + "-" * 60)
    print("  💡 Gemini 성장 조언")
    print(result.ai_advice)
    print("=" * 60 + "\n")


def result_to_dict(result: GrowthIndexResult) -> dict:
    """API 응답용 dict 변환 (외주팀 연동용)"""
    return {
        "company_name": result.company_name,
        "growth_index": result.total_score,
        "grade": result.grade,
        "categories": [
            {
                "name": c.name,
                "score": c.score,
                "max_score": c.max_score,
                "breakdown": c.breakdown,
                "note": c.note,
            }
            for c in result.categories
        ],
        "ai_advice": result.ai_advice,
        "ai_personnel_reason": result.ai_personnel_reason,
        "calculation_notes": result.calculation_notes,
    }


# ─────────────────────────────────────────────
# 9. 실행 예시 (테스트용)
# ─────────────────────────────────────────────

if __name__ == "__main__":
    sample = StartupInput(
        company_name="브이휠 (V-Wheel)",
        industry="핀테크 / AI 성장금융",
        founded_year=2023,
        financial=FinancialData(
            revenue_current_year=80_000,        # 8억원
            revenue_previous_year=30_000,       # 3억원
            operating_profit_rate=-12.0,        # 적자 12%
            debt_ratio=80.0,
            current_ratio=130.0,
            estimated_revenue_target=100_000,   # 추정 목표 10억원
        ),
        runway=RunwayData(
            cash_balance=200_000,               # 20억원
            monthly_burn_rate=8_000,            # 월 8천만원
        ),
        investment=InvestmentData(
            investment_stage="시리즈A",
            total_investment_amount=500_000,    # 50억원
            months_since_last_investment=8,
        ),
        valuation=ValuationData(
            current_valuation=3_000_000,        # 300억원
            industry_avg_per=0,                 # 미입력 → 절대 규모 기준
            net_profit=-9_600,                  # 적자
        ),
        key_people=KeyPeopleData(
            solution_description=(
                "AI 기반 비상장 스타트업 가치평가 및 투자 매칭 플랫폼. "
                "Growth Index, Value-Band 예측, Exit 확률 계산 등 ML 모델로 "
                "스타트업과 투자기관을 연결하는 SaaS 플랫폼."
            ),
            members=[
                KeyPersonData(
                    role="CEO",
                    name="김루센",
                    education=["연세대학교 경영학과 학사"],
                    career=["삼성증권 IB팀 7년", "VC 심사역 3년", "스타트업 창업 1회"],
                ),
                KeyPersonData(
                    role="CTO",
                    name="이기술",
                    education=["KAIST 전산학과 석사"],
                    career=["카카오 ML팀 5년", "AI 스타트업 CTO 2년"],
                ),
                KeyPersonData(
                    role="CFO",
                    name="박재무",
                    education=["고려대학교 경제학과 학사"],
                    career=["딜로이트 회계법인 6년", "핀테크 CFO 2년"],
                ),
            ],
        ),
    )

    result = calculate_growth_index(sample)
    print_result(result)

    # API 응답 형태 확인
    # import json
    # print(json.dumps(result_to_dict(result), ensure_ascii=False, indent=2))