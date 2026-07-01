"""core/stage_router.py — IR 신호 기반 IPO트랙 vs 다음라운드트랙 자동 판별.

배경
----
지금까지는 "이 회사는 IPO를 목표로 하는 단계인가, 다음 투자 라운드를 목표로
하는 단계인가"를 사람이 IR을 읽고 매번 수동으로 판단해서
calculate_valuation_band()(IPO 트랙, PER 기반 예상공모시총) 또는
calculate_equity_value_psr_band()(다음 라운드 트랙, PSR 기반 가치밴드) 중
하나를 직접 골라 호출했다.

실제 플랫폼에서는 사용자가 "예상 밸류에이션" 페이지에서 IR을 업로드하면
그 즉시 결과가 나와야 하므로, 이 모듈이 그 핵심 분기를 자동화한다.

설계 원칙
--------
1. Zero-label 상황이므로 지도학습이 아니라 규칙 기반 가중치 점수제로 시작한다.
   추후 실제 판단 결과 + 사람 피드백(수정 이력)이 쌓이면, 아래 스키마
   (StageSignals 입력 / StageDecision 출력)를 그대로 유지한 채 스코어링
   로직만 학습된 분류기로 교체할 수 있도록 설계했다.
2. 단일 신호로 자르지 않는다. IPO 언급, 최근 투자 라운드, IPO 목표 시점,
   설립연도, 재무 프로필을 각각 점수화해서 합산하고, 신뢰도(confidence)와
   판단 근거(reasoning)를 함께 반환한다 — 블랙박스가 아니라 "왜 이렇게
   판단했는지"를 외주 프론트가 사용자에게 그대로 보여줄 수 있어야 한다.
3. 신호가 부족하거나(총점 0) 신호끼리 상충하면(예: Seed 단계인데 IPO
   임박 언급) 자동 계산은 하되 review_needed=True로 표시한다. 완전 수동으로
   되돌리지 않고, "자동 판단 + 확인 요청" 수준으로 UX를 설계한다.
4. 애매한 케이스에서 실수 방향은 next_round_track(PSR, 더 넓은 밴드) 쪽으로
   둔다. IPO 트랙(PER 기반 공모시총)은 신뢰구간이 좁아서, 근거 없이 IPO
   트랙으로 잘못 분류하면 사용자에게 과신 위험한 숫자를 보여주게 된다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import List, Optional


# ─────────────────────────────────────────────────────────────────
# 1. 투자 라운드 정규화
# ─────────────────────────────────────────────────────────────────

class InvestmentRound(str, Enum):
    PRE_SEED = "pre_seed"
    SEED = "seed"
    ANGEL = "angel"
    SERIES_A = "series_a"
    SERIES_B = "series_b"
    SERIES_C = "series_c"
    SERIES_D_PLUS = "series_d_plus"
    PRE_IPO = "pre_ipo"
    UNKNOWN = "unknown"


# 낮은 단계 → 높은 단계 순서. normalize_round에서 "가장 앞선 단계"를 고를 때 사용.
_ROUND_ORDER = [
    InvestmentRound.PRE_SEED,
    InvestmentRound.SEED,
    InvestmentRound.ANGEL,
    InvestmentRound.SERIES_A,
    InvestmentRound.SERIES_B,
    InvestmentRound.SERIES_C,
    InvestmentRound.SERIES_D_PLUS,
    InvestmentRound.PRE_IPO,
]

_ROUND_PATTERNS = [
    (r"pre[-\s]?seed|프리\s*시드", InvestmentRound.PRE_SEED),
    # "Series Pre-A"/"Pre-Series A"처럼 Seed와 Series A 사이 단계 — Series A
    # 패턴보다 먼저 검사해야 하지만, "series\s*a\b"는 애초에 "series"와 "a"
    # 사이에 "pre-"가 끼면 매칭 안 되므로 순서 자체보다 패턴 존재가 핵심.
    # 실제 IR에서 나온 표기라 이번에 추가(2026-07 실사례로 발견된 누락).
    (r"pre[-\s]?series\s*a\b|series\s*pre[-\s]?a\b|프리\s*시리즈\s*a|프리a\b",
     InvestmentRound.SEED),
    (r"\bseed\b|시드", InvestmentRound.SEED),
    (r"angel|엔젤", InvestmentRound.ANGEL),
    (r"series\s*a\b|시리즈\s*a\b|시리즈에이", InvestmentRound.SERIES_A),
    (r"series\s*b\b|시리즈\s*b\b|시리즈비", InvestmentRound.SERIES_B),
    (r"series\s*c\b|시리즈\s*c\b|시리즈씨", InvestmentRound.SERIES_C),
    (r"series\s*d\b|series\s*e\b|시리즈\s*d\b|시리즈\s*e\b", InvestmentRound.SERIES_D_PLUS),
    (r"pre[-\s]?ipo|프리\s*아이피오|상장\s*전\s*투자", InvestmentRound.PRE_IPO),
]


def normalize_round(raw: Optional[str]) -> InvestmentRound:
    """IR 원문 텍스트("Seed+Angel", "Series A", "Pre-IPO" 등)에서
    가장 앞선(최신) 투자 라운드를 뽑아 InvestmentRound로 정규화한다.

    "Seed+Angel"처럼 복수 라운드가 한 문자열에 섞여 있으면, 더 진행된
    단계(=마지막에 받은 라운드로 간주)를 채택한다.
    """
    if not raw:
        return InvestmentRound.UNKNOWN
    text = raw.lower()
    matched = [rnd for pattern, rnd in _ROUND_PATTERNS if re.search(pattern, text)]
    if not matched:
        return InvestmentRound.UNKNOWN
    return max(matched, key=_ROUND_ORDER.index)


# ─────────────────────────────────────────────────────────────────
# 2. 입력 / 출력 스키마
# ─────────────────────────────────────────────────────────────────

@dataclass
class StageSignals:
    """IR에서 추출 가능한 필드. 값을 모르면 None으로 둔다 — 신호 부족은
    review_needed로 자연스럽게 반영된다."""

    company_name: str
    founding_year: Optional[int] = None
    latest_round: Optional[str] = None          # 원문 텍스트, 예: "Series A", "Seed+Angel"
    ipo_mentioned: bool = False                  # IR에 IPO 계획/시점이 명시되어 있는가
    ipo_target_year: Optional[int] = None
    net_income: Optional[float] = None           # 최근 순이익(억원, 트레일링)
    future_net_income: Optional[float] = None    # 추정 순이익(억원, 기술특례 트랙용)
    revenue: Optional[float] = None               # 최근 매출(억원)
    current_year: int = field(default_factory=lambda: datetime.now().year)


@dataclass
class StageDecision:
    company_name: str
    track: str                       # "ipo_track" | "next_round_track"
    sub_track: Optional[str]         # "general" | "tech_special" | None (다음라운드 트랙이면 None)
    confidence: float                # 0.0 ~ 1.0
    score_ipo: float
    score_next_round: float
    reasoning: List[str]
    review_needed: bool
    normalized_round: str
    source: str = "auto"             # "auto" | "manual_override"

    def to_api_response(self) -> dict:
        """외주 프론트가 그대로 렌더링할 수 있는 JSON 직렬화 형태."""
        return {
            "company_name": self.company_name,
            "track": self.track,
            "sub_track": self.sub_track,
            "confidence": round(self.confidence, 3),
            "review_needed": self.review_needed,
            "reasoning": self.reasoning,
            "normalized_round": self.normalized_round,
            "source": self.source,
            "_debug_scores": {
                "ipo_track": round(self.score_ipo, 2),
                "next_round_track": round(self.score_next_round, 2),
            },
        }


# ─────────────────────────────────────────────────────────────────
# 3. IPO 트랙 내부 서브트랙 판별 (일반 vs 기술특례)
# ─────────────────────────────────────────────────────────────────

def determine_ipo_subtrack(
    net_income: Optional[float], future_net_income: Optional[float]
) -> Optional[str]:
    """일반 트랙(트레일링 실적 기반) vs 기술특례 트랙(추정 실적 + 현재가치 할인)
    판별. 순이익이 흑자면 일반 트랙, 적자인데 미래 추정치가 있으면 기술특례,
    아무 정보도 없으면 판별 불가(None)로 두고 review_needed로 넘긴다."""
    if net_income is not None and net_income > 0:
        return "general"
    if future_net_income is not None:
        return "tech_special"
    if net_income is not None and net_income <= 0:
        # 적자인데 미래 추정치가 없음 → 기술특례로 가야 하지만 입력 보강 필요
        return "tech_special"
    return None


# ─────────────────────────────────────────────────────────────────
# 4. 메인 라우팅 함수
# ─────────────────────────────────────────────────────────────────

# 라운드별 트랙 점수 가중치. (ipo_score 기여분, next_round_score 기여분)
_ROUND_WEIGHTS = {
    InvestmentRound.PRE_SEED: (0.0, 3.0),
    InvestmentRound.SEED: (0.0, 3.0),
    InvestmentRound.ANGEL: (0.0, 3.0),
    InvestmentRound.SERIES_A: (0.0, 2.0),
    InvestmentRound.SERIES_B: (0.5, 0.5),   # 중간 단계 — 다른 신호에 판단을 맡김
    InvestmentRound.SERIES_C: (2.0, 0.0),
    InvestmentRound.SERIES_D_PLUS: (3.0, 0.0),
    InvestmentRound.PRE_IPO: (4.0, 0.0),
    InvestmentRound.UNKNOWN: (0.0, 0.0),
}

_EARLY_ROUNDS = {InvestmentRound.PRE_SEED, InvestmentRound.SEED, InvestmentRound.ANGEL}

_CONFIDENCE_REVIEW_THRESHOLD = 0.25

# 신호가 "방향은 만장일치인데 양이 적은" 경우(예: 신호 1개, +1.0점 뿐인데
# 반대쪽이 0점이라 비율상 100% 만장일치) confidence가 부풀려지는 문제가
# 실사례(피엔피 — 신호가 설립연도 하나뿐인데 confidence=1.00으로 나옴)에서
# 발견됨. IPO 언급(+5.0, 가장 강한 단일 신호) 하나면 "증거 충분"으로 보고,
# 그 이하는 비례해서 confidence를 깎는다.
_EVIDENCE_SATURATION = 5.0

# score_ipo가 score_next보다 커도, 절대량이 이 미만이면(예: "설립 7년+" 같은
# 약한 신호 하나뿐) ipo_track으로 확정하지 않고 next_round_track 기본값으로
# 되돌린다. founding_year age>=7(+1.0), 흑자+매출(+1.0) 같은 신호들에 스스로
# "(약한 신호)"라고 주석까지 달아놓고 정작 그거 하나로 트랙이 뒤집히게 둔 게
# 설계원칙 4번("애매하면 next_round_track 쪽으로")과 어긋났던 실제 버그
# (2026-07 피엔피 실사례에서 발견: ipo_mentioned=False인데도 ipo_track으로 감).
_MIN_IPO_SCORE_TO_COMMIT = 2.0


def determine_company_stage(
    signals: StageSignals, manual_override: Optional[str] = None
) -> StageDecision:
    """IR 신호를 받아 "ipo_track" vs "next_round_track"을 자동 판별한다.

    Parameters
    ----------
    signals : StageSignals
        IR에서 추출한 필드.
    manual_override : "ipo_track" | "next_round_track" | None
        사용자가 자동 판별 결과를 직접 뒤집고 싶을 때만 사용. 기본은 None
        (완전 자동). override를 줘도 자동 점수는 그대로 계산해서 reasoning에
        남긴다 — "자동 판단과 다르게 선택했다"는 사실 자체가 로그로 남아야
        나중에 스코어링 규칙을 튜닝할 때 근거가 된다.
    """
    reasoning: List[str] = []
    score_ipo = 0.0
    score_next = 0.0

    normalized_round = normalize_round(signals.latest_round)

    # 신호 1: 최근 투자 라운드
    w_ipo, w_next = _ROUND_WEIGHTS[normalized_round]
    score_ipo += w_ipo
    score_next += w_next
    if normalized_round != InvestmentRound.UNKNOWN:
        reasoning.append(
            f"최근 투자 라운드 '{signals.latest_round}' → {normalized_round.value} "
            f"(IPO +{w_ipo}, 다음라운드 +{w_next})"
        )
    else:
        reasoning.append("최근 투자 라운드 정보 없음 또는 인식 불가")

    # 신호 2: IPO 언급 (가장 강한 직접 신호)
    if signals.ipo_mentioned:
        score_ipo += 5.0
        reasoning.append("IR에 IPO 계획/시점 명시 → IPO +5.0 (가장 강한 신호)")

    # 신호 3: IPO 목표 시점의 임박도
    if signals.ipo_target_year is not None:
        horizon = signals.ipo_target_year - signals.current_year
        if 0 <= horizon <= 3:
            score_ipo += 3.0
            reasoning.append(f"IPO 목표 {horizon}년 후로 임박 → IPO +3.0")
        elif 3 < horizon <= 6:
            score_ipo += 1.5
            reasoning.append(f"IPO 목표 {horizon}년 후 → IPO +1.5")
        else:
            score_ipo += 0.5
            reasoning.append(
                f"IPO 목표 시점이 {horizon}년으로 매우 멀거나 이미 지남 → IPO +0.5 (약한 신호)"
            )

    # 신호 4: 설립연도(회사 나이)
    if signals.founding_year is not None:
        age = signals.current_year - signals.founding_year
        if age <= 2:
            score_next += 2.0
            reasoning.append(f"설립 {age}년차로 초기 단계 → 다음라운드 +2.0")
        elif age >= 7:
            score_ipo += 1.0
            reasoning.append(f"설립 {age}년차로 업력 충분 → IPO +1.0 (약한 신호)")

    # 신호 5: 재무 프로필 (매출 규모 — IPO 준비도의 약한 보조 신호)
    if signals.net_income is not None and signals.net_income > 0 and \
       signals.revenue is not None and signals.revenue >= 100:
        score_ipo += 1.0
        reasoning.append(
            f"흑자 + 매출 {signals.revenue}억원(≥100억) → IPO +1.0 (약한 신호)"
        )

    total = score_ipo + score_next
    margin_confidence = abs(score_ipo - score_next) / total if total > 0 else 0.0
    # 증거량 보정: 총점이 saturation(5.0) 미만이면 그 비율만큼 confidence를 깎는다.
    # 예) 신호 1개(+1.0)만으로 만장일치여도 evidence_strength=0.2 → confidence 0.2로 낮아짐.
    evidence_strength = min(1.0, total / _EVIDENCE_SATURATION) if total > 0 else 0.0
    confidence = margin_confidence * evidence_strength
    if 0 < total < _EVIDENCE_SATURATION:
        reasoning.append(
            f"증거량 보정: 총점 {total:.1f}점(<{_EVIDENCE_SATURATION:.0f}) → "
            f"신뢰도에 {evidence_strength:.0%} 가중 적용"
        )
    # 동점(특히 0:0)일 때는 next_round_track으로 — IPO 트랙은 신뢰구간이 좁아
    # 근거 없이 배정하면 위험하므로, "이겨야만" IPO 트랙을 준다. 추가로,
    # 상대적으로 이겨도 절대량이 _MIN_IPO_SCORE_TO_COMMIT 미만이면(약한 신호
    # 하나뿐인 경우) 여전히 next_round_track — "이기기만 하면 OK"가 아니라
    # "이길 만큼 이겨야 OK".
    ipo_wins_relatively = score_ipo > score_next
    ipo_clears_minimum = score_ipo >= _MIN_IPO_SCORE_TO_COMMIT
    auto_track = "ipo_track" if (ipo_wins_relatively and ipo_clears_minimum) else "next_round_track"
    if ipo_wins_relatively and not ipo_clears_minimum:
        reasoning.append(
            f"IPO 신호가 근소 우위(IPO {score_ipo:.1f} vs 다음라운드 {score_next:.1f})지만 "
            f"절대량이 약함(<{_MIN_IPO_SCORE_TO_COMMIT:.1f}) → 안전하게 다음라운드 트랙 기본값 적용"
        )

    review_needed = total == 0 or confidence < _CONFIDENCE_REVIEW_THRESHOLD

    # 상충 신호 명시적 탐지: 초기 라운드인데 IPO를 언급한 이례적 조합
    if signals.ipo_mentioned and normalized_round in _EARLY_ROUNDS:
        review_needed = True
        reasoning.append(
            "⚠ 초기 라운드(Seed/Angel 등)인데 IPO가 언급된 이례적 조합 → 확인 필요"
        )

    if total == 0:
        reasoning.append("판별 가능한 신호가 전혀 없음 → 기본값(다음라운드 트랙)으로 처리, 확인 필요")

    # 최종 트랙 결정: override가 있으면 override를 따르되 자동 판단은 로그로 남김
    if manual_override in ("ipo_track", "next_round_track"):
        if manual_override != auto_track:
            reasoning.append(
                f"수동 override 적용: 자동 판단은 '{auto_track}'이었으나 "
                f"'{manual_override}'로 사용자가 직접 변경"
            )
        track = manual_override
        source = "manual_override"
    else:
        track = auto_track
        source = "auto"

    sub_track = None
    if track == "ipo_track":
        sub_track = determine_ipo_subtrack(signals.net_income, signals.future_net_income)
        if sub_track is None:
            review_needed = True
            reasoning.append("IPO 트랙이지만 순이익/추정 순이익 정보 부족 → 서브트랙(일반/기술특례) 판별 불가, 확인 필요")
        else:
            reasoning.append(f"서브트랙: {sub_track}")

    return StageDecision(
        company_name=signals.company_name,
        track=track,
        sub_track=sub_track,
        confidence=confidence,
        score_ipo=score_ipo,
        score_next_round=score_next,
        reasoning=reasoning,
        review_needed=review_needed,
        normalized_round=normalized_round.value,
        source=source,
    )