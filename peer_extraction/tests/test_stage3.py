"""Stage 3 단위테스트 (네트워크 불필요, 가짜 판정기).

검증: embed_score 상위K 선별 · tier 선별/정렬 · LLM 캐싱(중복호출 방지) ·
      JSON 파싱 가드(코드펜스/오류tier/범위) · 캐시키 결정론.
실행: peer_extraction 루트에서  python tests/test_stage3.py
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

from config import Stage3Config
from core.llm_judge import LlmJudge, parse_judgment, _key, TIERS
from stages.stage3_business import apply_stage3, business_text, final_peers

results: list[tuple[str, bool, str]] = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))


class MockJudge(LlmJudge):
    """genai 없이 동작. _call만 오버라이드(캐시·키·judge_many는 실제 로직 사용)."""

    def __init__(self, table):
        super().__init__(client="dummy", cache_path=None)   # client≠None → genai 미사용
        self.table = table
        self.call_count = 0

    def _call(self, issuer, candidate):
        self.call_count += 1
        return self.table.get(candidate, {"tier": "낮음", "score": 10, "rationale": "기본"})


POOL = pd.DataFrame({
    "회사명": ["A압출", "B전지", "C제과", "D다이캐스팅"],
    "종목코드": ["000001", "000002", "000003", "000004"],
    "주요제품": ["알루미늄 압출", "2차전지 부품", "제과", "알루미늄 다이캐스팅"],
    "업종": ["1차금속", "전기장비", "음식료품", "1차금속"],
    "embed_score": [0.90, 0.80, 0.30, 0.85],
})
TABLE = {
    "알루미늄 압출 | 1차금속": {"tier": "높음", "score": 90, "rationale": "직접경쟁"},
    "2차전지 부품 | 전기장비": {"tier": "중간", "score": 60, "rationale": "일부중첩"},
    "알루미늄 다이캐스팅 | 1차금속": {"tier": "높음", "score": 88, "rationale": "동일밸류체인"},
    # C제과(제과 | 음식료품)는 미등록 → 기본 '낮음'
}
ISSUER = "알루미늄 압출·다이캐스팅 자동차부품 제조"


def run():
    # ---- business_text ----
    check("business_text 결합", business_text(POOL.iloc[0]) == "알루미늄 압출 | 1차금속",
          business_text(POOL.iloc[0]))

    # ---- apply_stage3: 전체 판정 ----
    j = MockJudge(TABLE)
    s3 = apply_stage3(POOL, ISSUER, j, Stage3Config(top_k_judge=4), progress=False)
    fp = final_peers(s3)
    fpn = list(fp["회사명"])
    check("낮음(C제과) 탈락", "C제과" not in fpn, str(fpn))
    check("높음·중간 통과(A,B,D)", set(fpn) == {"A압출", "B전지", "D다이캐스팅"}, str(fpn))
    # 정렬: 높음(score순) → 중간.  A(90),D(88) 먼저, B(60) 뒤
    check("정렬: 높음 먼저(A,D) 그다음 중간(B)", fpn == ["A압출", "D다이캐스팅", "B전지"], str(fpn))
    check("rationale 부착", s3.iloc[0]["rationale"] == "직접경쟁")

    # ---- top_k 선별: 상위 2개만 판정 ----
    j2 = MockJudge(TABLE)
    s3b = apply_stage3(POOL, ISSUER, j2, Stage3Config(top_k_judge=2), progress=False)
    check("top_k=2 → 2건만 판정(embed 상위 A,D)", j2.call_count == 2, f"calls={j2.call_count}")
    check("top_k=2 대상 = A,D (embed 0.90,0.85)", set(s3b["회사명"]) == {"A압출", "D다이캐스팅"}, str(list(s3b["회사명"])))

    # ---- 캐싱: 재실행 시 추가 호출 없음 ----
    before = j.call_count
    apply_stage3(POOL, ISSUER, j, Stage3Config(top_k_judge=4), progress=False)   # 같은 판정기 재사용
    check("캐시 적중 → 재호출 0", j.call_count == before, f"{before}→{j.call_count}")

    # ---- parse_judgment 가드 ----
    check("정상 JSON", parse_judgment('{"tier":"높음","score":95,"rationale":"r"}')["tier"] == "높음")
    check("코드펜스 제거", parse_judgment('```json\n{"tier":"중간","score":50,"rationale":"r"}\n```')["tier"] == "중간")
    check("오류 tier → 낮음", parse_judgment('{"tier":"매우높음","score":50}')["tier"] == "낮음")
    check("score 상한 클램프", parse_judgment('{"tier":"높음","score":150}')["score"] == 100)
    check("score 하한 클램프", parse_judgment('{"tier":"낮음","score":-5}')["score"] == 0)
    check("본문 중 JSON 구제", parse_judgment('판단: {"tier":"중간","score":40,"rationale":"r"} 끝')["score"] == 40)

    # ---- 캐시 키 결정론 ----
    k1 = _key("m", "issuer", "cand")
    check("같은 입력 → 같은 키", k1 == _key("m", "issuer", "cand"))
    check("후보 다르면 키 다름", k1 != _key("m", "issuer", "cand2"))
    check("발행사 다르면 키 다름", k1 != _key("m", "issuer2", "cand"))


def main():
    run()
    passed = sum(1 for _, ok, _ in results if ok)
    print("=" * 60)
    for name, ok, detail in results:
        mark = "PASS" if ok else "FAIL"
        line = f"[{mark}] {name}"
        if not ok and detail:
            line += f"   <- {detail}"
        print(line)
    print("=" * 60)
    print(f"{passed}/{len(results)} passed")
    sys.exit(0 if passed == len(results) else 1)


if __name__ == "__main__":
    main()