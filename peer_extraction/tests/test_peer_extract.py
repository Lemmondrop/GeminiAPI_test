"""피어 라벨 추출기 단위테스트 (네트워크 불필요, Gemini 모킹).

실제 이지트로닉스 신고서 구조 기반 픽스처:
  - 1차 KSIC 선정기준표 / 최종 선정표(46.12 등) / 탈락사(영화테크 137.79).
구간수집 · KSIC 파싱 · 최종피어+PER · 탈락사 · 회사명 검증 · 객체 파싱.
실행: peer_extraction 루트에서  python tests\\test_peer_extract.py
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.peer_extract import (clean_company_name, collect_valuation_text,
                               extract_filing_labels, extract_peers,
                               is_valid_company_name, _parse_labels)

results: list[tuple[str, bool, str]] = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))


# 실제 신고서 구조 축약 (앞: 위험고지 오탐 유도 / 중간: KSIC표 / 끝: 최종표)
FILING = """가. 비교회사 부적합 가능성. 당사는 유사회사를 선정하여 공모가액 산출에 적용하였습니다.
(위험고지 서술 — 여기의 '유사회사'는 표가 아님)
... 중략 ...
【 국내 유사회사 선정 기준표 】
구 분 선정기준 세부 검토기준 선정회사
1차 업종유사성 한국표준산업분류상 다음 중 하나에 속하는 상장회사
① (C26429) 기타 무선 통신장비 제조업 ② (C26410) 유선 통신장비 제조업
③ (C28121) 전기회로 개폐 제조업 ④ (C30399) 그외 기타 자동차 부품 제조업
⑤ (C29172) 공기조화장치 제조업
이노와이어리스, 오이솔루션, 와이엠텍, 영화테크, 우리산업 등 122개사
... 중략 ...
【 유사회사 PER 산출】 평균 PER 산정
4. P/E 50배 이상 비경상 멀티플 제외(Outlier)
선정여부
이노와이어리스 적정 충족 충족 충족(46.12) 선정
오이솔루션 적정 충족 충족 충족(23.52) 선정
와이엠텍 적정 충족 충족 충족(32.39) 선정
영화테크 적정 충족 충족 미충족(137.79) 미선정
우리산업 적정 충족 충족 충족(28.16) 선정
【 최종 유사기업 선정결과 요약 】
최종 유사기업 이노와이어리스, 오이솔루션, 와이엠텍, 우리산업
"""

LABELS_JSON = """{
  "issuer_ksic": "C28112",
  "comparison_ksic": ["C26429","C26410","C28121","C30399","C29172"],
  "final_peers": [
    {"name":"이노와이어리스","per":46.12},
    {"name":"오이솔루션","per":23.52},
    {"name":"와이엠텍","per":32.39},
    {"name":"우리산업","per":28.16}
  ],
  "excluded": [{"name":"영화테크","reason":"PER 137.79 아웃라이어(50배 초과)"}]
}"""


class FakeResp:
    def __init__(self, t): self.text = t


class FakeClient:
    def __init__(self, payload): self._p = payload
    class _M:
        def __init__(self, p): self._p = p
        def generate_content(self, **kw): return FakeResp(self._p)
    @property
    def models(self): return self._M(self._p)


def run():
    # ----- 구간 수집: KSIC표 + 최종표 포함, 위험고지 단독 '유사회사'는 앵커 아님 -----
    regions = collect_valuation_text(FILING)
    check("구간에 KSIC표 포함", "한국표준산업분류" in regions)
    check("구간에 최종표 포함", "최종 유사기업" in regions)
    check("구간 비어있지 않음", len(regions) > 100)

    # ----- 회사명 검증 -----
    check("정상 통과", is_valid_company_name("이노와이어리스"))
    check("선정여부 머리글 제외", not is_valid_company_name("선정여부"))
    check("제조업 토큰 제외", not is_valid_company_name("그외 기타 자동차 부품 제조업"))
    check("PER 괄호 정제", clean_company_name("이노와이어리스(46.12)") == "이노와이어리스")

    # ----- 라벨 파싱 (모킹) -----
    labels = extract_filing_labels(FILING, client=FakeClient(LABELS_JSON))
    names = [p["name"] for p in labels["final_peers"]]
    check("최종피어 4사", names == ["이노와이어리스","오이솔루션","와이엠텍","우리산업"], str(names))
    check("탈락 영화테크 제외", "영화테크" not in names)
    check("PER 동반", labels["final_peers"][0]["per"] == 46.12, str(labels["final_peers"][0]))
    check("1차 KSIC 5개", labels["comparison_ksic"] ==
          ["C26429","C26410","C28121","C30399","C29172"], str(labels["comparison_ksic"]))
    check("발행사 KSIC", labels["issuer_ksic"] == "C28112", str(labels["issuer_ksic"]))
    check("탈락사+사유", labels["excluded"] and labels["excluded"][0]["name"] == "영화테크",
          str(labels["excluded"]))

    # ----- extract_peers 단축 -----
    check("extract_peers 4사", extract_peers(FILING, client=FakeClient(LABELS_JSON)) ==
          ["이노와이어리스","오이솔루션","와이엠텍","우리산업"])

    # ----- 노이즈/코드펜스 견고성 -----
    noisy = '```json\n{"final_peers":["유사회사","이노와이어리스","평균"],"comparison_ksic":["c26429"," X "]}\n```'
    lab2 = _parse_labels(noisy)
    check("코드펜스+머리글 제거", [p["name"] for p in lab2["final_peers"]] == ["이노와이어리스"],
          str(lab2["final_peers"]))
    check("KSIC 소문자 정규화", lab2["comparison_ksic"] == ["C26429"], str(lab2["comparison_ksic"]))

    # 앵커 없으면 빈 라벨
    empty = extract_filing_labels("관련 평가 내용 없음", client=FakeClient(LABELS_JSON))
    check("앵커 없음 빈 라벨", empty["final_peers"] == [] and empty["comparison_ksic"] == [])


def main():
    run()
    passed = sum(1 for _, ok, _ in results if ok)
    print("=" * 60)
    for name, ok, detail in results:
        print(f"[{'PASS' if ok else 'FAIL'}] {name}" + ("" if ok or not detail else f"   <- {detail}"))
    print("=" * 60)
    print(f"{passed}/{len(results)} passed")
    sys.exit(0 if passed == len(results) else 1)


if __name__ == "__main__":
    main()