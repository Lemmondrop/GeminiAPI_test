"""유사회사 섹션 진단 — 본문 내 관련 키워드의 '모든' 등장 위치와 주변 텍스트 출력.

섹션 국소화가 실패할 때, 실제 비교회사 표가 본문 어디에 있는지 눈으로 찾는 용도.
실행: python tests\\diag_peer_section.py --rcept 20220111000212
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.dart_client import DartClient

KEYWORDS = ["유사회사", "비교회사", "비교기업", "유사기업", "비교 대상",
            "비교대상", "선정", "주당 평가가액", "PER", "적용 PER", "할인율"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rcept", required=True)
    ap.add_argument("--window", type=int, default=90, help="각 등장 위치 주변 글자수")
    ap.add_argument("--max-hits", type=int, default=8, help="키워드당 최대 표시 횟수")
    a = ap.parse_args()

    text = DartClient().fetch_document(a.rcept)
    print(f"본문 {len(text):,}자\n")

    for kw in KEYWORDS:
        idxs = []
        start = 0
        while True:
            i = text.find(kw, start)
            if i < 0 or len(idxs) >= a.max_hits:
                break
            idxs.append(i)
            start = i + 1
        total = text.count(kw)
        print(f"### '{kw}'  총 {total}회" + (f" (상위 {a.max_hits} 표시)" if total > a.max_hits else ""))
        for i in idxs:
            ctx = text[max(0, i - 20):i + a.window].replace("\n", " ")
            print(f"  @{i:>6}: ...{ctx}...")
        print()


if __name__ == "__main__":
    main()
