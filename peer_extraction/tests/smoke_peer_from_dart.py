"""유사회사 라벨 추출 — DART 신고서 1건 라이브 스모크 (Phase B 게이트).

발행조건확정본엔 표가 없으므로 --corp로 주면 최초 증권신고서를 자동 탐색.
실행 (peer_extraction 루트):
  python tests\\smoke_peer_from_dart.py --rcept 20211222000428      # 최초신고서 직접
  python tests\\smoke_peer_from_dart.py --corp 01140837             # corp_code→최초신고서 자동
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.dart_client import DartClient
from core.peer_extract import collect_valuation_text, extract_filing_labels


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rcept", help="최초 증권신고서 접수번호(직접)")
    ap.add_argument("--corp", help="corp_code(8자리) → 최초 증권신고서 자동 탐색")
    ap.add_argument("--show-regions", action="store_true", help="수집된 평가구간 출력")
    a = ap.parse_args()
    if not (a.rcept or a.corp):
        sys.exit("--rcept 또는 --corp 필요")

    dc = DartClient()
    rcept = a.rcept
    if not rcept:
        init = dc.find_initial_registration(a.corp)
        if not init:
            sys.exit(f"[중단] corp_code={a.corp} 증권신고서 없음")
        rcept = init["rcept_no"]
        print(f"최초 증권신고서: {init.get('rcept_dt')}  {rcept}  {init.get('report_nm')}")

    print(f"\n=== 라벨 추출: rcept={rcept} ===")
    print("  [1/3] 본문 다운로드...")
    text = dc.fetch_document(rcept)
    print(f"  본문 {len(text):,}자")

    print("  [2/3] 평가구간 수집...")
    regions = collect_valuation_text(text)
    print(f"  수집 {len(regions):,}자")
    if not regions:
        print("  ⚠ 평가구간 못 찾음 — 발행조건확정본이거나 앵커 불일치. --corp로 최초신고서 시도.")
        return
    if a.show_regions:
        print("-" * 55); print(regions[:2500]); print("-" * 55)

    print("  [3/3] Gemini 라벨 추출...")
    lab = extract_filing_labels(text)

    print(f"\n발행사 KSIC: {lab['issuer_ksic']}")
    print(f"1차 비교 KSIC ({len(lab['comparison_ksic'])}): {', '.join(lab['comparison_ksic'])}")
    print(f"\n최종 선정 유사회사 ({len(lab['final_peers'])}):")
    for p in lab["final_peers"]:
        print(f"  - {p['name']:<16} PER {p['per']}")
    if lab["excluded"]:
        print(f"\n탈락사 ({len(lab['excluded'])}):")
        for e in lab["excluded"]:
            print(f"  - {e['name']:<16} {e['reason']}")

    print("\n→ 이지트로닉스 기대: 이노와이어리스/오이솔루션/와이엠텍/우리산업, 탈락=영화테크(137.79)")
    print("  맞으면 배치(발행조건확정 rcept→최초신고서→추출→peer_labels.csv)로 확장.")


if __name__ == "__main__":
    main()