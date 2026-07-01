"""유사회사 피어 추출 — 단일 신고서 PDF 스모크 (게이트 검증).

한 건이 정확히 나오는지 먼저 확인 → 그 다음 배치(전체 200건)로 확장.
실행 (peer_extraction 루트):
  python tests\\smoke_peer_extract.py --pdf "경로\\한라캐스트_증권신고서.pdf"
  python tests\\smoke_peer_extract.py --pdf "...pdf" --no-llm   # 정규식 폴백만
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.peer_extract import (extract_pdf_text, extract_peers,
                               extract_peer_names_regex, find_peer_section)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True, help="증권신고서 PDF 경로")
    ap.add_argument("--no-llm", action="store_true", help="Gemini 없이 정규식 폴백만")
    ap.add_argument("--show-section", action="store_true", help="국소화된 섹션 원문 출력")
    a = ap.parse_args()

    if not Path(a.pdf).exists():
        sys.exit(f"[중단] PDF 없음: {a.pdf}")

    print(f"=== 피어 추출 스모크: {Path(a.pdf).name} ===")
    text = extract_pdf_text(a.pdf)
    print(f"본문 텍스트 {len(text):,}자")

    section = find_peer_section(text)
    if not section:
        print("⚠ 유사회사 섹션 못 찾음 — SECTION_START 키워드 조정 필요.")
        print("  본문에서 '유사회사/비교회사' 등장 위치 진단:")
        for kw in ("유사회사", "비교회사", "비교기업", "유사기업"):
            i = text.find(kw)
            print(f"    {kw}: {'위치 '+str(i) if i >= 0 else '없음'}")
        return
    print(f"유사회사 섹션 국소화: {len(section)}자")
    if a.show_section:
        print("-" * 50); print(section[:1500]); print("-" * 50)

    if a.no_llm:
        peers = extract_peer_names_regex(section)
        print(f"\n[정규식 폴백] 피어 후보 {len(peers)}개:")
    else:
        peers = extract_peers(text, use_llm=True)
        print(f"\n[Gemini] 최종 선정 피어 {len(peers)}개:")
    for p in peers:
        print(f"  - {p}")

    if not peers:
        print("  (없음 — --show-section으로 섹션 원문 확인 후 키워드/프롬프트 조정)")
    else:
        print("\n→ 이 목록이 신고서 실제 비교회사와 맞으면 배치로 확장.")


if __name__ == "__main__":
    main()
