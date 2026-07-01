"""최초 증권신고서 탐색 + 비교회사 표 진단.

발행조건확정본엔 비교회사 PER 표가 없음 → 같은 corp_code의 '최초 증권신고서'를 찾아
그 본문에서 유사회사/비교회사/PER 키워드 위치를 확인.

실행:
  python tests\\diag_initial_filing.py --corp 01140837                  # 공시목록 + 최초신고서 탐색
  python tests\\diag_initial_filing.py --corp 01140837 --probe          # 최초신고서 본문 키워드 진단
corp_code는 enriched.csv의 corp_code 컬럼(이지트로닉스=01140837).
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.dart_client import DartClient

KW = ["유사회사", "비교회사", "비교기업", "유사기업", "상대가치", "PER", "주당 평가가액"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--corp", required=True, help="8자리 corp_code")
    ap.add_argument("--probe", action="store_true", help="최초신고서 본문 키워드 위치 진단")
    a = ap.parse_args()

    dc = DartClient()
    print(f"=== 공시목록 (corp_code={a.corp}) ===")
    filings = dc.list_filings(a.corp, pblntf_ty="C")
    regs = [f for f in filings if "증권신고서" in (f.get("report_nm") or "")]
    print(f"발행공시 {len(filings)}건 / 증권신고서 {len(regs)}건")
    for f in regs:
        print(f"  {f.get('rcept_dt')}  {f.get('rcept_no')}  {f.get('report_nm')}")

    init = dc.find_initial_registration(a.corp)
    if not init:
        print("\n⚠ 증권신고서 없음")
        return
    print(f"\n→ 최초 증권신고서 선택: {init.get('rcept_dt')}  {init.get('rcept_no')}  "
          f"{init.get('report_nm')}")

    if a.probe:
        rcept = init["rcept_no"]
        print(f"\n=== 최초신고서 본문 진단 (rcept={rcept}) ===")
        text = dc.fetch_document(rcept)
        print(f"본문 {len(text):,}자")
        for kw in KW:
            cnt = text.count(kw)
            i = text.find(kw)
            ctx = ""
            if i >= 0:
                ctx = "  ...{}...".format(text[max(0, i-15):i+70].replace("\n", " "))
            print(f"  '{kw}': {cnt}회" + (f"  첫@{i}{ctx}" if i >= 0 else ""))


if __name__ == "__main__":
    main()
