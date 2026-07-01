"""탈락 정답 피어의 실제 재무값 확인 — 연간/LTM 순이익이 진짜 음수인지.

연간 기준으로 바꿔도 안 살아난 이유 규명:
  - net_income_annual도 음수 → 현재 진짜 적자(시점 불일치). 필터 문제 아님.
  - net_income_annual>0인데 탈락 → 필터/조인 버그.

실행: python scripts\\diag_peer_financials.py
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from core.cooccur import _norm_name

# 직전 검증에서 탈락한 대표 정답 피어들 (사유가 net_income이었던 것)
SUSPECTS = ["넥스틴", "SKC", "천보", "텔레칩스", "녹십자", "코윈테크", "삼익THK",
            "진시스템", "수젠텍", "아이센스", "비나텍", "영진약품", "한국비엔씨"]


def main():
    fin = pd.read_csv(config.UNIVERSE_FIN_CSV, dtype=str)
    # 회사명 컬럼 확인
    name_col = config.NAME_COL if config.NAME_COL in fin.columns else None
    uni = pd.read_csv(config.LISTED_UNIVERSE_CSV, encoding=config.CSV_ENCODING, dtype=str)
    code2name = dict(zip(uni[config.STOCK_COL], uni[config.NAME_COL]))
    name2code = {_norm_name(v): k for k, v in code2name.items()}

    cols = ["net_income_annual", "net_income_ltm", "op_income_annual", "op_income_ltm",
            "결산월", "audit_opinion", "has_financials"]
    print(f"재무 레코드: {len(fin)}개\n")
    print(f"{'회사명':<14} {'종목코드':<8} {'net_annual':>12} {'net_ltm':>12} {'결산':>5} {'감사':<8}")
    print("-" * 70)
    for name in SUSPECTS:
        code = name2code.get(_norm_name(name))
        if not code:
            print(f"{name:<14} (종목코드 못 찾음)")
            continue
        row = fin[fin[config.STOCK_COL] == code] if config.STOCK_COL in fin.columns else \
              fin[fin["종목코드"] == code]
        if not len(row):
            print(f"{name:<14} {code:<8} (재무 레코드 없음)")
            continue
        r = row.iloc[0]
        na = r.get("net_income_annual", "")
        lt = r.get("net_income_ltm", "")
        fm = r.get("결산월", "")
        au = str(r.get("audit_opinion", ""))[:6]
        print(f"{name:<14} {code:<8} {str(na):>12} {str(lt):>12} {str(fm):>5} {au:<8}")


if __name__ == "__main__":
    main()
