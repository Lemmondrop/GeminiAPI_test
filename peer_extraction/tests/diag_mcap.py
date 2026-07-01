"""시총 소스 진단 — 전체 backfill 전에 어느 소스가 동작하는지 빠르게 확인.

실행: peer_extraction 루트에서  python tests\\diag_mcap.py [YYYYMMDD]
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.financials import market_caps

SAMPLE = {"125490": "한라캐스트", "001780": "알루코", "047310": "파워로직스"}


def main():
    date = sys.argv[1] if len(sys.argv) > 1 else "20260619"
    print(f"시총 수집 시도 (기준일 {date})...")
    caps = market_caps(date)
    print(f"\n확보: {len(caps)}종목")
    if caps:
        for code, name in SAMPLE.items():
            v = caps.get(code)
            s = f"{v/1e8:,.0f}억" if isinstance(v, (int, float)) else "—(없음)"
            print(f"  {name}({code}): {s}")
        print("\n→ 정상. scripts\\collect_financials.py --mcap 로 전체 backfill 가능.")
    else:
        print("→ 모든 소스 실패. pip install finance-datareader 확인,")
        print("  또는 KRX 공식 활용신청 승인 대기. 위 경고 메시지 참고.")


if __name__ == "__main__":
    main()
