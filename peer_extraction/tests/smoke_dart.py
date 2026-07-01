"""DART 라이브 스모크 — 3개사로 실제 응답 필드·추출을 검증.

전체 3,281사 수집(collect_universe) 전에, 알려진 회사로 다음을 눈으로 확인:
  - corp_code 매핑(종목코드→고유번호)이 되는가
  - 영업이익/당기순이익 온기·LTM 값이 그럴듯한가 (음수/적자도 정상)
  - 감사의견(adt_opinion)·감사인 필드가 실제로 채워지는가
  - 연결(CFS) 우선 선택이 맞는가
  - KRX 시가총액이 종목코드로 조인되는가

선행: 상위 .env에 DART_API_BUSINESS_KEY=... 와 KRX_API_BUSINESS_KEY=...  /  pip install -r requirements.txt
실행(peer_extraction 루트): python tests/smoke_dart.py
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from core.dart_client import DartClient
from core.financials import collect_one, market_caps

NAMES = ["한라캐스트", "알루코", "파워로직스"]
MCAP_DATE = "20260616"   # 최근 영업일 (휴장일이면 직전 영업일로 조정)


def _read_csv(path):
    try:
        return pd.read_csv(path, encoding=config.CSV_ENCODING)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="utf-8-sig")


def main():
    if not config.LISTED_UNIVERSE_CSV.exists():
        sys.exit(f"[중단] CSV 없음: {config.LISTED_UNIVERSE_CSV}")
    df = _read_csv(config.LISTED_UNIVERSE_CSV)

    client = DartClient()
    print("corp_code 매핑 다운로드/로드 중...")
    code_map = client.corp_code_map()
    print(f"  매핑 {len(code_map)}건 준비")

    caps = {}
    try:
        caps = market_caps(MCAP_DATE)
        print(f"  KRX 시총 {len(caps)}종목 ({MCAP_DATE})")
    except Exception as e:
        print(f"  [경고] KRX 시총 실패: {e}  (시총 None으로 진행)")

    print(f"\n기간: 온기={config.ANNUAL_YEAR} 사업보고서 / "
          f"LTM={config.LTM_QUARTER_YEAR}-{config.LTM_QUARTER_CODE}\n")

    for nm in NAMES:
        row = df[df[config.NAME_COL] == nm]
        if row.empty:
            print(f"=== {nm}: CSV에 없음 ===\n"); continue
        sc = str(row.iloc[0][config.STOCK_COL]).split(".")[0].zfill(6)
        cc = code_map.get(sc)
        print(f"=== {nm}  (종목={sc}  corp_code={cc}) ===")
        if not cc:
            print("  corp_code 매핑 실패\n"); continue
        m = collect_one(client, cc)
        u = 1e8  # 억원 환산 (DART 단위 원)
        def f(v): return f"{v/u:,.0f}억" if isinstance(v, (int, float)) else "—"
        print(f"  매출액    온기={f(m['revenue_annual'])}   LTM={f(m['revenue_ltm'])}")
        print(f"  영업이익  온기={f(m['op_income_annual'])}   LTM={f(m['op_income_ltm'])}")
        print(f"  당기순익  온기={f(m['net_income_annual'])}   LTM={f(m['net_income_ltm'])}")
        print(f"  감사의견  {m['audit_opinion']}  (감사인 {m['auditor']})")
        print(f"  FS기준    {m.get('net_income_fs')}   시가총액={f(caps.get(sc))}\n")

    print("점검: 영업이익/순이익 온기·LTM 값이 공시와 맞는지, 감사의견이 '적정'으로 뜨는지 확인.")


if __name__ == "__main__":
    main()