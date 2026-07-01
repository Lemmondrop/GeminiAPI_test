"""전체 유니버스 재무 일괄 수집 러너.

DART 재무(온기/LTM)·감사의견을 3,281사 수집 → data/cache/universe_financials.csv.
시총(KRX)은 인가된 경우에만 --mcap로 함께, 아니면 None으로 두고 나중에 백필.

체크포인트(financials_progress.json)로 중단 시 재개. 시총은 done 레코드도 백필됨.

실행 (peer_extraction 루트):
  python scripts/collect_financials.py --limit 20          # 먼저 20개로 점검
  python scripts/collect_financials.py                      # 전체 (시총 제외)
  python scripts/collect_financials.py --mcap 20260616      # KRX 인가 후: 시총까지
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from core.dart_client import DartClient
from core.financials import collect_universe


def _read_csv(path):
    try:
        return pd.read_csv(path, encoding=config.CSV_ENCODING)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="utf-8-sig")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mcap", default=None, help="시총 기준일 YYYYMMDD (KRX 인가 시)")
    ap.add_argument("--limit", type=int, default=None, help="앞 N개만 (점검용)")
    args = ap.parse_args()

    if not config.LISTED_UNIVERSE_CSV.exists():
        sys.exit(f"[중단] CSV 없음: {config.LISTED_UNIVERSE_CSV}")
    df = _read_csv(config.LISTED_UNIVERSE_CSV)

    n = len(df) if args.limit is None else min(args.limit, len(df))
    mcap_note = f"+KRX 시총({args.mcap})" if args.mcap else "시총 제외(나중 백필)"
    print(f"유니버스 {len(df)}행 → {n}건 수집  [{mcap_note}]")
    print(f"기간: 온기={config.ANNUAL_YEAR} 사업보고서 / LTM={config.LTM_QUARTER_YEAR}-{config.LTM_QUARTER_CODE}")
    print("회사당 DART 2~3콜. 중단되면 같은 명령으로 재개됨.\n")

    client = DartClient()
    out = collect_universe(df, client, mcap_date=args.mcap, limit=args.limit)

    # 요약
    ok = int(out["has_financials"].sum()) if "has_financials" in out else 0
    no_corp = int(out["corp_code"].isna().sum()) if "corp_code" in out else 0
    print(f"\n요약: 재무확보 {ok}/{len(out)}  | corp_code 매핑실패 {no_corp}")
    if args.mcap and "market_cap" in out:
        mc = int(out["market_cap"].notna().sum())
        print(f"      시총 확보 {mc}/{len(out)}")


if __name__ == "__main__":
    main()
