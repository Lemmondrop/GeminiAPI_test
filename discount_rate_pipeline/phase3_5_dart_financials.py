"""
Phase 3.5: DART 재무API로 재무 데이터 보완

XML 추출에서 누락된 재무 피처(매출액, 영업이익, 당기순이익, 자산총계, 자본총계)를
DART fnlttSinglAcnt API로 보완합니다.

전략:
  - 상장 직전 사업연도 사업보고서 우선 (reprt_code=11011)
  - 없으면 반기보고서 (11012) → 3분기보고서 (11014) 순으로 fallback
  - 연결재무제표 우선, 없으면 별도재무제표

추출 계정과목:
  - 매출액       : ifrs-full:Revenue / dart:SalesRevenue
  - 영업이익     : dart:OperatingIncomeLoss / ifrs-full:ProfitLossFromOperatingActivities
  - 당기순이익   : ifrs-full:ProfitLoss / dart:NetIncome
  - 자산총계     : ifrs-full:Assets
  - 자본총계     : ifrs-full:Equity

입력:  ipo_features_raw.csv
출력:  ipo_features_enriched.csv  (재무 피처 보완 완료)
       dart_financial_log.csv     (API 호출 로그)

사용법:
  python phase3_5_dart_financials.py --input ipo_features_raw.csv --output ipo_features_enriched.csv
  python phase3_5_dart_financials.py --input ipo_features_raw.csv --output ipo_features_enriched.csv --only-missing
"""

import os, re, sys, time, argparse, logging
import requests
import pandas as pd
import numpy as np
from dotenv import load_dotenv

load_dotenv()
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

DART_FIN_URL  = "https://opendart.fss.or.kr/api/fnlttSinglAcnt.json"
RETRY         = 3
RETRY_WAIT    = 2.0

# 보고서 우선순위: 사업보고서 → 반기 → 3분기 → 1분기
REPORT_CODES  = [
    ('11011', '사업보고서'),
    ('11012', '반기보고서'),
    ('11014', '3분기보고서'),
    ('11013', '1분기보고서'),
]

# 계정과목 매핑 (우선순위 순)
ACCOUNT_MAP = {
    'revenue': [
        'ifrs-full_Revenue',
        'dart_SalesRevenue',
        'ifrs-full_RevenueFromContractsWithCustomers',
        'dart_Revenue',
    ],
    'operating_income': [
        'dart_OperatingIncomeLoss',
        'ifrs-full_ProfitLossFromOperatingActivities',
        'dart_OperatingIncome',
    ],
    'net_income': [
        'ifrs-full_ProfitLoss',
        'dart_NetIncome',
        'ifrs-full_ProfitLossAttributableToOwnersOfParent',
        'dart_ProfitLossAttributableToOwnersOfParent',
    ],
    'total_assets': [
        'ifrs-full_Assets',
        'dart_Assets',
    ],
    'equity': [
        'ifrs-full_Equity',
        'dart_Equity',
        'ifrs-full_EquityAttributableToOwnersOfParent',
    ],
}

ACCOUNT_NAME_MAP = {
    'revenue': [
        '매출액',
        '수익(매출액)',
        '영업수익',
    ],
    'operating_income': [
        '영업이익',
        '영업이익(손실)',
    ],
    'net_income': [
        '당기순이익',
        '당기순이익(손실)',
        '연결당기순이익',
        '분기순이익',
        '반기순이익',
    ],
    'total_assets': [
        '자산총계',
    ],
    'equity': [
        '자본총계',
        '자본 총계',
    ],
}


def normalize_corp_code(value) -> str:
    """DART corp_code를 8자리 문자열로 정규화."""
    if value is None:
        return ''
    text = str(value).strip()
    if text == '' or text.lower() == 'nan':
        return ''
    if re.fullmatch(r'\d+\.0', text):
        text = text[:-2]
    text = re.sub(r'\D', '', text)
    return text.zfill(8) if text else ''


# ══════════════════════════════════════════════════════════
# 1. API 호출
# ══════════════════════════════════════════════════════════

def fetch_financials(api_key: str, corp_code: str,
                     bsns_year: int, reprt_code: str,
                     rate_limit: float) -> dict:
    """DART 단일회사 재무제표 API 호출.

    Returns:
        {'status': 'ok'/'notfound'/'error', 'data': DataFrame or None}
    """
    corp_code_padded = str(corp_code).zfill(8)
    params = {
        'crtfc_key': api_key,
        'corp_code': corp_code_padded,
        'bsns_year': str(bsns_year),
        'reprt_code': reprt_code,
    }

    for attempt in range(RETRY):
        try:
            r = requests.get(DART_FIN_URL, params=params, timeout=30)
            time.sleep(rate_limit)

            if r.status_code != 200:
                log.warning(f"HTTP {r.status_code} — {corp_code} {bsns_year} {reprt_code}")
                time.sleep(RETRY_WAIT)
                continue

            data = r.json()
            status = data.get('status', '')

            if status == '000':  # 정상
                df = pd.DataFrame(data.get('list', []))
                return {'status': 'ok', 'data': df}
            elif status == '013':  # 데이터 없음
                return {'status': 'notfound', 'data': None}
            else:
                log.warning(f"API 오류 {status}: {data.get('message')} — {corp_code}")
                return {'status': f'error_{status}', 'data': None}

        except requests.RequestException as e:
            log.warning(f"요청 실패 ({attempt+1}/{RETRY}): {e}")
            time.sleep(RETRY_WAIT)

    return {'status': 'error_timeout', 'data': None}


# ══════════════════════════════════════════════════════════
# 2. 계정과목 추출
# ══════════════════════════════════════════════════════════

def _parse_amount(val_str: str) -> float | None:
    """'-1,234,567' → -1234567.0"""
    if not val_str or str(val_str).strip() in ('', '-', 'None'):
        return None
    s = str(val_str).replace(',', '').strip()
    try:
        return float(s)
    except ValueError:
        return None


def extract_accounts(fin_df: pd.DataFrame) -> dict:
    """재무제표 DataFrame에서 핵심 계정과목 추출.

    연결재무제표(CFS) 우선, 없으면 별도(OFS).
    DART 단일회사 주요계정 API는 일부 회사에서 account_id를 비워두고
    account_nm만 제공하므로 account_id 매칭 후 계정명 매칭으로 fallback한다.
    """
    if fin_df is None or len(fin_df) == 0:
        return {}

    result = {}

    # 연결 우선 → 별도 fallback
    for fs_div in ('CFS', 'OFS'):
        sub = fin_df[fin_df['fs_div'] == fs_div] if 'fs_div' in fin_df.columns else fin_df

        for key, acct_ids in ACCOUNT_MAP.items():
            if key in result:
                continue

            candidates = []

            if 'account_id' in sub.columns:
                for acct_id in acct_ids:
                    rows = sub[sub['account_id'] == acct_id]
                    if len(rows) > 0:
                        candidates.append((rows, acct_id))

            if 'account_nm' in sub.columns:
                account_names = ACCOUNT_NAME_MAP.get(key, [])
                account_nm = sub['account_nm'].astype(str).str.strip()
                for name in account_names:
                    rows = sub[account_nm == name]
                    if len(rows) == 0:
                        rows = sub[account_nm.str.contains(re.escape(name), na=False, regex=True)]
                    if len(rows) > 0:
                        candidates.append((rows, name))

            for rows, source_name in candidates:
                # thstrm_amount: 당기, frmtrm_amount: 전기
                for amount_col in ('thstrm_amount', 'frmtrm_amount'):
                    if amount_col not in rows.columns:
                        continue
                    val = _parse_amount(rows.iloc[0][amount_col])
                    if val is not None:
                        result[key] = val
                        result[f'{key}_source'] = f'{fs_div}_{source_name}'
                        break
                if key in result:
                    break

        if len(result) >= 3:  # 핵심 3개 이상 찾으면 중단
            break

    return result


def to_billion(won: float | None) -> float | None:
    """원 → 억원 변환."""
    if won is None:
        return None
    return round(won / 1_0000_0000, 4)


# ══════════════════════════════════════════════════════════
# 3. 메인
# ══════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input',          required=True)
    parser.add_argument('--output',         required=True)
    parser.add_argument('--only-missing',   action='store_true',
                        help='net_income_bn 없는 행만 API 호출')
    parser.add_argument('--rate-limit-sec', type=float, default=0.5)
    parser.add_argument('--max-rows',       type=int,   default=None)
    args = parser.parse_args()

    api_key = (os.environ.get('DART_API_BUSINESS_KEY') or
               os.environ.get('OPENDART_API_KEY'))
    if not api_key:
        print("❌ DART_API_BUSINESS_KEY 없음"); return

    df = pd.read_csv(
        args.input,
        encoding='utf-8-sig',
        dtype={'corp_code': str, 'stock_code': str, 'rcept_no': str},
    )
    print(f"\n{'='*60}")
    print(f"Phase 3.5: DART 재무API 보완")
    print(f"{'='*60}")
    print(f"전체: {len(df)}건")

    # 대상 필터
    if args.only_missing:
        targets = df[df['net_income_bn'].isna()].copy()
        print(f"net_income 없는 행만 처리: {len(targets)}건")
    else:
        targets = df.copy()
        print(f"전체 처리: {len(targets)}건")

    if args.max_rows:
        targets = targets.head(args.max_rows)
        print(f"max_rows 제한: {args.max_rows}건")

    logs = []
    stats = {'ok': 0, 'notfound': 0, 'error': 0, 'skip': 0}

    for i, (idx, row) in enumerate(targets.iterrows()):
        company     = row['company_name']
        corp_code   = normalize_corp_code(row.get('corp_code'))
        listing_dt  = pd.to_datetime(row['listing_date'])
        listing_yr  = listing_dt.year
        listing_mo  = listing_dt.month

        # 직전 사업연도 결정
        # 상장월이 1~4월이면 전전년도도 고려 (직전 사업보고서 미제출 가능)
        candidate_years = [listing_yr - 1]
        if listing_mo <= 4:
            candidate_years.append(listing_yr - 2)

        print(f"\n[{i+1:3d}/{len(targets)}] {company} (corp={corp_code}, 상장={listing_yr}-{listing_mo:02d})")

        found = False
        log_entry = {'company_name': company, 'corp_code': corp_code,
                     'listing_date': row['listing_date']}

        if not corp_code:
            print(f"  ⚠️ corp_code 없음 - 재무API 스킵")
            log_entry['status'] = 'no_corp_code'
            stats['skip'] += 1
            logs.append(log_entry)
            continue

        for bsns_year in candidate_years:
            for reprt_code, reprt_name in REPORT_CODES:
                result = fetch_financials(
                    api_key, corp_code, bsns_year, reprt_code, args.rate_limit_sec
                )

                if result['status'] == 'ok' and result['data'] is not None:
                    accounts = extract_accounts(result['data'])

                    if accounts.get('net_income') is not None or accounts.get('revenue') is not None:
                        # 억원 변환 후 DataFrame에 반영
                        df.loc[idx, 'dart_revenue_bn']    = to_billion(accounts.get('revenue'))
                        df.loc[idx, 'dart_op_income_bn']  = to_billion(accounts.get('operating_income'))
                        df.loc[idx, 'dart_net_income_bn'] = to_billion(accounts.get('net_income'))
                        df.loc[idx, 'dart_assets_bn']     = to_billion(accounts.get('total_assets'))
                        df.loc[idx, 'dart_equity_bn']     = to_billion(accounts.get('equity'))
                        df.loc[idx, 'dart_bsns_year']     = bsns_year
                        df.loc[idx, 'dart_reprt_code']    = reprt_code

                        # 부채비율, 영업이익률 파생
                        assets = accounts.get('total_assets')
                        eq     = accounts.get('equity')
                        rev    = accounts.get('revenue')
                        op     = accounts.get('operating_income')
                        if assets and eq and eq != 0:
                            df.loc[idx, 'dart_debt_ratio'] = round((assets - eq) / eq * 100, 2)
                        if rev and op and rev != 0:
                            df.loc[idx, 'dart_op_margin'] = round(op / rev * 100, 2)

                        # net_income_bn이 비어있으면 채움
                        if pd.isna(row.get('net_income_bn')) and accounts.get('net_income') is not None:
                            df.loc[idx, 'net_income_bn'] = to_billion(accounts['net_income'])
                            df.loc[idx, 'net_income_source'] = 'dart_api'

                        ni  = to_billion(accounts.get('net_income'))
                        rev = to_billion(accounts.get('revenue'))
                        print(f"  ✅ {bsns_year} {reprt_name} | 순이익={ni}억 매출={rev}억")

                        log_entry.update({
                            'status': 'ok', 'bsns_year': bsns_year,
                            'reprt_code': reprt_code, 'reprt_name': reprt_name,
                            'net_income_bn': ni, 'revenue_bn': rev,
                        })
                        stats['ok'] += 1
                        found = True
                        break

                elif result['status'] == 'notfound':
                    pass  # 다음 보고서 코드 시도

                else:
                    log.warning(f"  {bsns_year} {reprt_name}: {result['status']}")

            if found:
                break

        if not found:
            print(f"  ❌ 재무데이터 없음")
            log_entry['status'] = 'notfound'
            stats['notfound'] += 1

        logs.append(log_entry)

        # 10건마다 중간 저장
        if (i + 1) % 10 == 0:
            df.to_csv(args.output.replace('.csv', '_partial.csv'),
                      index=False, encoding='utf-8-sig')
            print(f"\n  💾 중간 저장 ({i+1}건)")

    # 최종 저장
    df.to_csv(args.output, index=False, encoding='utf-8-sig')
    pd.DataFrame(logs).to_csv(
        args.output.replace('.csv', '_dartlog.csv'),
        index=False, encoding='utf-8-sig'
    )

    print(f"\n{'='*60}")
    print(f"Phase 3.5 완료")
    print(f"{'='*60}")
    print(f"  성공:   {stats['ok']:3d}건")
    print(f"  미발견: {stats['notfound']:3d}건")
    print(f"  오류:   {stats['error']:3d}건")
    print(f"\n저장: {args.output}")

    # 피처 커버리지 개선 확인
    print(f"\n[재무 피처 커버리지]")
    for col in ['net_income_bn', 'dart_net_income_bn', 'dart_revenue_bn',
                'dart_op_income_bn', 'dart_assets_bn', 'dart_debt_ratio']:
        if col in df.columns:
            n = df[col].notna().sum()
            pct = n / len(df) * 100
            bar = '█' * int(pct / 5) + '░' * (20 - int(pct / 5))
            print(f"  {col:25s}: {bar} {n:3d}/{len(df)} ({pct:.0f}%)")


if __name__ == '__main__':
    main()
