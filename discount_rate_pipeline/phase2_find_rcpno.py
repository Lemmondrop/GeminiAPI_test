"""
Phase 2 v2: 증권신고서 rcpNo 검색 — 우선순위 로직 개선

변경사항 (v1 → v2):
  1. get_ipo_filing: 공시 우선순위 적용
       1순위: [발행조건확정]증권신고서(지분증권)
       2순위: 증권신고서(지분증권) (정정 포함)
       3순위: [기재정정]투자설명서 / 투자설명서 (fallback)
  2. track_type: classify_track()이 덮어쓰는 버그 수정
       → 입력 CSV의 track_type을 보존하고, 없을 때만 report_nm 기반 분류
"""
import os
import time
import argparse
import re
from typing import Dict, Optional
from datetime import datetime, timedelta

from dotenv import load_dotenv
load_dotenv()

# ── 공시 우선순위 ──────────────────────────────────────────
PRIORITY_KEYWORDS = [
    ('발행조건확정', 1),   # [발행조건확정]증권신고서 — 공모가 확정본, 최우선
    ('증권신고서',   2),   # 원본 또는 [첨부정정]증권신고서
    ('투자설명서',   3),   # [기재정정]투자설명서 — 최후 fallback
]

def filing_priority(report_nm: str) -> int:
    """낮을수록 우선순위 높음. 매칭 안 되면 99."""
    for kw, pri in PRIORITY_KEYWORDS:
        if kw in report_nm:
            return pri
    return 99


def get_ipo_filing(dart, corp_code: str, listing_date: str) -> Optional[Dict]:
    """특정 기업의 IPO 관련 증권신고서 검색.

    우선순위:
      1. [발행조건확정]증권신고서(지분증권)
      2. 증권신고서(지분증권) 계열 (정정 포함)
      3. 투자설명서 계열 (fallback)

    동일 우선순위 내에서는 rcept_dt 기준 최신 선택.
    """
    if not corp_code or not listing_date:
        return None

    try:
        listing_dt = datetime.strptime(listing_date, '%Y-%m-%d')
    except ValueError:
        return None

    start = (listing_dt - timedelta(days=180)).strftime('%Y%m%d')
    end   = (listing_dt + timedelta(days=30)).strftime('%Y%m%d')

    try:
        filings = dart.list(corp_code, start=start, end=end, kind='C')
        if filings is None or len(filings) == 0:
            return None

        col      = 'report_nm' if 'report_nm' in filings.columns else 'report_name'
        date_col = 'rcept_dt'  if 'rcept_dt'  in filings.columns else 'rcept_date'

        # 증권신고 / 투자설명서 필터
        ipo_filings = filings[
            filings[col].str.contains('증권신고|투자설명서', na=False, regex=True)
        ].copy()

        if len(ipo_filings) == 0:
            return None

        # 우선순위 컬럼 추가 후 정렬: 우선순위 오름차순, 날짜 내림차순
        ipo_filings['_pri'] = ipo_filings[col].apply(filing_priority)
        ipo_filings = ipo_filings.sort_values(['_pri', date_col], ascending=[True, False])

        best = ipo_filings.iloc[0]
        rcp_no = str(best['rcept_no'])

        return {
            'rcept_no':           rcp_no,
            'report_nm':          str(best[col]),
            'rcept_dt':           str(best[date_col]),
            'dart_url':           f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcp_no}",
            'filing_priority':    int(best['_pri']),
            'all_filings_count':  len(ipo_filings),
        }
    except Exception as e:
        return {'error': str(e)}


def classify_track_from_report(report_nm: str) -> str:
    """report_nm 기반 트랙 분류 — track_type 미지정 시에만 사용."""
    text = (report_nm or '').lower()
    if any(kw.lower() in text for kw in ['스팩', 'spac', '기업인수목적']):
        return 'spac'
    if any(kw in text for kw in ['기술특례', '기술성장기업', '기술평가']):
        return 'tech_special'
    if any(kw in text for kw in ['성장성특례', '성장성 추천']):
        return 'growth_special'
    return 'general'


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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input',          required=True)
    parser.add_argument('--output',         required=True)
    parser.add_argument('--rate-limit-sec', type=float, default=0.3)
    args = parser.parse_args()

    api_key = os.environ.get('DART_API_BUSINESS_KEY')
    if not api_key:
        print("❌ DART_API_BUSINESS_KEY 환경변수가 없습니다 (.env 확인)")
        return

    import pandas as pd
    import OpenDartReader

    print(f"\n{'='*60}")
    print(f"Phase 2 v2: 증권신고서 rcpNo 검색 (우선순위 개선)")
    print(f"{'='*60}")

    df = pd.read_csv(
        args.input,
        encoding='utf-8-sig',
        dtype={'corp_code': str, 'stock_code': str},
    )
    if 'track_type' not in df.columns:
        df['track_type'] = 'general'
    else:
        df['track_type'] = df['track_type'].fillna('').astype(str).str.strip()
        df.loc[df['track_type'].isin(['', 'nan', 'None']), 'track_type'] = 'general'
    print(f"\n입력: {len(df)}개 회사")

    has_corp = df[df['corp_code'].notna() & (df['corp_code'].astype(str).str.strip() != '')]
    print(f"corp_code 보유: {len(has_corp)}개")

    dart = OpenDartReader(api_key)

    results  = []
    stats    = {'found': 0, 'not_found': 0, 'error': 0}
    pri_dist = {1: 0, 2: 0, 3: 0, 99: 0}   # 우선순위 분포 집계

    print(f"\n[Phase 2.1] DART 공시 검색 중...")

    for i, (idx, row) in enumerate(df.iterrows()):
        corp_code    = normalize_corp_code(row.get('corp_code'))
        listing_date = row['listing_date']
        # ★ 입력 CSV의 track_type 보존
        existing_track = str(row.get('track_type', 'general')).strip() or 'general'

        out_row = dict(row)

        if not corp_code:
            out_row.update(rcept_no='', report_nm='', rcept_dt='',
                           dart_url='', filing_search_result='no_corp_code',
                           track_type=existing_track)
            results.append(out_row)
            continue

        filing = get_ipo_filing(dart, corp_code, listing_date)

        if filing is None:
            stats['not_found'] += 1
            out_row.update(rcept_no='', report_nm='', rcept_dt='',
                           dart_url='', filing_search_result='not_found',
                           track_type=existing_track)
            # track_type 보존 (없으면 빈 문자열 유지)
        elif 'error' in filing:
            stats['error'] += 1
            out_row.update(rcept_no='', report_nm='', rcept_dt='',
                           dart_url='', filing_search_result=f"error: {filing['error']}",
                           track_type=existing_track)
        else:
            stats['found'] += 1
            pri = filing['filing_priority']
            pri_dist[pri] = pri_dist.get(pri, 0) + 1

            out_row['rcept_no']           = filing['rcept_no']
            out_row['report_nm']          = filing['report_nm']
            out_row['rcept_dt']           = filing['rcept_dt']
            out_row['dart_url']           = filing['dart_url']
            out_row['filing_search_result'] = f"ok_pri{pri}"

            # ★ track_type: 입력값 보존 우선, 없으면 report_nm 기반 분류
            if existing_track and existing_track not in ('', 'nan'):
                out_row['track_type'] = existing_track
            else:
                out_row['track_type'] = classify_track_from_report(filing['report_nm'])

        results.append(out_row)

        if (i + 1) % 10 == 0:
            print(f"  진행: {i+1}/{len(df)} | "
                  f"found={stats['found']} not_found={stats['not_found']} err={stats['error']}")

        time.sleep(args.rate_limit_sec)

    out_df = pd.DataFrame(results)
    out_df.to_csv(args.output, index=False, encoding='utf-8-sig')
    print(f"\n저장: {args.output}")

    # ── 결과 요약 ──────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"rcpNo 검색 결과")
    print(f"{'='*60}")
    print(f"  발견:      {stats['found']:3d}개")
    print(f"  미발견:    {stats['not_found']:3d}개")
    print(f"  오류:      {stats['error']:3d}개")

    print(f"\n[공시 우선순위 분포]")
    labels = {1: '발행조건확정 (최우선)', 2: '증권신고서 계열', 3: '투자설명서 (fallback)', 99: '기타'}
    for pri, cnt in sorted(pri_dist.items()):
        if cnt:
            print(f"  P{pri} {labels.get(pri,'?'):25s}: {cnt:3d}개")

    print(f"\n[트랙 분포]")
    for track, cnt in out_df['track_type'].value_counts().items():
        if track:
            print(f"  {str(track):20s}: {cnt}개")

    # 학습 제외 대상 경고
    exclude = out_df[out_df['track_type'].isin(['spac', 'growth_special'])]
    if len(exclude):
        print(f"\n⚠️  학습 제외 대상 (SPAC·성장성특례): {len(exclude)}개")
        for _, r in exclude.iterrows():
            print(f"  - {r['company_name']} ({r['track_type']})")

    # konex_transfer 확인
    konex = out_df[out_df['track_type'] == 'konex_transfer']
    if len(konex):
        print(f"\n[konex_transfer 기업] — {len(konex)}개 (track_type 보존 확인)")
        for _, r in konex.iterrows():
            print(f"  ✅ {r['company_name']:15s} | {r['report_nm']} | P{str(r.get('filing_search_result',''))[-1]}")


if __name__ == '__main__':
    main()
