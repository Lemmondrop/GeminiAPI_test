"""
Phase 1: 91개 회사명 → DART corp_code 매핑

매핑 전략:
  1) OpenDartReader.corp_codes로 전체 기업 마스터 로드
  2) 회사명 정확 매칭 시도 (정식명, '(주)' 제거 후 매칭 등)
  3) finance-datareader로 상장일 검증 (있으면 더 정확)
  4) 동명 회사 다수 발견 시 사용자 확인 필요로 마킹

입력:  halla_pdf_extracted_91rows.csv  (listing_date, company_name, discount_rate_*)
출력:  ipo_with_corp_code.csv  (위 + corp_code, stock_code, match_quality, manual_check_needed)

사용법:
  python phase1_map_corp_codes.py --input halla_pdf_extracted_91rows.csv --output ipo_with_corp_code.csv
"""
import os
import re
import csv
import argparse
import time
from typing import Dict, List, Optional, Tuple
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()


# ============ 매칭 유틸리티 ============
def normalize_company_name(name: str) -> str:
    """회사명 정규화: 공백·괄호·접두사 제거"""
    if not name:
        return ""
    s = name.strip()
    # '(주)', '㈜', '주식회사' 등 접두사 제거
    s = re.sub(r'^\(주\)|^㈜|^주식회사\s*', '', s)
    # 뒤쪽 (주), ㈜
    s = re.sub(r'\(주\)$|㈜$|\s*주식회사$', '', s)
    # 모든 공백 제거 (PDF 추출 시 공백 차이 흡수)
    s = re.sub(r'\s+', '', s)
    return s.strip()


def company_name_variants(name: str) -> List[str]:
    """매칭 시도용 변형 목록 생성"""
    base = normalize_company_name(name)
    variants = {base, name.strip()}
    # 'X로보틱스' ↔ 'X Robotics'  (대소문자 무시 매칭은 별도)
    # 'X.AI', 'X에이아이' 등 변형 처리는 추후 필요시
    return list(variants)


# ============ DART corp_code 마스터 ============
def load_dart_corp_master() -> 'pd.DataFrame':
    """OpenDartReader로 전체 기업 마스터 로드.
    
    Returns:
        DataFrame with columns: corp_code, corp_name, stock_code, modify_date
    """
    api_key = os.environ.get('DART_API_BUSINESS_KEY')
    if not api_key:
        raise RuntimeError("OPENDART_API_KEY 환경변수가 설정되지 않았습니다 (.env 확인)")
    
    import OpenDartReader
    dart = OpenDartReader(api_key)
    
    print("[Phase 1.1] DART 기업 마스터 로드 중...")
    df = dart.corp_codes
    print(f"  전체 등록 기업: {len(df):,}개")
    
    # 상장사만 필터 (stock_code 있는 기업)
    listed = df[df['stock_code'].notna() & (df['stock_code'].astype(str).str.strip() != '')].copy()
    listed['stock_code'] = listed['stock_code'].astype(str).str.zfill(6)
    print(f"  상장사: {len(listed):,}개")
    
    # 정규화된 이름 컬럼 추가 (매칭 속도 ↑)
    listed['_normalized'] = listed['corp_name'].apply(normalize_company_name)
    return listed


# ============ 매칭 ============
def match_company(target_name: str, listing_date: Optional[str], 
                  dart_master: 'pd.DataFrame',
                  fdr_listing: Optional['pd.DataFrame'] = None) -> Dict:
    """단일 회사 매칭.
    
    Returns:
        {
          'corp_code': '00...',
          'stock_code': '012345',
          'matched_name': '(주)XXX',  # DART에 등록된 정식명
          'match_quality': 'exact' | 'normalized' | 'multiple' | 'not_found',
          'candidates': [...]  # 동명 후보 다수일 경우
        }
    """
    result = {
        'corp_code': None,
        'stock_code': None,
        'matched_name': None,
        'match_quality': 'not_found',
        'candidates': [],
        'notes': '',
    }
    
    if not target_name:
        return result
    
    norm_target = normalize_company_name(target_name)
    
    # 1) 정규화 이름 정확 매칭
    matches = dart_master[dart_master['_normalized'] == norm_target]
    
    if len(matches) == 1:
        row = matches.iloc[0]
        result['corp_code'] = row['corp_code']
        result['stock_code'] = row['stock_code']
        result['matched_name'] = row['corp_name']
        result['match_quality'] = 'exact'
        return result
    
    elif len(matches) > 1:
        # 동명 다수 - 상장일로 추가 필터
        result['candidates'] = [
            {'corp_code': r['corp_code'], 'name': r['corp_name'], 'stock': r['stock_code']}
            for _, r in matches.iterrows()
        ]
        
        if fdr_listing is not None and listing_date:
            # finance-datareader의 ListingDate 컬럼으로 매칭
            for _, r in matches.iterrows():
                fdr_match = fdr_listing[fdr_listing['Code'] == r['stock_code']]
                if len(fdr_match) > 0:
                    fdr_date = str(fdr_match.iloc[0].get('ListingDate', ''))[:10]
                    if fdr_date == listing_date:
                        result['corp_code'] = r['corp_code']
                        result['stock_code'] = r['stock_code']
                        result['matched_name'] = r['corp_name']
                        result['match_quality'] = 'exact_with_date'
                        return result
        
        result['match_quality'] = 'multiple'
        result['notes'] = f'{len(matches)}개 후보 - 수동 확인 필요'
        return result
    
    # 2) 부분 매칭 시도 (정규화 이름이 포함되는 회사 찾기)
    partial = dart_master[dart_master['_normalized'].str.contains(norm_target, na=False, regex=False)]
    if len(partial) >= 1:
        # 가장 짧은 이름 우선 (정확 매칭에 가까움)
        partial = partial.copy()
        partial['_len_diff'] = partial['_normalized'].str.len() - len(norm_target)
        partial = partial.sort_values('_len_diff')
        
        if len(partial) == 1:
            row = partial.iloc[0]
            result['corp_code'] = row['corp_code']
            result['stock_code'] = row['stock_code']
            result['matched_name'] = row['corp_name']
            result['match_quality'] = 'partial'
            result['notes'] = f'부분 매칭: {target_name} → {row["corp_name"]}'
            return result
        else:
            result['candidates'] = [
                {'corp_code': r['corp_code'], 'name': r['corp_name'], 'stock': r['stock_code']}
                for _, r in partial.head(5).iterrows()
            ]
            result['match_quality'] = 'multiple_partial'
            result['notes'] = f'{len(partial)}개 부분 매칭 후보'
            return result
    
    return result


# ============ FDR 상장일 데이터 (선택) ============
def load_fdr_listing_dates() -> Optional['pd.DataFrame']:
    """finance-datareader로 KOSPI+KOSDAQ 상장 기업 + 상장일 로드"""
    try:
        import FinanceDataReader as fdr
    except ImportError:
        print("⚠️  finance-datareader 미설치 (상장일 검증 스킵): pip install finance-datareader")
        return None
    
    try:
        print("[Phase 1.2] KRX 상장 기업 + 상장일 로드 중...")
        kospi = fdr.StockListing('KOSPI')
        kosdaq = fdr.StockListing('KOSDAQ')
        merged = pd.concat([kospi, kosdaq], ignore_index=True)
        merged['Code'] = merged['Code'].astype(str).str.zfill(6)
        print(f"  KOSPI+KOSDAQ: {len(merged):,}개")
        return merged
    except Exception as e:
        print(f"⚠️  FDR 로드 실패: {e}")
        return None


# ============ 메인 ============
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', required=True, help='91개 회사 CSV (halla_pdf_extracted_91rows.csv)')
    parser.add_argument('--output', required=True, help='출력 CSV')
    parser.add_argument('--no-fdr', action='store_true', help='finance-datareader 사용 안 함')
    args = parser.parse_args()
    
    import pandas as pd
    globals()['pd'] = pd  # 다른 함수에서 사용
    
    print(f"\n{'='*60}")
    print(f"Phase 1: 회사명 → corp_code 매핑")
    print(f"{'='*60}")
    
    # 입력 로드
    df = pd.read_csv(args.input, encoding='utf-8-sig')
    print(f"\n입력: {len(df)}개 회사")
    
    # DART 마스터 로드
    try:
        dart_master = load_dart_corp_master()
    except Exception as e:
        print(f"❌ DART 마스터 로드 실패: {e}")
        return
    
    # FDR 상장일 (선택)
    fdr_listing = None if args.no_fdr else load_fdr_listing_dates()
    
    # 매칭 진행
    print(f"\n[Phase 1.3] 91개 회사 매칭 중...")
    results = []
    quality_counts = {'exact': 0, 'exact_with_date': 0, 'partial': 0, 
                      'multiple': 0, 'multiple_partial': 0, 'not_found': 0}
    
    for idx, row in df.iterrows():
        company = row['company_name']
        listing_date = row['listing_date']
        
        match = match_company(company, listing_date, dart_master, fdr_listing)
        quality_counts[match['match_quality']] = quality_counts.get(match['match_quality'], 0) + 1
        
        results.append({
            'listing_date': listing_date,
            'company_name': company,
            'corp_code': match['corp_code'] or '',
            'stock_code': match['stock_code'] or '',
            'matched_name': match['matched_name'] or '',
            'match_quality': match['match_quality'],
            'manual_check_needed': match['match_quality'] in ('multiple', 'multiple_partial', 'not_found'),
            'discount_rate_low_pct': row['discount_rate_low_pct'],
            'discount_rate_high_pct': row['discount_rate_high_pct'],
            'discount_rate_mid_pct': row['discount_rate_mid_pct'],
            'notes': match['notes'],
            'candidates': str(match['candidates']) if match['candidates'] else '',
        })
        
        if (idx + 1) % 20 == 0:
            print(f"  진행: {idx+1}/{len(df)}")
    
    # 저장
    out_df = pd.DataFrame(results)
    out_df.to_csv(args.output, index=False, encoding='utf-8-sig')
    print(f"\n저장: {args.output}")
    
    # 통계
    print(f"\n{'='*60}")
    print(f"매칭 결과 통계")
    print(f"{'='*60}")
    for q, cnt in quality_counts.items():
        pct = cnt / len(df) * 100
        print(f"  {q:25s}: {cnt:3d}개 ({pct:5.1f}%)")
    
    success = quality_counts.get('exact', 0) + quality_counts.get('exact_with_date', 0) + quality_counts.get('partial', 0)
    need_check = quality_counts.get('multiple', 0) + quality_counts.get('multiple_partial', 0) + quality_counts.get('not_found', 0)
    print(f"\n  ✅ 자동 매칭 성공: {success}개")
    print(f"  ⚠️  수동 확인 필요: {need_check}개")
    
    if need_check > 0:
        print(f"\n수동 확인 필요한 회사:")
        for r in results:
            if r['manual_check_needed']:
                print(f"  - {r['company_name']} ({r['listing_date']}): {r['match_quality']} | {r['notes']}")


if __name__ == '__main__':
    main()
