"""
Phase 1.5: Phase 1 결과의 수동 확인 회사 8개 일괄 보강

처리 내역:
  - 노이즈 행 3개 제거 (보통주 전환, 액면분할, 제3자배정 유상증자)
  - 사명 변경 2개 corp_code 매핑 (인스웨이브시스템즈, 삼기이브이)
  - 코넥스 이전상장 2개 매핑 + 트랙 표시 (빅텐츠, 시큐센)
  - 동명회사 1개 후보 선택 (위너스 → 위너스일렉)
  - track_type 컬럼 추가 (general / konex_transfer)

종목코드 → corp_code 변환은 DART 마스터에서 자동 조회.

입력:  ipo_with_corp_code.csv      (Phase 1 출력)
출력:  ipo_with_corp_code_v2.csv   (보강 완료)

사용법:
  python phase1_5_fix_manual_cases.py --input ipo_with_corp_code.csv --output ipo_with_corp_code_v2.csv
"""
import os
import argparse
from dotenv import load_dotenv
load_dotenv()


# ============ 수동 매핑 테이블 ============
# 종목코드(stock_code) 기준 매핑 - corp_code는 DART 마스터에서 자동 조회
MANUAL_MAPPING = {
    # PDF 추출명 -> {stock_code, dart_name, track_type, note}
    '인스웨이브시스템즈': {
        'stock_code': '450520',
        'expected_dart_name': '인스웨이브',
        'track_type': 'general',
        'note': '사명 변경: 인스웨이브시스템즈 → 인스웨이브',
    },
    '빅텐츠': {
        'stock_code': '210120',
        'expected_dart_name': '빅토리콘텐츠',
        'track_type': 'konex_transfer',
        'note': '코넥스→코스닥 이전상장. 정식명: 빅토리콘텐츠',
    },
    '시큐센': {
        'stock_code': '232830',
        'expected_dart_name': '시큐센',
        'track_type': 'konex_transfer',
        'note': '코넥스→코스닥 이전상장',
    },
    '삼기이브이': {
        'stock_code': '419050',
        'expected_dart_name': '삼기에너지솔루션즈',
        'track_type': 'general',
        'note': '사명 변경: 삼기이브이 → 삼기에너지솔루션즈',
    },
    '위너스': {
        'stock_code': '479960',
        'expected_dart_name': '위너스일렉',
        'track_type': 'general',
        'note': '동명회사 다수 → 위너스일렉(479960) 선택 (FN가이드 확인)',
    },
}

# 제거할 노이즈 행 (회사명 기준)
NOISE_ROWS = {'보통주 전환', '액면분할', '제3자배정 유상증자'}


def load_dart_master():
    """DART 기업 마스터 로드 (stock_code → corp_code 변환용)"""
    api_key = os.environ.get('DART_API_BUSINESS_KEY')
    if not api_key:
        raise RuntimeError("OPENDART_API_KEY 환경변수가 없습니다 (.env 확인)")
    
    import OpenDartReader
    dart = OpenDartReader(api_key)
    print("[Phase 1.5.1] DART 기업 마스터 로드 중...")
    df = dart.corp_codes
    df['stock_code'] = df['stock_code'].astype(str).str.strip()
    return df


def lookup_corp_code(stock_code: str, dart_master) -> dict:
    """종목코드로 corp_code 역조회"""
    # 6자리 zero-padding
    sc = str(stock_code).zfill(6)
    matches = dart_master[dart_master['stock_code'] == sc]
    if len(matches) == 0:
        return None
    row = matches.iloc[0]
    return {
        'corp_code': row['corp_code'],
        'corp_name': row['corp_name'],
        'stock_code': sc,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', required=True, help='Phase 1 출력 CSV')
    parser.add_argument('--output', required=True, help='보강 후 출력 CSV')
    args = parser.parse_args()
    
    import pandas as pd
    
    print(f"\n{'='*60}")
    print(f"Phase 1.5: 수동 확인 회사 일괄 보강")
    print(f"{'='*60}")
    
    # 입력 로드
    df = pd.read_csv(args.input, encoding='utf-8-sig', dtype={'stock_code': str, 'corp_code': str})
    print(f"\n입력: {len(df)}개 행")
    
    # track_type 컬럼이 없으면 기본값 'general'로 추가
    if 'track_type' not in df.columns:
        df['track_type'] = 'general'
    
    # 1) 노이즈 행 제거
    before_count = len(df)
    df = df[~df['company_name'].isin(NOISE_ROWS)].copy()
    removed = before_count - len(df)
    print(f"\n[1] 노이즈 행 제거: {removed}개 제거")
    for noise in NOISE_ROWS:
        print(f"    - {noise}")
    
    # 2) DART 마스터 로드 (수동 매핑에 사용)
    try:
        dart_master = load_dart_master()
    except Exception as e:
        print(f"\n❌ DART 마스터 로드 실패: {e}")
        return
    
    # 3) 수동 매핑 적용
    print(f"\n[2] 수동 매핑 적용:")
    update_count = 0
    for pdf_name, mapping in MANUAL_MAPPING.items():
        idx_match = df['company_name'] == pdf_name
        if not idx_match.any():
            print(f"    ⚠️  {pdf_name}: 입력에 없음 (스킵)")
            continue
        
        # corp_code 역조회
        lookup = lookup_corp_code(mapping['stock_code'], dart_master)
        if lookup is None:
            print(f"    ❌ {pdf_name}: stock_code {mapping['stock_code']}로 corp_code 조회 실패")
            continue
        
        # 업데이트
        df.loc[idx_match, 'corp_code'] = lookup['corp_code']
        df.loc[idx_match, 'stock_code'] = lookup['stock_code']
        df.loc[idx_match, 'matched_name'] = lookup['corp_name']
        df.loc[idx_match, 'match_quality'] = 'manual'
        df.loc[idx_match, 'manual_check_needed'] = False
        df.loc[idx_match, 'track_type'] = mapping['track_type']
        df.loc[idx_match, 'notes'] = mapping['note']
        
        print(f"    ✅ {pdf_name:20s} → corp_code {lookup['corp_code']} ({lookup['corp_name']}) [{mapping['track_type']}]")
        update_count += 1
    
    print(f"\n  총 {update_count}개 수동 매핑 완료")
    
    # 4) 통계
    print(f"\n[3] 최종 통계:")
    print(f"  전체 행: {len(df)}개")
    
    quality_counts = df['match_quality'].value_counts()
    print(f"\n  매칭 품질:")
    for q, cnt in quality_counts.items():
        print(f"    {q:20s}: {cnt}개")
    
    track_counts = df['track_type'].value_counts()
    print(f"\n  트랙 분포:")
    for t, cnt in track_counts.items():
        print(f"    {t:20s}: {cnt}개")
    
    # 여전히 확인 필요한 행
    still_need = df[df['manual_check_needed'] == True]
    if len(still_need) > 0:
        print(f"\n  ⚠️ 여전히 확인 필요: {len(still_need)}개")
        for _, r in still_need.iterrows():
            print(f"    - {r['company_name']} ({r['listing_date']}): {r['match_quality']}")
    else:
        print(f"\n  ✅ 모든 행 매핑 완료 (수동 확인 필요 0개)")
    
    # 저장
    df.to_csv(args.output, index=False, encoding='utf-8-sig')
    print(f"\n저장: {args.output}")


if __name__ == '__main__':
    main()
