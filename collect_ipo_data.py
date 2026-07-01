"""
할인율 모델용 IPO 데이터 수집 파이프라인 (Phase A + B)

사용 방법:
1. https://opendart.fss.or.kr 에서 API 키 무료 발급
2. .env 파일에 DART_API_BUSINESS_KEY=xxx 저장 또는 환경변수 설정
3. python collect_ipo_data.py --start 2023-01-01 --end 2025-12-31 --output ipo_meta.csv

Phase A: 2023~2025 IPO 목록 수집 (KRX KIND + DART 매칭)
Phase B: 각 IPO의 증권신고서 rcpNo + 메타데이터 수집
Phase C: (별도 스크립트) PDF 본문에서 평가가액·할인율 추출
"""
import os
import time
import json
import argparse
import pandas as pd
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# OpenDART 라이브러리
import OpenDartReader

load_dotenv()

# ============ 설정 ============
BASE_URL = "https://opendart.fss.or.kr/api"
API_KEY = (os.getenv("DART_API_BUSINESS_KEY") or "").strip()
OUTPUT_DIR = Path('./data')
OUTPUT_DIR.mkdir(exist_ok=True)

# IPO 관련 공시 코드
# pblntf_detail_ty:
#   - C001: 증권신고서(지분증권)  ← IPO
#   - C002: 증권신고서(채무증권)
#   - C004: 증권발행실적보고서

# 트랙 분류 키워드 (증권신고서 본문/제목에서 찾음)
TRACK_KEYWORDS = {
    'tech_special': ['기술특례', '기술성장기업', '기술평가'],
    'growth_special': ['성장성특례', '성장성 추천'],
    'spac': ['스팩', 'SPAC', '기업인수목적'],
    'general': [],  # 위에 해당 안 되면 일반
}

# ============ Phase A: IPO 목록 수집 ============
def collect_ipo_list(start_date, end_date, output_path):
    """KRX KIND에서 신규상장 기업 목록 다운로드.
    
    실제로는 KRX KIND 페이지(https://kind.krx.co.kr/corpgeneral/corpList.do)에서
    상장공시 검색 → '신규상장' 필터 → 엑셀 다운로드가 가장 빠름.
    여기서는 OpenDART 공시검색을 활용하는 방법을 보여줌.
    """
    print(f"[Phase A] IPO 목록 수집: {start_date} ~ {end_date}")
    
    if API_KEY == 'YOUR_API_KEY_HERE':
        print("⚠️  OPENDART_API_KEY 환경변수가 설정되지 않았습니다.")
        print("    https://opendart.fss.or.kr 에서 무료 발급 후 설정하세요.")
        return None
    
    dart = OpenDartReader(API_KEY)
    
    # OpenDART list_date_ex로 기간 내 모든 공시 조회 후 신규상장 필터
    # 단, 신규상장은 자료실 카테고리(C)가 아니라 거래소공시(I)에 있음
    # 가장 확실한 방법: KRX KIND에서 직접 다운로드 후 corp_code 매핑
    
    # 여기서는 간단화: corp_code 매핑 테이블을 받아서, 신규 상장 종목 필터링
    try:
        # 전체 기업 리스트 (KRX 상장사)
        corp_list = dart.corp_codes
        print(f"  전체 등록 기업: {len(corp_list)}개")
        
        # 종목코드 있는 기업만 (상장사)
        listed = corp_list[corp_list['stock_code'].notna() & (corp_list['stock_code'] != '')]
        print(f"  상장사: {len(listed)}개")
        
        # ⚠️ DART에는 "최초 상장일"이 직접 노출되지 않음
        # 실제로는 KRX KIND 또는 finance-datareader 같은 라이브러리로 보강 필요
        print("\n  ⚠️ 다음 단계: KRX KIND 신규상장 공시에서 상장일 매핑 필요")
        print("  추천: pip install finance-datareader → fdr.StockListing('KOSDAQ')")
        
        return listed
        
    except Exception as e:
        print(f"  ❌ 오류: {e}")
        return None

# ============ Phase B: 증권신고서 메타데이터 수집 ============
def get_ipo_filing_metadata(corp_code, listing_date, dart):
    """특정 기업의 IPO 관련 증권신고서 정보 수집"""
    # 상장일 6개월 전부터 1개월 후까지 조회 (정정신고서 포함)
    listing_dt = datetime.strptime(listing_date, '%Y-%m-%d')
    start = (listing_dt - pd.Timedelta(days=180)).strftime('%Y-%m-%d')
    end = (listing_dt + pd.Timedelta(days=30)).strftime('%Y-%m-%d')
    
    try:
        # 해당 기업의 기간 내 공시 조회
        filings = dart.list(corp_code, start=start, end=end, kind='C')  # C: 증권신고
        
        if filings is None or len(filings) == 0:
            return None
        
        # 증권신고서(지분증권) 필터: 보고서명에 '증권신고서(지분증권)' 또는 '투자설명서'
        ipo_filings = filings[
            filings['report_nm'].str.contains('증권신고|투자설명서', na=False, regex=True)
        ].copy()
        
        if len(ipo_filings) == 0:
            return None
        
        # 가장 최근(정정 반영) 신고서 선택
        ipo_filings = ipo_filings.sort_values('rcept_dt', ascending=False)
        latest = ipo_filings.iloc[0]
        
        return {
            'rcept_no': latest['rcept_no'],
            'report_nm': latest['report_nm'],
            'rcept_dt': latest['rcept_dt'],
            'dart_url': f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={latest['rcept_no']}",
            'all_filings_count': len(ipo_filings),
        }
        
    except Exception as e:
        print(f"  ❌ {corp_code} 오류: {e}")
        return None

def classify_track(report_nm, business_desc=""):
    """상장 트랙 분류 (일반 / 기술특례 / 성장성특례 / SPAC)"""
    text = (report_nm + " " + business_desc).lower()
    
    for keyword in TRACK_KEYWORDS['spac']:
        if keyword.lower() in text:
            return 'spac'
    for keyword in TRACK_KEYWORDS['tech_special']:
        if keyword in text:
            return 'tech_special'
    for keyword in TRACK_KEYWORDS['growth_special']:
        if keyword in text:
            return 'growth_special'
    return 'general'

# ============ Phase B-2: 재무 데이터 수집 ============
def get_financial_data(corp_code, year, dart):
    """OpenDART의 fnlttSinglAcnt API로 재무 데이터 수집"""
    try:
        # 최근 사업연도 재무제표 (CFS: 연결재무제표)
        fs = dart.finstate(corp_code, year, reprt_code='11011')  # 11011 = 사업보고서
        
        if fs is None or len(fs) == 0:
            return None
        
        # 핵심 지표 추출
        result = {}
        for _, row in fs.iterrows():
            acc = row['account_nm']
            if acc == '당기순이익':
                result['net_income'] = row.get('thstrm_amount', '')
            elif acc == '매출액' or acc == '영업수익':
                result['revenue'] = row.get('thstrm_amount', '')
            elif acc == '영업이익':
                result['operating_income'] = row.get('thstrm_amount', '')
            elif acc == '자산총계':
                result['total_assets'] = row.get('thstrm_amount', '')
            elif acc == '부채총계':
                result['total_liabilities'] = row.get('thstrm_amount', '')
        
        return result
    except Exception as e:
        return None

# ============ Phase C 예고: PDF 표 추출 ============
def extract_valuation_from_pdf(rcpNo, save_path=None):
    """증권신고서 PDF에서 평가가액·할인율 표 추출 (별도 스크립트로 분리 예정)
    
    워크플로우:
    1. DART 문서뷰어 페이지에서 PDF 첨부파일 URL 추출
       URL 패턴: https://dart.fss.or.kr/pdf/download/main.do?rcp_no={rcpNo}
    2. PDF 다운로드 → pdfplumber로 표 추출
    3. '주당 평가가액 산출' / '희망공모가액' 키워드로 해당 페이지 식별
    4. 표 셀 파싱 → 한라캐스트 케이스의 9,910 / 25.93 / 7,033 / 17.54~27.49 형태로 정규화
    5. 실패 시: 페이지 이미지화 → Gemini Vision으로 표 읽기
    """
    print(f"[Phase C] {rcpNo} - PDF 추출은 별도 스크립트에서 구현")
    pass

# ============ 메인 ============
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--start', default='2023-01-01', help='시작 날짜')
    parser.add_argument('--end', default='2025-12-31', help='종료 날짜')
    parser.add_argument('--output', default='./data/ipo_metadata.csv', help='출력 파일')
    parser.add_argument('--limit', type=int, default=None, help='수집 건수 제한 (테스트용)')
    args = parser.parse_args()
    
    print(f"\n{'='*60}")
    print(f"할인율 모델용 IPO 데이터 수집 파이프라인")
    print(f"{'='*60}\n")
    
    # Phase A: 상장 기업 목록
    listed = collect_ipo_list(args.start, args.end, args.output)
    if listed is None:
        return
    
    if API_KEY == 'YOUR_API_KEY_HERE':
        print("\n계속하려면 API 키를 설정하고 다시 실행하세요.")
        return
    
    # ⚠️ 이 시점에서 listed 데이터프레임에는 '상장일'이 없음
    # 실제 사용 시에는 다음 중 하나 필요:
    #   A) finance-datareader: fdr.StockListing('KOSDAQ')로 ListingDate 컬럼 받기
    #   B) KRX KIND 신규상장 공시 엑셀 다운로드
    
    print("\n[중요] 이 스크립트는 스켈레톤입니다. 실제 사용 시:")
    print("  1) OPENDART_API_KEY 환경변수 설정")
    print("  2) finance-datareader 또는 KRX KIND로 상장일 보강")
    print("  3) Phase C (PDF 추출)는 extract_pdf_tables.py에서 처리")

if __name__ == '__main__':
    main()