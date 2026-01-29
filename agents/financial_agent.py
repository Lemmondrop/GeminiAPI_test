### 재무제표, 매출 추이, 투자 이력을 추출하는 PROMPT 입니다.

import json
from utils import call_gemini, safe_json_loads

FINANCIAL_SCHEMA = """
{
  "Report_Header": {
    "Company_Name": "기업명",
    "CEO_Name": "대표자명",
    "Industry_Sector": "세부 산업분야",
    "Industry_Classification": "산업 대분류 (바이오/IT/제조/기타)",
    "Investment_Rating": "종합 등급"
  },
  "Financial_Status": {
    "Detailed_Balance_Sheet": {
       "Years": ["YYYY", "YYYY", "YYYY"],
       "Current_Assets": ["값", "값", "값"],
       "Non_Current_Assets": ["값", "값", "값"],
       "Total_Assets": ["값", "값", "값"],
       "Current_Liabilities": ["값", "값", "값"],
       "Non_Current_Liabilities": ["값", "값", "값"],
       "Total_Liabilities": ["값", "값", "값"],
       "Capital_Stock": ["값", "값", "값"],
       "Retained_Earnings_Etc": ["값", "값", "값"],
       "Total_Equity": ["값", "값", "값"]
    },
    "Income_Statement_Summary": {
       "Years": ["YYYY", "YYYY", "YYYY"],
       "Total_Revenue": ["값", "값", "값"],
       "Operating_Profit": ["값", "값", "값"],
       "Net_Profit": ["값", "값", "값"]
    },
    "Key_Financial_Commentary": "재무 실적 상세 분석 (매출 변동 원인, 비용 구조, 향후 실적 추정 근거 등)",
    "Investment_History": [
      { "Date": "YYYY.MM", "Round": "Series A", "Amount": "금액(단위)", "Investor": "투자자" }
    ],
    "Future_Revenue_Structure": {
        "Business_Model": "수익 모델 상세",
        "Future_Cash_Cow": "미래 핵심 수익원"
    }
  },
  "Investment_Thesis_Summary": "핵심 투자 하이라이트 요약"
}
"""

def analyze(pdf_path: str) -> dict:
    print(f"   [Financial Agent] 재무 데이터 및 헤더 정보 분석 중...")
    
    prompt = f"""
    당신은 기업 재무 분석 전문가(CFA)입니다.
    제공된 IR 자료(PDF)에서 **기본 기업 정보(Header)**와 **재무 데이터(Financial)**를 정밀하게 추출하십시오.

    [필수 추출 항목]
    1. **Report_Header**: 기업명, 대표자명, 산업분야를 정확히 추출해야 다른 에이전트가 작동합니다.
    2. **재무제표**: 모든 숫자는 단위(백만원, 억원 등)를 명시하십시오. 표가 없으면 텍스트에서 유추하십시오.
    3. **투자 이력**: 기존 투자 유치 이력을 날짜순으로 정리하십시오.

    [Output Schema]
    {FINANCIAL_SCHEMA}
    """
    
    res = call_gemini(prompt, pdf_path=pdf_path)
    if res.get("ok"):
        return safe_json_loads(res["text"])
    else:
        print(f"   [Financial Agent] Error: {res.get('error')}")
        return {}