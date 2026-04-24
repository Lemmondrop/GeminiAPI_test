import json
from pydantic import BaseModel, Field
from typing import List, Optional
from utils import call_gemini, safe_json_loads

# =========================================================================
# 1. [Pydantic 스키마 정의] - FINANCIAL_SCHEMA 문자열을 객체지향적으로 완벽 대체
# =========================================================================
class ReportHeader(BaseModel):
    Company_Name: str = Field(description="기업명")
    CEO_Name: str = Field(description="대표자명")
    Industry_Sector: str = Field(description="세부 산업분야")
    Industry_Classification: str = Field(description="산업 대분류 (바이오/IT/제조/기타 중 택 1)")
    Investment_Rating: str = Field(description="종합 투자 등급 (긍정적/관심/관망/부정적 중 택 1)")

class BalanceSheet(BaseModel):
    Unit: Optional[str] = Field(None, description="표기된 단위 (예: 단위 : 백만원, 억원). 없으면 null")
    Columns: List[str] = Field(description="연도 헤더 (예: ['구분', '2022', '2023']). 데이터 없으면 빈 배열 []")
    Rows: List[List[str]] = Field(description="계정별 실적 데이터 배열. 데이터 없으면 빈 배열 []")

class IncomeStatement(BaseModel):
    Unit: Optional[str] = Field(None, description="표기된 단위 (예: 단위 : 백만원, 억원). 없으면 null")
    Columns: List[str] = Field(description="연도 헤더 (예: ['구분', '2025(E)', '2026(E)']). 데이터 없으면 빈 배열 []")
    Rows: List[List[str]] = Field(description="계정별 추정손익 데이터 배열. 데이터 없으면 빈 배열 []")

class InvestmentHistoryItem(BaseModel):
    Date: Optional[str] = Field(None, description="투자 일자 (예: 2024.01). 🚨명확하지 않다면 '2000.00' 등 가짜 숫자를 짓지 말고 무조건 빈칸(\"\") 또는 null 처리")
    Round: str = Field(description="투자 라운드 (과거 이력뿐만 아니라 '현재 투자 라운드'나 '목표 라운드'도 무조건 포함. 예: Seed + Angel)")
    Amount: str = Field(description="투자 금액 또는 목표 투자 금액 (예: 10억)")
    Investor: str = Field(description="투자자 (문서에 없으면 '정보 부재' 또는 '미공개')")

class FutureRevenueStructure(BaseModel):
    Business_Model: str = Field(description="수익 모델 상세")
    Future_Cash_Cow: str = Field(description="미래 핵심 수익원")

class HighlightItem(BaseModel):
    Highlight_Title: str = Field(description="투자 매력도를 요약하는 1줄 소제목 (반드시 대괄호 [] 로 시작할 것. 예: [독점적 기술 해자])")
    Highlight_Logic: str = Field(description="해당 소제목을 뒷받침하는 객관적 수치, 파트너십, 시장 우위 등 VC 심사역 관점의 근거 (2문장 이내)")

class FinancialStatus(BaseModel):
    Balance_Sheet: BalanceSheet
    Income_Statement: IncomeStatement
    Key_Financial_Commentary: str = Field(description="재무 실적 상세 분석 (매출 변동 원인, 비용 구조, 향후 실적 추정 근거 등)")
    Investment_History: List[InvestmentHistoryItem] = Field(description="투자 유치 이력 배열. 과거 투자가 없고 현재 라운드(펀딩 계획)만 있어도 무조건 포함할 것.")
    Future_Revenue_Structure: FutureRevenueStructure

class FinancialResponseSchema(BaseModel):
    Report_Header: ReportHeader
    Financial_Status: FinancialStatus
    Investment_Highlights: List[HighlightItem] = Field(description="이 회사에 투자해야 하는 가장 강력한 이유 3가지", min_length=3, max_length=3)

# =========================================================================
# 2. [분석 엔진(Agent) 실행부]
# =========================================================================
def analyze(pdf_path: str, extra_text: str = "") -> dict:
    print(f"   [Financial Agent] 재무 데이터 및 헤더 정보 분석 중...")
    
    # 🚨 기존 프롬프트의 지시사항은 그대로 유지하되, 스키마 설명은 Pydantic이 대체하므로 제거
    prompt = f"""
    당신은 기업 재무 분석 전문가(CFA)이자 벤처캐피탈(VC)의 시니어 심사역입니다.
    제공된 메인 IR 자료(PDF)와 **[보충 문서 데이터]**(엑셀 재무제표, Markdown 등)를 종합하여 정밀하게 추출하십시오.
    엑셀(표 형태) 또는 Markdown으로 제공된 재무 데이터가 있다면 이를 최우선으로 신뢰하십시오.

    [보충 문서 데이터]
    {extra_text}

    [필수 추출 항목 및 중요 작성 지침]
    1. Report_Header: 기업명, 대표자명, 산업분야를 정확히 추출해야 다른 에이전트가 작동합니다.
       - Investment_Rating: 사업 계획과 기술력을 종합적으로 판단하여 기초 투자 등급을 기재하십시오.
    2. 🚨 [Executive Summary 개편 (매우 중요)]: Investment_Highlights 필드는 이 기업에 투자해야 하는 가장 강력한 논리 3가지를 도출하는 곳입니다.
       - '이 회사는 ~하는 업체입니다' 식의 단순 기업 소개나 설명문 작성을 절대 금지합니다.
       - VC 투심위원을 설득할 수 있도록 [독점적 기술력], [안정적 전방/고객사 확보], [압도적 수익성/마진], [규제 수혜] 등 강력한 키워드를 Highlight_Title에 적으십시오.
       - Highlight_Logic에는 그 주장을 뒷받침하는 객관적인 사실(특허 개수, 레퍼런스 기업명, 추정 매출액 등)을 1~2문장으로 날카롭게 적으십시오.
    3. 🚨 [표 데이터 단위 제거 (매우 중요)]: 전체 표의 기준 단위(백만원, 억원 등)는 오직 Unit 필드에만 명시하십시오. Balance_Sheet 및 Income_Statement의 Rows 데이터(값)를 추출할 때, 숫자 뒤에 '억원', '백만원', '%' 등의 문자열 단위를 절대 포함하지 마십시오. 오직 순수 숫자 형태(예: "183.73", "1,040", "-20")로만 기재해야 워드 변환기에서 에러가 발생하지 않습니다.
    4. 투자 이력: 기존 투자 유치 이력을 날짜순으로 정리하십시오.
    5. 🚨 [추정손익 표 추출]: 보충 문서(Markdown 등)의 후반부를 꼼꼼히 탐색하여 '5개년 추정손익', '예상 실적', 'Financial Projection' 관련 표가 존재하는 경우에만 추출하십시오.
    6. 🚨 [Hallucination 방지]: 만약 문서 내에 미래 추정 실적 데이터가 명시되어 있지 않다면, 절대 임의의 숫자를 계산하거나 지어내지 마십시오. 이 경우 Income_Statement의 Columns와 Rows 필드를 빈 배열([])로 남겨두십시오.
    7. 표 데이터 전사: 찾은 데이터가 있을 경우에만 Income_Statement 필드의 Columns(연도 헤더)와 Rows(매출액/영업수익, 영업이익, 당기순이익 등 핵심 계정) 배열에 2차원 표 형태로 정확하게 전사하십시오. 과거 재무 데이터 역시 존재할 경우에만 Balance_Sheet 배열에 작성하고, 없다면 빈 배열([])로 두십시오.
    8. 투자 일자(Date)가 문서에 명확히 기재되어 있지 않다면, 절대로 '2000.00', '0000', 'N/A' 등의 임의의 숫자나 문자를 지어내서 채우지 마십시오. 반드시 값을 빈 문자열("") 또는 null로 비워두어야 합니다.
    """
    
    # 🚨 call_gemini 함수에 Pydantic 스키마(response_schema)를 전달합니다.
    # (utils.py의 call_gemini 함수가 response_schema kwargs를 지원해야 완벽하게 작동합니다)
    res = call_gemini(prompt, pdf_path=pdf_path, response_schema=FinancialResponseSchema)
    
    if res.get("ok"):
        return safe_json_loads(res["text"])
    else:
        print(f"   [Financial Agent] Error: {res.get('error')}")
        return {}