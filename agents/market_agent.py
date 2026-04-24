import json
from pydantic import BaseModel, Field
from typing import List, Optional
from utils import call_gemini, safe_json_loads

# =========================================================================
# 1. [Pydantic 스키마 정의] - 기존 MARKET_SCHEMA 문자열을 완벽하게 대체
# =========================================================================
class TargetMarketAnalysis(BaseModel):
    TAM: str = Field(description="TAM (전체 시장)에 대한 서술 (숫자 또는 명확한 시장 정의)")
    SAM: str = Field(description="SAM (유효 시장)에 대한 서술")
    SOM: str = Field(description="SOM (수익 시장)에 대한 서술")

class CompetitorItem(BaseModel):
    Company: str = Field(description="Target 기업 및 경쟁사명 (배열의 첫 번째는 반드시 Target 기업)")
    Product: str = Field(description="주요 제품 및 서비스")
    Target_Market: str = Field(description="주요 타겟 고객/시장")
    Core_Tech: str = Field(description="핵심 기술력 및 차별점 (강점)")

class MarketTrendItem(BaseModel):
    Type: str = Field(description="트렌드 분류 (예: 뉴스/보고서, 기술동향, 규제)")
    Source: str = Field(description="출처")
    Content: str = Field(description="트렌드 상세 내용 (3문장 이상)")

class ExpectedLOScenario(BaseModel):
    Category: str = Field(description="M&A 또는 IPO")
    Probability: str = Field(description="높음/보통/낮음 중 택 1")
    Comment: str = Field(description="M&A 관점에서의 구체적인 타겟 기업(대기업/PEF 등) 및 인수 매력도 분석, 또는 IPO 관점에서의 상장 가능성 및 투자 매력도 분석")

class LOExitStrategy(BaseModel):
    Verified_Signals: List[str] = Field(description="기존 레퍼런스, 파트너십 목록")
    Expected_LO_Scenarios: List[ExpectedLOScenario] = Field(description="M&A와 IPO 시나리오를 반드시 각각 1개씩 분리하여 배열에 담을 것")
    Valuation_Range: str = Field(description="시장 관점에서의 가치 평가 (적정 기업 가치 밴드 등)")

class ExportAndContractStats(BaseModel):
    Export_Graph_Data: List[List[str]] = Field(description="수출 실적 연도와 수치 쌍 2D 배열 (예: [['2023', '100']]). 없으면 빈 배열 []")
    Contract_Count_Graph_Data: List[List[str]] = Field(description="계약 건수 연도와 수치 쌍 2D 배열. 없으면 빈 배열 []")
    Sales_Graph_Data: List[List[str]] = Field(description="매출 규모 추이 연도와 수치 쌍 2D 배열. 없으면 빈 배열 []")

class GrowthPotential(BaseModel):
    Target_Market_Analysis: TargetMarketAnalysis
    Competitors_Comparison: List[CompetitorItem]
    Target_Market_Trends: List[MarketTrendItem]
    LO_Exit_Strategy: LOExitStrategy
    Export_and_Contract_Stats: ExportAndContractStats

class MarketResponseSchema(BaseModel):
    Growth_Potential: GrowthPotential

# =========================================================================
# 2. [분석 엔진(Agent) 실행부]
# =========================================================================
def analyze(pdf_path: str, company_name: str, industry_sector: str, extra_text: str = "") -> dict:
    print(f"   [Market Agent] 시장 동향 및 경쟁사 검색 중...")

    # 1. RAG: 최신 시장 동향 및 경쟁사 정밀 검색 (검색어 개편)
    search_query = f"""
    1. {company_name} 주요 경쟁사 현황 및 제품 비교
    2. {industry_sector} 시장 규모(TAM, SAM, SOM) 및 전망
    3. {industry_sector} 최신 트렌드 및 규제 이슈
    """
    rag_res = call_gemini(f"'{company_name}'이 속한 {industry_sector} 시장의 최신 동향과 경쟁사 정보를 검색하십시오.\n{search_query}", tools=[{"google_search": {}}])
    rag_context = rag_res.get("text", "") if rag_res.get("ok") else ""

    # 2. Analysis
    # 🚨 프롬프트 지시사항은 완벽히 유지하되, 하드코딩된 Output Schema 문자열만 제거
    prompt = f"""
    당신은 벤처캐피탈(VC)의 시니어 산업 분석 애널리스트입니다.
    메인 IR 자료(PDF), 검색된 시장 데이터(RAG Context), 그리고 **[보충 문서 데이터]**를 모두 결합하여 시장성을 심층 분석하십시오.

    [보충 문서 데이터]
    {extra_text}

    [RAG Context]
    {rag_context}

    [작성 지침 - 매우 중요]
    1. **시장 규모 (TAM/SAM/SOM)**: 
       - 무리하게 정확한 숫자를 만들어내지 마십시오. 숫자가 없다면 해당 시장의 '정의(Definition)'와 '타겟 고객층'을 서술하여 방어하십시오.
       - 줄글이 아닌, 핵심 내용만 간결하게 요약하여 작성하십시오.
    2. 🚨 [경쟁사 비교 분석 - 동기화 강제 지침]: 
       - 하단에 [참고: 주요 상장 경쟁사(Peer Group) 명단]으로 제공된 기업들은 5장의 밸류에이션 산정 기준이 되는 핵심 기업들입니다.
       - **보고서의 논리적 일관성을 위해, 제공된 [Peer Group] 명단 중 최소 2개 이상의 기업을 반드시 경쟁사 비교표(Competitors_Comparison)에 포함시키십시오.**
       - 만약 타겟 기업과 더 직접적으로 경쟁하는 글로벌 선도 기업이나 스타트업이 있다면, 앞서 언급한 Peer Group 기업들과 함께 섞어서 총 3~4개 사로 구성하십시오.
       - 'A사', 'B사' 등의 가명 사용은 절대 금지하며, 반드시 실존 기업명만 사용하십시오.
    3. **시장 트렌드**: 최근 뉴스나 규제 등 주요 트렌드 3가지를 도출하십시오.
    4. **Exit 전략**: 동종 업계의 M&A 사례나 IPO 트렌드를 참고하여 회수 가능성을 시나리오로 제시하십시오.
    5. **[성장 지표 추출]**: 문서 내에 연도별 수출 실적, 계약 건수, 매출 규모 추이 데이터가 표나 그래프로 존재한다면 `Export_and_Contract_Stats`에 연도와 수치 쌍(예: ["2024", "150"])으로 전사하십시오. 없으면 빈 배열 `[]`을 반환하십시오.
    6. **[엑시트 전략 분리 작성 엄수]**: 'Expected_LO_Scenarios' 작성 시 M&A와 IPO는 엑시트의 성격이 완전히 다르므로 반드시 2개의 독립된 시나리오로 분리하여 각 상황에 맞게 날카롭게 작성하십시오.
    """
    
    # 🚨 [핵심 변경] utils.py의 call_gemini에 response_schema 파라미터를 넘겨 JSON 출력을 완벽히 통제합니다.
    res = call_gemini(prompt, pdf_path=pdf_path, response_schema=MarketResponseSchema)
    
    if res.get("ok"):
        return safe_json_loads(res["text"])
    else:
        print(f"   [Market Agent] Error: {res.get('error')}")
        return {}