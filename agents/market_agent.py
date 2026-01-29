### 시장 규모, 경쟁사, 트렌드 담당 (RAG Heavy)
import json
from utils import call_gemini, safe_json_loads

MARKET_SCHEMA = """
{
  "Growth_Potential": {
    "Target_Market_Analysis": {
        "Target_Area": "타겟 시장 정의 (TAM/SAM/SOM)",
        "Market_Characteristics": "시장 규모, CAGR, 성장 요인, 규제 환경",
        "Competitive_Positioning": "경쟁사 현황 및 당사의 포지셔닝 (구체적 비교)"
    },
    "Target_Market_Trends": [
      { "Type": "뉴스/보고서", "Source": "출처", "Content": "트렌드 상세 내용 (3문장 이상)" }
    ],
    "LO_Exit_Strategy": {
        "Verified_Signals": ["기존 레퍼런스, 파트너십"],
        "Expected_LO_Scenarios": [
            { "Category": "M&A/IPO", "Probability": "가능성", "Comment": "시나리오 상세" }
        ],
        "Valuation_Range": "시장 관점에서의 가치 평가"
    },
    "Export_and_Contract_Stats": {
      "Export_Graph_Data": [["Year", "Value"]],
      "Contract_Count_Graph_Data": [["Year", "Count"]],
      "Sales_Graph_Data": [["Year", "Revenue"]]
    }
  }
}
"""

def analyze(pdf_path: str, company_name: str, industry_sector: str) -> dict:
    print(f"   [Market Agent] 시장 동향 및 경쟁사 검색 중...")

    # 1. RAG: 최신 시장 동향 검색
    search_query = f"""
    1. {company_name} 경쟁사 시장점유율
    2. {industry_sector} 시장 규모 전망 CAGR
    3. {industry_sector} 최신 트렌드 규제 이슈
    """
    rag_res = call_gemini(f"'{company_name}'이 속한 {industry_sector} 시장의 최신 동향과 경쟁사 정보를 검색하십시오.\n{search_query}", tools=[{"google_search": {}}])
    rag_context = rag_res.get("text", "") if rag_res.get("ok") else ""

    # 2. Analysis
    prompt = f"""
    당신은 산업 분석 애널리스트입니다.
    IR 자료와 검색된 시장 데이터(RAG Context)를 결합하여 시장성을 심층 분석하십시오.

    [RAG Context]
    {rag_context}

    [분석 지침]
    1. **시장 규모**: TAM/SAM/SOM 관점에서 시장의 크기와 성장성(CAGR)을 수치로 제시하십시오.
    2. **경쟁 현황**: 주요 경쟁사들의 이름을 명시하고, 당사의 기술적/사업적 차별점을 부각하십시오.
    3. **Exit 전략**: 동종 업계의 M&A 사례나 IPO 트렌드를 참고하여 회수 가능성을 시나리오로 제시하십시오.

    [Output Schema]
    {MARKET_SCHEMA}
    """
    
    res = call_gemini(prompt, pdf_path=pdf_path)
    if res.get("ok"):
        return safe_json_loads(res["text"])
    else:
        print(f"   [Market Agent] Error: {res.get('error')}")
        return {}