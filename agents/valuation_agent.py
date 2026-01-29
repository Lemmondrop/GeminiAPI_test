import os
import json
from utils import call_gemini, safe_json_loads, load_industry_codes, get_companies_by_code

# 밸류에이션 전용 스키마
VALUATION_SCHEMA = """
{
  "Valuation_and_Judgment": {
      "Valuation_Table": [
          { "Round": "Round", "Pre_Money": "Pre-Val", "Post_Money": "Post-Val", "Comment": "비고" }
      ],
    "Valuation_Logic_Detail": {
          "Step1_Selected_Industries": "AI가 선정한 3가지 핵심 산업내용 및 코드",
          "Step2_Combined_Peer_Pool": "3개 코드로 확장된 상장사 모집단 규모",
          "Step3_Final_Peer_Selection": "재무/사업 연관성으로 최종 필터링된 Peer (3~5개)",
          "Step4_Valuation_Calculation": "가치 산출 과정 및 결과",
          "Applied_Multiple": "적용 배수",
          "Target_Net_Income": "적용 실적",
          "Total_Shares_Outstanding": "주식수",
          "Estimated_Share_Price": "주당 가격",
          "Calculation_Rationale": "수식 포함 근거"
      },
      "Suitable_Investor_Type": "적합 투자자 유형"
  }
}
"""

def analyze(pdf_path: str, company_name: str, ceo_name: str, industry_sector: str) -> dict:
    print(f"   [Valuation Agent] '{company_name}' 데이터 기반 정밀 타겟팅 시작...")

    # ==============================================================================
    # [1] 파일 경로 설정
    # ==============================================================================
    # 1. 산업 정의 파일 (preprocess_corp_code_list) - 사용자 지정 경로
    industry_def_path = r"C:\Users\Woochul\Desktop\Project V\OCR Sample\data\metadata\preprocess_corp_code_list(prototype).csv"
    
    # 2. 회사 리스트 파일 (company_cord_prototype) - 같은 폴더 가정
    base_dir = os.path.dirname(industry_def_path)
    company_list_path = os.path.join(base_dir, "company_cord_prototype.csv")
    
    # 만약 위 경로에 없다면 루트나 기본 data 폴더 확인 (Fallback)
    if not os.path.exists(company_list_path):
        company_list_path = "company_cord_prototype.csv"

    # ==============================================================================
    # [2] Step 1: 다중 산업내용 선정 (LLM Selection)
    # ==============================================================================
    industry_map = load_industry_codes(industry_def_path) # { "산업내용": "코드" }
    
    target_codes = []
    selected_industries = []
    
    if not industry_map:
        print("   ⚠️ 산업 정의 파일을 로드할 수 없어 앵커 검색 모드로 전환합니다.")
        peers_str = f"{industry_sector} 관련 상장사"
    else:
        # LLM에게 선택지 제공 (산업내용 리스트)
        industry_names = list(industry_map.keys())
        context_str = "\n".join(industry_names[:1500]) 

        prompt_step1 = f"""
        당신은 산업분류 전문가입니다.
        IR 자료를 분석하여, 아래 **[산업내용 목록]** 중 대상 기업 '{company_name}'의 사업과 연관성이 높은 **'산업내용'을 3개** 선택하십시오.

        [분석 가이드]
        - 핵심 사업뿐만 아니라 전방/후방 산업도 고려하여 **3가지를 선정**하십시오.
        - 예: 바이오 기업 -> '의약품 제조업' + '자연과학 연구개발업' + '의료용 기기 제조업'

        [산업내용 목록 (일부)]
        {context_str}

        [Output JSON]
        {{ 
            "Selected_Industries": ["산업내용1", "산업내용2", "산업내용3"], 
            "Reason": "선정 이유" 
        }}
        """
        
        res1 = call_gemini(prompt_step1, pdf_path=pdf_path)
        selection = safe_json_loads(res1.get("text", ""))
        selected_industries = selection.get("Selected_Industries", [])
        
        print(f"   👉 AI 선정 산업내용 (3개): {selected_industries}")
        
        # 코드 매핑 (3개 각각 수행)
        for ind_name in selected_industries:
            code = industry_map.get(ind_name.strip(), "")
            if not code:
                # 부분 일치 검색
                for name, c in industry_map.items():
                    if ind_name.strip() in name:
                        code = c
                        break
            if code:
                target_codes.append(code)
        
        # 중복 제거
        target_codes = list(set(target_codes))
        print(f"   👉 매핑된 산업분류코드 리스트: {target_codes}")
    # ==============================================================================
    # [3] Step 2: 코드 -> 회사명 추출 (Deterministic Matching)
    # ==============================================================================
    final_peer_names = []
    
    if target_codes:
        for code in target_codes:
            # 각 코드별로 기업 추출
            peers = get_companies_by_code(code, company_list_path)
            
            # 만약 정확히 일치하는 기업이 없으면 상위 분류 검색 (Fallback)
            if not peers and len(code) > 3:
                broad_code = code[:-1]
                peers = get_companies_by_code(broad_code, company_list_path)
                
            final_peer_names.extend(peers)

    # 통합 후 중복 제거
    final_peer_names = list(set(final_peer_names))
    
    # 상위 10개 선정 (랜덤성을 줄이기 위해 정렬 후 선택하거나, 그냥 앞에서부터)
    # 데이터가 많으면 더 다양한 기업이 포함될 확률이 높아짐
    peers_to_search = final_peer_names[:10]
    
    print(f"   👉 [통합] 최종 검색 대상 Peer ({len(peers_to_search)}개 / 전체 Pool {len(final_peer_names)}개): {peers_to_search}")
    
    peers_str = ", ".join(peers_to_search)
    if not peers_str: peers_str = f"{industry_sector} 대표 상장사"

    # ==============================================================================
    # [4] Step 3: RAG & Valuation
    # ==============================================================================
    search_prompt = f"""
    비상장 기업 '{company_name}'의 가치평가를 수행합니다.
    3가지 산업분류코드를 통해 식별된 유사 상장사들: **{peers_str}**
    
    [지시사항]
    1. 위 기업들 중 **대표적인 3~5개사**의 '현재 시가총액', '매출액', '영업이익', 'PER'를 검색하십시오.
       - 검색어: `site:finance.naver.com {peers_str} 시가총액 PER`
    2. 타겟 기업('{company_name}')의 '발행주식수'와 '자본금' 정보를 찾으십시오.
    """
    
    rag_res = call_gemini(search_prompt, tools=[{"google_search": {}}])
    rag_context = rag_res.get("text", "")

    main_prompt = f"""
    당신은 밸류에이션 전문가입니다.
    
    [Data]
    - Selected Industries: {selected_industries}
    - Industry Codes: {target_codes}
    - Identified Peers: {peers_str}
    - Market Data: 
    {rag_context}

    [Logic]
    1. **Filtering**: 식별된 기업 중 적자거나 시총이 너무 작은 곳은 제외하고 최종 Peer를 선정하십시오.
    2. **Valuation**: [Target 추정 실적] x [Peer 평균 Multiple] = **Total Value(A)**
    3. **Share Price**: (A) / [Target 주식수(B)] = **주당 가격**
       - (B가 없으면 자본금/500원 역산 또는 50만주 가정)
    4. **Output**: 반드시 수식(Equation)을 포함하여 결과를 작성하십시오.

    [Output Schema]
    {VALUATION_SCHEMA}
    """
    
    res = call_gemini(main_prompt, pdf_path=pdf_path)
    if res.get("ok"):
        return safe_json_loads(res["text"])
    else:
        print(f"   [Valuation Agent] Error: {res.get('error')}")
        return {}