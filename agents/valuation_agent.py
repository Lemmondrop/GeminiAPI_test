import os
import json
import math
import datetime
import re
from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
from utils import call_gemini, safe_json_loads, load_industry_codes, get_companies_by_code
from utils_extended import full_peer_filtering_pipeline

# =========================================================================
# 1. [Pydantic 스키마 정의] - 각 LLM 호출 단계별로 완벽한 타입 강제
# =========================================================================

# [Step 1 & Step 1-확장] 산업 분류 추출용 스키마
class IndustrySelection(BaseModel):
    Selected_Industries: List[str] = Field(description="선택된 산업내용 목록 배열")

# [Step 8] Target 수치 핀셋 추출용 스키마
class ExtractionData(BaseModel):
    target_net_income: Optional[int] = Field(description="추정 당기순이익 (원 단위 정수). 모르면 null")
    total_shares: Optional[int] = Field(description="총 발행주식수. 3가지 방법으로도 유추할 수 없으면 null")
    target_round: Optional[str] = Field(description="현재 진행 중이거나 목표로 하는 투자 유치 라운드 영문. 없으면 null")

# [Step 8-Fallback] RAG 자본금 변환용 스키마
class CapitalConversion(BaseModel):
    capital_amount_int: Optional[int] = Field(description="원(₩) 단위의 정수로 변환된 자본금. 도저히 알 수 없으면 null")
    capital_amount_str: str = Field(description="원본 자본금 문자열 (예: '15억 1천만원')")
    source: str = Field(description="출처 (예: 잡코리아, 사람인 등)")

# [Step 9] 최종 Valuation 정성평가 출력용 스키마
class ValuationTableItem(BaseModel):
    Round: str
    Pre_Money: str
    Post_Money: str
    Comment: str

class ThreeAxisAssessment(BaseModel):
    Technology_Rating: str = Field(description="기술성 평가 요약 (핵심 경쟁력 및 진입장벽, 1~2문장)")
    Growth_Rating: str = Field(description="성장성 평가 요약 (시장 확장성 및 매출 성장 잠재력, 1~2문장)")
    Exit_Rating: str = Field(description="회수성 평가 요약 (IPO/M&A 가능성 등 엑시트 시나리오, 1~2문장)")

class ValuationAndJudgment(BaseModel):
    Valuation_Table: List[ValuationTableItem]
    Valuation_Logic_Detail: Dict[str, Any] = Field(default_factory=dict, description="시스템에서 파이썬 변수를 강제 주입할 빈 딕셔너리")
    Three_Axis_Assessment: ThreeAxisAssessment
    Suitable_Investor_Type: str = Field(description="적합 투자자 유형")

class InvestmentRiskItem(BaseModel):
    Risk_Title: str = Field(description="리스크 핵심 요약 (예: 전방시장 악화에 따른 매출 감소 우려)")
    Risk_Detail: str = Field(description="해당 리스크가 기업에 미칠 부정적 영향 및 원인 (특수기호 없이 평문)")
    Mitigation_and_Outlook: str = Field(description="기업의 대응책 또는 애널리스트의 긍정적 판단/해소 가능성 (특수기호 없이 평문)")

class ValuationResponseSchema(BaseModel):
    Valuation_and_Judgment: ValuationAndJudgment
    Investment_Risks: List[InvestmentRiskItem]
    Investment_Rating: str = Field(description="종합 등급 (예: 긍정적, 관심, 관망, 부정적 등)")
    Final_Conclusion: str = Field(description="해당 기업에 대한 종합 투자 판단 및 밸류에이션 결론 (3~4문장 내외)")


# =========================================================================
# 2. [분석 엔진(Agent) 메인 실행부]
# =========================================================================
def analyze(pdf_path: str, company_name: str, ceo_name: str, industry_sector: str, extra_text: str = "") -> dict:
    print(f"   [Valuation Agent] '{company_name}' 정밀 타겟팅 (Full 4-Stage Pipeline) 시작...")

    industry_def_path = r"C:\Users\Researcher\Desktop\Project V\OCR Sample\data\metadata\preprocess_corp_code_list(prototype).csv"
    base_dir = os.path.dirname(industry_def_path)
    company_list_path = os.path.join(base_dir, "company_cord_prototype.csv")
    if not os.path.exists(company_list_path): 
        company_list_path = "company_cord_prototype.csv"

    # [Step 1] 산업분류코드 추출
    industry_map = load_industry_codes(industry_def_path)
    target_codes = []
    selected_industries = []
    
    if industry_map:
        industry_names = list(industry_map.keys())
        context_str = "\n".join(industry_names[:1500]) 

        prompt_step1 = f"""
당신은 산업분류 전문가입니다.
IR 자료(PDF)와 **[보충 문서 데이터]**를 분석하여, 아래 [산업내용 목록] 중 대상 기업 '{company_name}'의 사업과 가장 연관성이 높은 '산업내용' 3가지를 선택하십시오.

[보충 문서 데이터]
{extra_text}

[산업내용 목록 (일부)]
{context_str}
"""
        # 🚨 [최적화 1] 산업 선택 Pydantic 주입
        res = call_gemini(prompt_step1, pdf_path=pdf_path, response_schema=IndustrySelection)
        selection = safe_json_loads(res.get("text", ""))
        selected_industries = selection.get("Selected_Industries", [])
        print(f"   👉 [Step 1] AI 선정 산업: {selected_industries}")
        
        for ind in selected_industries:
            search_keyword = ind.strip()
            for k, v_list in industry_map.items():
                if search_keyword in k:
                    target_codes.extend(v_list)
        
        target_codes = list(set(target_codes))

    # [Step 2] 기업 리스트 추출
    raw_peers = []
    if target_codes:
        for code in target_codes:
            peers = get_companies_by_code(code, company_list_path)
            if not peers and len(code) > 3:
                peers = get_companies_by_code(code[:-1], company_list_path)
            raw_peers.extend(peers)
    
    raw_peers = list(set(raw_peers))
    
    # =========================================================
    # 동적 확장 로직 (Dynamic Expansion) - 최소 5개 ~ 최대 7개
    # =========================================================
    if len(raw_peers) <= 10:
        print(f"   ⚠️ [Fallback] 1차 매칭 기업이 {len(raw_peers)}개로 부족합니다. 밸류체 전반으로 산업분류를 최소 5개 이상 확장 탐색합니다.")
        
        prompt_step1_expand = f"""
당신은 산업분류 전문가입니다.
앞서 선택한 3개의 산업만으로는 비교할 상장사 그룹이 턱없이 부족합니다.
대상 기업 '{company_name}'의 밸류체인(원재료, 후방 제조 공정, 전방 적용 시장 등)을 모두 포괄하여, 
아래 [산업내용 목록] 중 관련성이 높은 '산업내용'을 **최소 5개에서 최대 7개까지** 넓게 다시 선택하십시오.

[보충 문서 데이터]
{extra_text}
[산업내용 목록 (일부)]
{context_str}
"""
        # 🚨 [최적화 2] 산업 확장 선택 Pydantic 주입
        res_expand = call_gemini(prompt_step1_expand, pdf_path=pdf_path, response_schema=IndustrySelection)
        selection_expand = safe_json_loads(res_expand.get("text", ""))
        expanded_industries = selection_expand.get("Selected_Industries", [])
        
        print(f"   👉 [Step 1-확장] AI 추가 선정 산업: {expanded_industries}")
        
        expanded_codes = []
        for ind in expanded_industries:
            search_keyword = ind.strip()
            for k, v_list in industry_map.items():
                if search_keyword in k:
                    expanded_codes.extend(v_list)
                    
        expanded_codes = list(set(expanded_codes))
        for code in expanded_codes:
            peers = get_companies_by_code(code, company_list_path)
            if not peers and len(code) > 3:
                peers = get_companies_by_code(code[:-1], company_list_path)
            raw_peers.extend(peers)
            
        raw_peers = list(set(raw_peers))
        print(f"   ✅ [Step 1-확장 완료] 최종 확보된 1차 모집단: {len(raw_peers)}개 사")

    # [Step 3~6] Full 4-Stage Filtering Pipeline
    pipeline_result = full_peer_filtering_pipeline(
        target_pdf_path=pdf_path,
        company_name=company_name,
        raw_peer_names=raw_peers,
        company_csv_path=company_list_path,
        similarity_threshold=0.3
    )
    
    stage2_dec = pipeline_result.get("stage2_dec_passed", [])
    stage2_profit = pipeline_result.get("stage2_profit_passed", [])
    stage3_business = pipeline_result.get("stage3_business_passed", [])
    stage4_final = pipeline_result.get("stage4_final_peers", [])
    
    final_peers = stage4_final[:10]
    peers_str = ", ".join(final_peers)

    # [Step 7] RAG 검색 (자유 텍스트 기반)
    if final_peers:
        search_prompt = f"Target Peers: {peers_str}\n각 기업의 시가총액, PER, 2024년 당기순이익, Target 기업 '{company_name}'의 발행주식수를 검색하십시오."
        rag_res = call_gemini(search_prompt, tools=[{"google_search": {}}])
        rag_context = rag_res.get('text', '') if rag_res.get('ok') else ''
    else:
        rag_context = "Peer 데이터 부족"

    # [Step 8] 파이썬 기반 Valuation 연산 엔진 (환각 0%)
    target_year = datetime.datetime.now().year + 1
    
    valid_pers = []
    scraped_per_info = []
    if "details" in pipeline_result and "stage4_requirements" in pipeline_result["details"]:
        for name, passed, reason, info in pipeline_result["details"]["stage4_requirements"]:
            if name in final_peers:
                per_val = info.get("per", "N/A")
                scraped_per_info.append(f"- {name}: {per_val}")
                if passed and isinstance(per_val, (int, float)):
                    valid_pers.append(per_val)
                    
    peer_avg_per = sum(valid_pers) / len(valid_pers) if valid_pers else 0.0

    extract_prompt = f"""
당신은 집요하고 정확한 재무 데이터 추출 전문가입니다.
대상 기업 '{company_name}'의 문서에서 아래 3가지 정보만 찾아내십시오. 계산 과정은 출력하지 말고 오직 최종 결과만 추출하십시오.

1. {target_year}년(또는 가장 인접한 미래 연도) 추정 당기순이익 (원 단위 정수 표기)
   - 영업이익만 있고 순이익이 없다면, 영업이익을 대안으로 추출하십시오.
2. 총 발행주식수 (아래 1~3순위의 방법을 순차적으로 동원하여 어떻게든 찾아내십시오.)
   - [1순위]: 문서에 명시된 '총 발행주식수', '보통주 주식수' 또는 '상장 예정 주식수'
   - [2순위]: '자본금'과 '1주당 액면가' 정보가 있다면 (자본금 ÷ 액면가)로 계산. 🚨만약 '자본금'만 있고 액면가가 없다면, 액면가를 500원으로 강제 가정하고 (자본금 ÷ 500)을 계산.
   - [3순위]: 주요 주주의 '보유 주식수'와 '지분율(%)'이 있다면 비례식으로 역산한 전체 주식수
   - 위 3가지 방법으로도 절대 유추할 수 없다면 null을 반환하십시오.
3. 현재 진행 중이거나 목표로 하는 '투자 유치 라운드'
   - 문서 내 투자 제안(Offering) 부분이나 요약본을 참고하여, 이 회사가 현재 어떤 라운드의 투자를 받고자 하는지 영문으로 기재하십시오.
   - (예: "Seed", "Pre-A", "Series A", "Series B", "Pre-IPO" 등). 명확히 알 수 없다면 null을 반환하십시오.

[보충 문서 데이터]
{extra_text}
"""
    # 🚨 [최적화 3] 데이터 핀셋 추출 Pydantic 주입
    res_extract = call_gemini(extract_prompt, pdf_path=pdf_path, response_schema=ExtractionData)
    extracted_data = safe_json_loads(res_extract.get("text", ""))
    t_net_income = extracted_data.get("target_net_income")
    t_shares = extracted_data.get("total_shares")
    t_round = extracted_data.get("target_round")

    fallback_msg = "" 

    if not t_shares:
        print(f"   🔍 [Smart Fallback] 문서 내 주식수 부재 감지. 구글 검색(RAG)을 통해 집중 탐색을 시작합니다...")
        try:
            clean_name = re.sub(r'\(.*?\)', '', company_name).replace('주식회사', '').replace('(주)', '').strip()
            
            rag_fallback_prompt = f"한국 비상장 기업 '{clean_name}'의 '자본금' 또는 '발행주식총수'를 검색해서 정확한 수치를 알려주세요. 잡코리아, 사람인, 혁신의숲 등 구인구직 사이트나 기사 데이터를 적극 참고하세요."
            rag_res = call_gemini(rag_fallback_prompt, tools=[{"google_search": {}}])
            fallback_context = rag_res.get('text', '') if rag_res.get('ok') else ''
            
            if fallback_context:
                conv_prompt = f"""
                당신은 재무 데이터 추출 전문가입니다.
                다음 검색 결과를 바탕으로 '{clean_name}'의 '자본금'을 찾아내십시오.
                그리고 찾은 자본금을 오직 '원(₩) 단위의 정수(Integer)'로만 변환하여 반환하십시오.
                (예: '15억 1천만원' -> 1510000000)
                
                절대 주식수를 직접 계산하거나 500으로 나누지 마십시오.
                신뢰성을 검증하기 위해, 원본 '자본금 문자열'과 '출처'도 함께 추출하십시오.
                
                [검색 결과 Context]
                {fallback_context}
                """
                # 🚨 [최적화 4] 자본금 정수 변환 Pydantic 주입
                res_conv = call_gemini(conv_prompt, response_schema=CapitalConversion)
                conv_data = safe_json_loads(res_conv.get("text", ""))
                
                capital_int = conv_data.get("capital_amount_int")
                capital_str = conv_data.get("capital_amount_str", "미상")
                source_name = conv_data.get("source", "웹 검색")
                
                if capital_int and isinstance(capital_int, (int, float)) and capital_int > 0:
                    t_shares = int(capital_int / 500)
                    fallback_msg = f" (출처: {source_name}, 적용 자본금: {capital_str})"
                    print(f"      ✅ 파이썬 엔진 기반 주식수 역산 완료: {t_shares:,}주{fallback_msg}")
                else:
                    print("      ⚠️ RAG 웹 검색으로 자본금 정보를 도출하지 못했습니다.")
            else:
                print("      ⚠️ RAG 검색 결과가 반환되지 않았습니다.")
                
        except Exception as e:
            print(f"      ⚠️ RAG 조건부 탐색 중 오류 발생: {e}")
            
    # 투자 라운드 기반 동적 할인율 매핑 로직
    d_rate_pv = 0.50  # 기본 할인율
    
    round_str = str(t_round).lower().replace(" ", "").replace("-", "") if t_round else ""
    
    # 초기 라운드
    if "seed" in round_str or "시드" in round_str or "angel" in round_str: 
        d_rate_pv = 0.80
    elif "prea" in round_str or "프리a" in round_str: 
        d_rate_pv = 0.60
    elif "seriesa" in round_str or "시리즈a" in round_str: 
        d_rate_pv = 0.50
        
    # 중기 라운드
    elif "seriesb" in round_str or "시리즈b" in round_str or "preb" in round_str: 
        d_rate_pv = 0.40
    elif "seriesc" in round_str or "시리즈c" in round_str or "prec" in round_str: 
        d_rate_pv = 0.30
        
    # 후기 라운드
    elif "ipo" in round_str or "상장" in round_str: 
        d_rate_pv = 0.20

    display_round = t_round if t_round else "미상 (기본값 적용)"

    # 3. 파이썬 3-Scenario 사칙연산 및 텍스트 조립
    calc_rationale = ""
    target_est_price = "-"
    scenario_data = [] 
    
    if peer_avg_per > 0 and t_net_income:
        d_rate_ipo = 0.40 # 공모 할인율 고정
        
        calc_rationale += f"1단계 (Peer 평균 PER): {peer_avg_per:.2f}배 적용\n"
        calc_rationale += f"2단계 (할인율 적용): 타겟 투자 라운드 [{display_round}] 리스크를 반영하여 현가 할인율 {d_rate_pv*100:.0f}% 적용\n"
        calc_rationale += f"3~4단계 (가치 산출): 하단의 시나리오별 밸류에이션 표 참조\n"
        
        if t_shares:
            if fallback_msg: calc_rationale += f"  * 참고: 기준 주식수 {t_shares:,}주 (문서 내 정보 부재로 외부 검색{fallback_msg} 바탕으로 액면가 500원 가정 역산)"
            else: calc_rationale += f"  * 참고: 기준 주식수 {t_shares:,}주 (명시된 주식수가 없을 경우, 자본금을 바탕으로 액면가 500원을 가정하여 역산된 주식수일 수 있습니다.)"
        else:
            calc_rationale += f"  * 참고: 자본금 및 발행주식수 정보 부재로 주당 단가 산출 생략"

        scenarios = [("낙관적", 1.0), ("중립적", 0.70), ("보수적", 0.50)]
        
        def get_approx_korean(val):
            if val >= 1000000000000: 
                jo = int(val // 1000000000000)
                eok = int((val % 1000000000000) // 100000000)
                return f"(약 {jo}조 {eok:,}억)" if eok > 0 else f"(약 {jo}조)"
            elif val >= 100000000: return f"(약 {int(val // 100000000):,}억)"
            else: return f"(약 {int(val // 10000):,}만)"

        for scen_name, ratio in scenarios:
            adj_income = t_net_income * ratio
            pv = adj_income / (1 + d_rate_pv)
            pv_trunc = math.floor(pv / 1000000) * 1000000
            ev = pv_trunc * peer_avg_per
            
            pv_approx = get_approx_korean(pv_trunc)
            ev_approx = get_approx_korean(ev)
            
            if t_shares:
                # [기초] 보통주 평가가액
                price_ps = ev / t_shares 
                
                # [계산] 최종 희망공모가 (보통주 기준 40% 할인)
                final_ipo_price = price_ps * (1 - d_rate_ipo)
                final_ipo_trunc = math.floor(final_ipo_price / 100) * 100
                                
                price_str = f"{price_ps:,.0f}원"
                final_ipo_str = f"{final_ipo_trunc:,.0f}원"
            else:
                price_str = "-"
                final_ipo_str = "-"
                
            scenario_data.append({
                "Scenario": scen_name,
                "Ratio": f"IR 추정치 {int(ratio*100)}% 달성",
                "PV": f"{pv_trunc:,.0f}원\n{pv_approx}",
                "EV": f"{ev:,.0f}원\n{ev_approx}",
                "Price": price_str,        # 주당 평가가액(보통주)
                "Final_Price": final_ipo_str, # 최종 희망공모가
            })
            
        if t_shares: target_est_price = scenario_data[0]["Final_Price"] 
        else: target_est_price = "산출 불가 (주식수 부재)"
            
    else:
        calc_rationale = "※ 산출 불가 사유:\n"
        if peer_avg_per <= 0: calc_rationale += "- 유효한 Peer PER 데이터 부재\n"
        if not t_net_income: calc_rationale += f"- {target_year}년 추정 당기순이익 정보 부재\n"

    # [Step 9] 정성적 평가 (LLM은 평가에만 집중)
    main_prompt = f"""
당신은 기업가치 평가 전문가입니다.
가치 평가(Valuation) 산출 및 계산은 파이썬 시스템이 이미 완벽하게 수행했습니다. 
당신은 아래 정보를 바탕으로 **정성적 분석 및 투자 판단**에만 집중하십시오.

[Input Data]
- 대상 기업: {company_name}
- 최종 Peer Group: {peers_str}
- 시장 데이터 (RAG): {rag_context}
- [보충 문서 데이터]: {extra_text}

[작성 지침]
- 'Three_Axis_Assessment' 필드에 기술성, 성장성, 회수성(Exit)을 각각 1~2문장으로 날카롭게 평가하십시오.
- 'Investment_Risks' 필드에 대상 기업의 시장 상황, 기술적 한계, 재무 상태 등을 종합하여 가장 치명적인 3가지 투자 리스크를 도출하고 완화 방안을 짝지어 작성하십시오.
- 'Investment_Rating'에 종합 등급을 적고, 'Final_Conclusion'에 3~4문장으로 투자 결론을 내리십시오.
- 계산 관련 수식은 시스템이 강제 주입할 것이므로 임의로 생성하지 마십시오.
- [전문가 문체(Tone & Manner) 엄수]: 'Three_Axis_Assessment'와 'Final_Conclusion'을 포함한 모든 정성적 평가를 작성할 때, "~할 것이다", "~다" 등의 단정적이거나 AI스러운 어미 사용을 절대 금지합니다. 실제 VC 심사역 및 증권사 애널리스트들이 사용하는 전문적인 어미("~로 사료된다", "~할 필요가 있다", "~할 것으로 전망된다", "~로 판단된다" 등)로 문장이 자연스럽게 귀결되도록 엄격히 작성하십시오.
"""
    # 🚨 [최적화 5] 최종 Valuation 정성평가 Pydantic 주입
    res = call_gemini(main_prompt, pdf_path=pdf_path, response_schema=ValuationResponseSchema, max_tokens=8192)
    val_data = safe_json_loads(res.get("text", "") if res.get("ok") else "{}")
    
    # [Data Injection] 
    if "Valuation_and_Judgment" not in val_data: val_data["Valuation_and_Judgment"] = {}
    v_judge = val_data["Valuation_and_Judgment"]
    if "Valuation_Logic_Detail" not in v_judge: v_judge["Valuation_Logic_Detail"] = {}
    
    detail = v_judge["Valuation_Logic_Detail"]
    detail["Step1_Industries"] = selected_industries
    detail["Step2_Raw_Pool"] = raw_peers
    detail["Step3_Dec_Filtered"] = stage2_dec
    detail["Step3_Profit_Filtered"] = stage2_profit
    detail["Step4_Business_Filtered"] = stage3_business
    detail["Step5_Final_Peers"] = final_peers
    
    # 파이썬 연산 결과를 LLM JSON 템플릿에 안전하게 강제 주입
    detail["Calculation_Rationale"] = calc_rationale
    detail["Scenario_Valuation"] = scenario_data
    detail["Estimated_Share_Price"] = target_est_price
    detail["Target_Net_Income"] = f"{t_net_income:,}원" if t_net_income else "정보 부재"
    detail["Total_Shares_Outstanding"] = f"{t_shares:,}주" if t_shares else "정보 부재"
    detail["Applied_Multiple"] = f"{peer_avg_per:.2f}배" if peer_avg_per > 0 else "산출 불가"

    detail["Discount_Rate_PV"] = f"{int(d_rate_pv*100)}%"
    detail["Discount_Rate_IPO"] = f"{int(d_rate_ipo*100)}%" if 'd_rate_ipo' in locals() else "40%"
    
    if "details" in pipeline_result:
        raw_s3 = pipeline_result["details"].get("stage3_similarity", [])[:5]
        detail["Stage3_Similarity_Scores"] = [
            (
                {"company": x.get("company") or x.get("name") or "", "score": x.get("score", 0.0), "reason": x.get("reason", ""), "main_products": x.get("main_products", "")}
                if isinstance(x, dict) else {"company": x[0] if len(x) > 0 else "", "score": x[1] if len(x) > 1 else 0.0, "reason": x[2] if len(x) > 2 else ""}
            ) for x in raw_s3
        ]
        detail["Stage4_Requirements_Check"] = [
            {"company": name, "passed": passed, "reason": reason, "info": info}
            for name, passed, reason, info in pipeline_result["details"].get("stage4_requirements", [])[:30] 
        ]
    
    if not v_judge.get("Valuation_Table"):
        v_judge["Valuation_Table"] = [{ "Round": "IPO (예상)", "Pre_Money": "-", "Post_Money": "-", "Comment": f"예상 공모가: {target_est_price}" }]

    return val_data

# import os
# import json
# import re
# import math
# from utils import call_gemini, safe_json_loads, load_industry_codes, get_companies_by_code
# from utils_extended import full_peer_filtering_pipeline

# # 밸류에이션 전용 스키마
# VALUATION_SCHEMA = """
# {
#   "Valuation_and_Judgment": {
#       "Valuation_Table": [
#           { "Round": "Round", "Pre_Money": "Pre-Val", "Post_Money": "Post-Val", "Comment": "비고" }
#       ],
#       "Valuation_Logic_Detail": {
#           "Step1_Industries": "AI 선정 산업내용 (3개)",
#           "Step2_Raw_Pool": "코드 매칭된 전체 기업 리스트",
#           "Step3_Dec_Filtered": "12월 결산 통과 기업 리스트",
#           "Step3_Profit_Filtered": "흑자 통과 기업 리스트",
#           "Step4_Business_Filtered": "사업 유사성 통과 기업 리스트",
#           "Step5_Final_Peers": "최종 Peer Group (3~10개)",
#           "Applied_Multiple": "적용 배수 (PER/PSR)",
#           "Target_Net_Income": "적용 실적",
#           "Total_Shares_Outstanding": "주식수",
#           "Estimated_Share_Price": "주당 가격",
#           "Calculation_Rationale": "산출 수식"
#       },
#       "Three_Axis_Assessment": {
#         "Technology_Rating": "기술성 평가 요약 (핵심 경쟁력 및 진입장벽, 1~2문장)",
#         "Growth_Rating": "성장성 평가 요약 (시장 확장성 및 매출 성장 잠재력, 1~2문장)",
#         "Exit_Rating": "회수성 평가 요약 (IPO/M&A 가능성 등 엑시트 시나리오, 1~2문장)"
#       },
#       "Suitable_Investor_Type": "적합 투자자 유형"
#   },
#     "Investment_Risks": [
#       {
#           "Risk_Title": "리스크 핵심 요약 (예: 전방시장 악화에 따른 매출 감소 우려)",
#           "Risk_Detail": "해당 리스크가 기업에 미칠 부정적 영향 및 원인 (- 로 시작하는 문장 형태)",
#           "Mitigation_and_Outlook": "기업의 대응책 또는 애널리스트의 긍정적 판단/해소 가능성 (--> 로 시작하는 문장 형태)"
#       }
#   ],
#   "Investment_Rating": "종합 등급 (예: 긍정적, 관심, 관망, 부정적 등)",
#   "Final_Conclusion": "해당 기업에 대한 종합 투자 판단 및 밸류에이션 결론 (3~4문장 내외로 명확하게 서술)"
# }
# """

# def analyze(pdf_path: str, company_name: str, ceo_name: str, industry_sector: str, extra_text: str = "") -> dict:
#     print(f"   [Valuation Agent] '{company_name}' 정밀 타겟팅 (Full 4-Stage Pipeline) 시작...")

#     industry_def_path = r"C:\Users\Researcher\Desktop\Project V\OCR Sample\data\metadata\preprocess_corp_code_list(prototype).csv"
#     base_dir = os.path.dirname(industry_def_path)
#     company_list_path = os.path.join(base_dir, "company_cord_prototype.csv")
#     if not os.path.exists(company_list_path): 
#         company_list_path = "company_cord_prototype.csv"

#     # [Step 1] 산업분류코드 추출
#     industry_map = load_industry_codes(industry_def_path)
#     target_codes = []
#     selected_industries = []
    
#     if industry_map:
#         industry_names = list(industry_map.keys())
#         context_str = "\n".join(industry_names[:1500]) 

#         prompt_step1 = f"""
# 당신은 산업분류 전문가입니다.
# IR 자료(PDF)와 **[보충 문서 데이터]**를 분석하여, 아래 [산업내용 목록] 중 대상 기업 '{company_name}'의 사업과 가장 연관성이 높은 '산업내용' 3가지를 선택하십시오.

# [보충 문서 데이터]
# {extra_text}

# [산업내용 목록 (일부)]
# {context_str}

# [Output JSON] 
# {{ "Selected_Industries": ["산업1", "산업2", "산업3"] }}
# """
#         res = call_gemini(prompt_step1, pdf_path=pdf_path)
#         # [Step 1] 산업분류코드 매칭 로직 (수정됨)
#         selection = safe_json_loads(res.get("text", ""))
#         selected_industries = selection.get("Selected_Industries", [])
#         print(f"   👉 [Step 1] AI 선정 산업: {selected_industries}")
        
#         for ind in selected_industries:
#             search_keyword = ind.strip()
#             # 딕셔너리를 전체 순회하며 키워드가 포함된 모든 산업을 찾음
#             for k, v_list in industry_map.items():
#                 if search_keyword in k:
#                     # 🚨 v_list는 이제 리스트(예: ['C26129', 'C02612'])이므로 extend를 사용하여 모두 흡수
#                     target_codes.extend(v_list)
        
#         # 중복 코드 깔끔하게 제거
#         target_codes = list(set(target_codes))

#     # [Step 2] 기업 리스트 추출
#     raw_peers = []
#     if target_codes:
#         for code in target_codes:
#             peers = get_companies_by_code(code, company_list_path)
#             if not peers and len(code) > 3:
#                 peers = get_companies_by_code(code[:-1], company_list_path)
#             raw_peers.extend(peers)
    
#     raw_peers = list(set(raw_peers))
#     # =========================================================
#     # 🚨 [추가] 동적 확장 로직 (Dynamic Expansion) - 최소 5개 ~ 최대 7개
#     # =========================================================
#     if len(raw_peers) <= 10:
#         print(f"   ⚠️ [Fallback] 1차 매칭 기업이 {len(raw_peers)}개로 부족합니다. 밸류체인 전반으로 산업분류를 최소 5개 이상 확장 탐색합니다.")
        
#         prompt_step1_expand = f"""
# 당신은 산업분류 전문가입니다.
# 앞서 선택한 3개의 산업만으로는 비교할 상장사 그룹이 턱없이 부족합니다.
# 대상 기업 '{company_name}'의 밸류체인(원재료, 후방 제조 공정, 전방 적용 시장 등)을 모두 포괄하여, 
# 아래 [산업내용 목록] 중 관련성이 높은 '산업내용'을 **최소 5개에서 최대 7개까지** 넓게 다시 선택하십시오.

# [보충 문서 데이터]
# {extra_text}

# [산업내용 목록 (일부)]
# {context_str}

# [Output JSON] 
# {{ "Selected_Industries": ["본업 산업1", "연관 산업2", "연관 산업3", "전방 산업4", "후방 산업5", "추가 산업6"] }}
# """
#         res_expand = call_gemini(prompt_step1_expand, pdf_path=pdf_path)
#         selection_expand = safe_json_loads(res_expand.get("text", ""))
#         expanded_industries = selection_expand.get("Selected_Industries", [])
        
#         print(f"   👉 [Step 1-확장] AI 추가 선정 산업: {expanded_industries}")
        
#         expanded_codes = []
#         for ind in expanded_industries:
#             search_keyword = ind.strip()
#             for k, v_list in industry_map.items():
#                 if search_keyword in k:
#                     expanded_codes.extend(v_list)
                    
#         expanded_codes = list(set(expanded_codes))
        
#         # 확장된 코드로 상장사 다시 긁어오기
#         for code in expanded_codes:
#             peers = get_companies_by_code(code, company_list_path)
#             if not peers and len(code) > 3:
#                 peers = get_companies_by_code(code[:-1], company_list_path)
#             raw_peers.extend(peers)
            
#         raw_peers = list(set(raw_peers))
#         print(f"   ✅ [Step 1-확장 완료] 최종 확보된 1차 모집단: {len(raw_peers)}개 사")
#     # =========================================================

#     # [Step 3~6] Full 4-Stage Filtering Pipeline
#     pipeline_result = full_peer_filtering_pipeline(
#         target_pdf_path=pdf_path,
#         company_name=company_name,
#         raw_peer_names=raw_peers,
#         company_csv_path=company_list_path,
#         similarity_threshold=0.3
#     )
#     # [Step 3~6] Full 4-Stage Filtering Pipeline
#     pipeline_result = full_peer_filtering_pipeline(
#         target_pdf_path=pdf_path,
#         company_name=company_name,
#         raw_peer_names=raw_peers,
#         company_csv_path=company_list_path,
#         similarity_threshold=0.3
#     )
    
#     stage2_dec = pipeline_result.get("stage2_dec_passed", [])
#     stage2_profit = pipeline_result.get("stage2_profit_passed", [])
#     stage3_business = pipeline_result.get("stage3_business_passed", [])
#     stage4_final = pipeline_result.get("stage4_final_peers", [])
    
#     final_peers = stage4_final[:10]
#     peers_str = ", ".join(final_peers)

#     # [Step 7] RAG
#     if final_peers:
#         search_prompt = f"Target Peers: {peers_str}\n각 기업의 시가총액, PER, 2024년 당기순이익, Target 기업 '{company_name}'의 발행주식수를 검색하십시오."
#         rag_res = call_gemini(search_prompt, tools=[{"google_search": {}}])
#         rag_context = rag_res.get('text', '') if rag_res.get('ok') else ''
#     else:
#         rag_context = "Peer 데이터 부족"

#     # 🚨 [추가] 파이썬이 이미 확보한 PER 데이터를 LLM 프롬프트에 강제 주입하기 위한 텍스트 생성
#     scraped_per_info = []
#     if "details" in pipeline_result and "stage4_requirements" in pipeline_result["details"]:
#         for name, passed, reason, info in pipeline_result["details"]["stage4_requirements"]:
#             if name in final_peers:
#                 per = info.get("per", "N/A")
#                 scraped_per_info.append(f"- {name}: {per}")
#     scraped_per_text = "\n".join(scraped_per_info) if scraped_per_info else "확보된 PER 데이터 없음"
    
# # [Step 8] Valuation 계산
#     import datetime
#     target_year = datetime.datetime.now().year + 1
#     main_prompt = f"""
# 당신은 기업가치 평가 전문가(Valuation Analyst)입니다.

# [Input Data]
# - 대상 기업: {company_name}
# - 산업분류: {selected_industries}
# - 최종 Peer Group: {peers_str}
# - 시장 데이터 (RAG): {rag_context}
# - 🚨 **[시스템 확보 Peer PER 데이터 (최우선 활용)]**:
# {scraped_per_text}
# - **[보충 문서 데이터]**: {extra_text}

# [🚨 숫자 표기 및 계산 규칙 (매우 중요) 🚨]
# 1. **[숫자 포맷팅 - 콤마 위치 엄수]**: 숫자를 출력할 때는 **반드시 오른쪽에서부터 3자리마다 콤마(,)**를 정확히 찍으십시오. 한국어의 만(10,000) 단위에 헷갈려 4자리마다 콤마를 찍는 기형적 표기(예: 3,0462원, 1,6790원)를 **절대 금지**합니다.
#    - 올바른 표기: 30,462원, 16,790원, 1,520,000,000원
#    - 틀린 표기: 3,0462원, 1,6790원, 15,2000,0000원
# 2. **[금액 표기 규칙 - 소수점 절대 금지]**: PER 배수를 제외한 '기업가치', '주당 평가가액', '최종 공모가액' 등 모든 원(KRW) 단위 금액은 소수점 표기를 절대 금지합니다. 나눗셈 등에서 소수점이 발생하면 즉시 반올림하여 **무조건 정수(원)로만 표기**하십시오.
# 3. **[데이터 절대 신뢰 규칙]**: 제공된 [시스템 확보 최종 Peer 데이터]는 파이썬 시스템이 이미 비정상 수치(10 미만, 100 초과) 및 MAX/MIN 아웃라이어를 완벽하게 제거한 '최종 정제 데이터'입니다. **LLM은 어떠한 경우에도 임의로 기업을 제외하거나 필터링하는 행위를 절대 하지 마십시오.**

# [평가 로직 및 필수 계산 지침]
# 제공된 문서(메인 PDF, 보충 문서)와 시장 데이터를 바탕으로 다음 5단계를 반드시 직접 계산하십시오.

# 1. **Peer 평균 PER 산출**: 위 제공된 [시스템 확보 Peer PER 데이터]를 최우선으로 확인하여 PER 10 미만 또는 100 초과 기업을 제외한 후 산술평균을 구하십시오.
# 2. **Target 추정 순이익의 현재가치(PV) 산출**: 본격적인 IPO 시점을 고려하여, 여러 연도를 평균 내지 말고 반드시 분석 기준 연도 다음 해인 🚨 **'{target_year}년'의 추정 당기순이익 금액 하나만** 추출하여 현재가치로 할인하십시오. (만약 {target_year}년 데이터가 없다면 가장 인접한 미래 연도를 대안으로 사용하십시오.)
#    - [할인율 및 기간 가정]: IR 자료에 연 할인율이 없다면 **'20%' (0.20)를 기본 적용**하십시오. {target_year}년은 분석 시점(현재 연도)으로부터 1년 뒤이므로 할인 기간(n)은 1로 적용합니다. (대안 연도 사용 시 그에 맞춰 n을 조정)
#    - [현가 수식]: {target_year}년 추정 순이익 / (1 + 0.20)^1
#    - [최종 금액 절사]: 도출된 최종 현재가치 금액의 **1백만 원 미만 단위는 무조건 버림(절사)** 하십시오. (예: 13,258,810,729원 -> 13,258,000,000원)
# 3. **Target 기업가치 산출**: 2단계에서 구한 'Target 현재가치 순이익' × 1단계에서 구한 'Peer 평균 PER'
#    *(계산 예시: 10,000,000,000원 × 16.88배 = 168,800,000,000원)*
# 4. **주당 평가가액 산출**: 3단계의 Target 기업가치 / Target 발행주식수 
#    *(계산 예시: 168,800,000,000원 / 21,000주 = 8,038,095원 <- 소수점 반올림)*
# 5. **최종 희망공모가액 산출 (40% 할인율 고정)**: 4단계에서 구한 주당 평가가액 × (1 - 0.40)
#    - 여기서는 계산 내역만 작성하십시오. 파이썬에서 최종 100원 단위 절사를 수행할 것입니다.

# [중요 작성 지침]
# - Target 기업의 '{target_year}년 추정 순이익'과 '발행주식수'를 제공된 자료에서 반드시 찾아내십시오. (명시되어 있지 않다면 합리적 추정치를 가정하여 명시할 것)
# - **[산출 근거 가독성 및 완결성 원칙 (매우 중요)]**: `Calculation_Rationale` 필드를 작성할 때, 여러 경우의 수(예: 2개년 평균, 3개년 평균 등)를 테스트하며 늘어놓는 '중간 계산 과정(연습장)'을 밖으로 절대 노출하지 마십시오. 내부적으로 가장 합리적인 '단 하나의 최종 기준'을 확정한 뒤, **오직 그 최종 기준에 대한 구체적인 수식과 데이터만** 기재하십시오. 
# 단, 산출 근거가 너무 생략되어서는 안 되며, 심사역이 계산의 출처를 정확히 파악할 수 있도록 각 스텝별로 아래의 필수 디테일을 포함하여 최대 1~2줄로 깔끔하게 요약하십시오.
#   * 1단계 (Peer PER): 필터링을 통과한 대상 기업들의 구체적인 PER 수치들과 최종 산술 평균식
#   * 2단계 (추정 순이익): {target_year}년 추정 순이익 금액, 20% 할인율 반영 수식 및 최종 도출된 현재가치 금액 명시 (예: 100억 / (1+0.2)^1 = 83.3억)
#   * 3~5단계 (가치 산출): 도출된 숫자(순이익, PER, 주식수, 할인율 등)가 모두 포함된 명확한 사칙연산 수식 표기
# - **[계산 불가 시 빠른 중단]**: 만약 발행주식수 등 핵심 데이터가 없어 어차피 4~5번 단계의 산출이 불가능하다면, 앞선 1~2번 단계에서 억지로 복잡한 순이익 계산을 길게 늘어놓지 마십시오. 가능한 단계까지만 간결히 적고 "발행주식수 정보 부재로 주당 가치 산출 불가"라고 깔끔하게 마무리하십시오.
# - `Target_Net_Income`, `Total_Shares_Outstanding`, `Estimated_Share_Price` 필드에는 도출된 수치를 명확히 기재하십시오.
# - 기업의 전반적인 데이터를 바탕으로 'Three_Axis_Assessment' 필드에 기술성, 성장성, 회수성(Exit)을 각각 1~2문장으로 날카롭게 평가하십시오.
# - [핵심: 투자 리스크 도출]: `Investment_Risks` 필드에는 대상 기업의 시장 상황, 기술적 한계, 재무 상태 등을 종합적으로 분석하여 **가장 핵심적이고 치명적인 3가지 투자 리스크**를 도출하십시오. 각 리스크마다 구체적인 내용(`Risk_Detail`)과 이를 상쇄할 수 있는 회사의 전략 또는 애널리스트의 긍정적 전망(`Mitigation_and_Outlook`)을 반드시 짝지어 작성해야 합니다.
# - 기업의 재무 상황, 시장성, 기술력 그리고 산출된 가치를 모두 종합하여 'Investment_Rating' 필드에 최종 투자 매력도를 나타내는 종합 등급을 기재해 주십시오.
# - 'Final_Conclusion' 필드에는 위 모든 분석 결과를 바탕으로 한 최종 투자 결론을 3~4문장으로 요약해 주십시오.
# [Output Schema]
# {VALUATION_SCHEMA}
# """
#     res = call_gemini(main_prompt, pdf_path=pdf_path, max_tokens=8192)
#     val_data = safe_json_loads(res.get("text", "") if res.get("ok") else "{}")
    
#     # [Data Injection] 
#     if "Valuation_and_Judgment" not in val_data: val_data["Valuation_and_Judgment"] = {}
#     v_judge = val_data["Valuation_and_Judgment"]
#     if "Valuation_Logic_Detail" not in v_judge: v_judge["Valuation_Logic_Detail"] = {}
    
#     detail = v_judge["Valuation_Logic_Detail"]
#     detail["Step1_Industries"] = selected_industries
#     detail["Step2_Raw_Pool"] = raw_peers
#     detail["Step3_Dec_Filtered"] = stage2_dec
#     detail["Step3_Profit_Filtered"] = stage2_profit
#     detail["Step4_Business_Filtered"] = stage3_business
#     detail["Step5_Final_Peers"] = final_peers

#     # 🚨 [Python 절사 로직 추가] LLM이 계산한 산출 근거 텍스트 및 주당 가격에서 최종 금액을 찾아 100원 미만 절사 적용
#     rationale = detail.get("Calculation_Rationale", "")
    
#     if rationale and "최종 희망공모가액" in rationale:
#         try:
#             lines = rationale.split('\n')
#             new_lines = []
#             final_price_val = None
            
#             for line in lines:
#                 if "최종 희망공모가액" in line:
#                     val = None
#                     # 1. '=' 기호가 있는 경우, '=' 뒤의 첫 번째 숫자가 타겟 금액
#                     if "=" in line:
#                         result_part = line.split("=")[-1]
#                         # 콤마가 있든 없든, 소수점이 있든 없든 숫자를 완벽하게 잡아내는 정규식
#                         matches = re.findall(r'((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)', result_part)
#                         if matches:
#                             val = float(matches[0].replace(',', ''))
#                     else:
#                         # 2. '=' 기호가 없는 경우, 라인 내의 마지막 큰 숫자를 타겟 금액으로 추정
#                         matches = re.findall(r'((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)', line)
#                         nums = [float(m.replace(',', '')) for m in matches]
#                         nums = [n for n in nums if n > 100]  # "100원" 이라는 텍스트 자체를 무시하기 위한 장치
#                         if nums:
#                             val = nums[-1]
                            
#                     # 3. 파이썬의 무자비한 100원 미만 강제 절사 로직 적용
#                     if val is not None:
#                         truncated_val = math.floor(val / 100) * 100
#                         final_price_val = truncated_val
#                         # 원래 문자열(LLM이 쓴 글)의 바로 아랫줄에 파이썬이 강제로 결과값을 추가 기입
#                         line = line + f"\n      -> [100원 미만 절사] 최종 {truncated_val:,.0f}원"
                
#                 new_lines.append(line)
            
#             detail["Calculation_Rationale"] = "\n".join(new_lines)
            
#             # 워드 표(Table)에 들어갈 주당 가격(Estimated_Share_Price)도 절사된 금액으로 통일
#             if final_price_val is not None:
#                 detail["Estimated_Share_Price"] = f"{final_price_val:,.0f}원"
                
#         except Exception as e:
#             print(f"      [DEBUG] Python 100원 절사 로직 적용 실패: {e}")

#     if "details" in pipeline_result:
#         raw_s3 = pipeline_result["details"].get("stage3_similarity", [])[:5]
#         detail["Stage3_Similarity_Scores"] = [
#             (
#                 {"company": x.get("company") or x.get("name") or "", "score": x.get("score", 0.0), "reason": x.get("reason", ""), "main_products": x.get("main_products", "")}
#                 if isinstance(x, dict) else {"company": x[0] if len(x) > 0 else "", "score": x[1] if len(x) > 1 else 0.0, "reason": x[2] if len(x) > 2 else ""}
#             ) for x in raw_s3
#         ]
#         detail["Stage4_Requirements_Check"] = [
#             {"company": name, "passed": passed, "reason": reason, "info": info}
#             for name, passed, reason, info in pipeline_result["details"].get("stage4_requirements", [])[:10]
#         ]
    
#     if not detail.get("Calculation_Rationale") or "미정" in detail.get("Calculation_Rationale", ""):
#         detail["Calculation_Rationale"] = f"Peer({len(final_peers)}개사) 평균 PER 적용 및 40% 고정 할인율 적용하여 산출"
    
#     if not v_judge.get("Valuation_Table"):
#         # LLM이 도출(또는 파이썬이 절사)한 주당 평가가액을 테이블 코멘트에 반영
#         est_price = detail.get("Estimated_Share_Price", "-")
#         v_judge["Valuation_Table"] = [{ "Round": "IPO (예상)", "Pre_Money": "-", "Post_Money": "-", "Comment": f"예상 공모가: {est_price}" }]

#     return val_data