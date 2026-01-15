import os
import re
import json
import time
import random
import base64
import requests
from json import JSONDecodeError
from dotenv import load_dotenv

load_dotenv()
api_key = (os.getenv("GEMINI_API_KEY") or "").strip()

TARGET_MODEL = "models/gemini-2.0-flash" 
API_BASE = "https://generativelanguage.googleapis.com/v1beta"
HEADERS = {"Content-Type": "application/json"}

# =========================================================
# 1. Helper Functions (String & JSON Processing)
# =========================================================
def _strip_code_fences(text: str) -> str:
    if not text: return ""
    s = text.strip()
    if s.startswith("```"):
        s = re.sub(r"^```[a-zA-Z]*\s*", "", s)
        s = re.sub(r"\s*```$", "", s)
    return s.strip()

def _extract_first_json_object(text: str) -> str | None:
    if not text: return None
    s = _strip_code_fences(text)
    start = s.find("{")
    if start == -1: return None
    
    depth = 0
    in_str = False
    esc = False

    for i in range(start, len(s)):
        ch = s[i]
        if in_str:
            if esc: esc = False
            elif ch == "\\": esc = True
            elif ch == '"': in_str = False
            continue
        else:
            if ch == '"': in_str = True
            continue
            if ch == "{": depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0: return s[start:i+1]
    return None

def _safe_json_loads(text: str) -> dict:
    raw = _strip_code_fences(text)
    try:
        obj = json.loads(raw)
        return obj if isinstance(obj, dict) else {"_raw_json": obj}
    except:
        pass
    extracted = _extract_first_json_object(raw)
    if not extracted:
        return {} 
    try:
        obj = json.loads(extracted)
        return obj if isinstance(obj, dict) else {"_raw_json" : obj}
    except:
        return {}

def _post_gemini(payload: dict, timeout: int = 120, max_retries: int = 3, base_wait: int = 5) -> dict:
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY가 설정되지 않았습니다.")
    url = f"{API_BASE}/{TARGET_MODEL}:generateContent?key={api_key}"

    for attempt in range(max_retries):
        try:
            resp = requests.post(url, headers=HEADERS, json=payload, timeout=timeout)
        except requests.RequestException:
            time.sleep(base_wait + random.uniform(0, 2))
            continue

        if resp.status_code == 200:
            return {"ok": True, "json": resp.json()}
        
        if resp.status_code == 429:
            time.sleep(base_wait * (2 ** attempt))
            continue

        if resp.status_code >= 500:
            time.sleep(base_wait)
            continue
        
        return {"ok": False, "status": resp.status_code, "text": resp.text}

    return {"ok": False, "status": 0, "text": "Max retries exceeded"}

def _extract_text(res_json: dict) -> str:
    if not res_json: return ""
    try:
        return res_json["candidates"][0]["content"]["parts"][0]["text"].strip()
    except (KeyError, IndexError, TypeError):
        return ""
    except:
        return ""

def _pdf_part_from_path(pdf_path: str) -> dict:
    with open(pdf_path, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode("utf-8")
    return {
        "inline_data": {
            "mime_type": "application/pdf",
            "data": b64
        }
    }

# =========================================================
# 2. JSON Schema Definition
# =========================================================
JSON_SCHEMA = r"""
{
  "Report_Header": {
    "Company_Name": "기업명",
    "CEO_Name": "대표자명",
    "Industry_Sector": "세부 산업분야 (예: 반도체 장비, AI 솔루션)",
    "Industry_Classification": "산업 대분류 (바이오 / IT / 제조업 / 금융업 / 모빌리티 / 기타)",
    "Analyst": "LUCEN Investment Intelligence",
    "Investment_Rating": "강력 매수 / 긍정적 / 관망 / 부정적"
  },
  "Investment_Thesis_Summary": "핵심 투자 하이라이트 (3~5줄로 상세하게 기술)",
  
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
    "Key_Financial_Commentary": "재무 실적 분석 및 향후 추정치에 대한 상세 설명 (성장률, 이익률 등 포함)",
    "Investment_History": [
      { "Date": "YYYY.MM", "Round": "Series A 등", "Amount": "금액", "Investor": "투자자" }
    ],
    "Future_Revenue_Structure": {
        "Business_Model": "비즈니스 모델(BM) 및 수익 창출 구조 상세 서술",
        "Future_Cash_Cow": "향후 주력 캐시카우(Cash Cow) 및 이익 기여도 분석"
    }
  },
  "Growth_Potential": {
    "Target_Market_Analysis": {
        "Target_Area": "타겟 영역 정의 (구체적 시장 세분화)",
        "Market_Characteristics": "해당 시장의 주요 특성 및 진입 장벽",
        "Competitive_Positioning": "경쟁사 대비 포지셔닝 및 차별점"
    },
    "Target_Market_Trends": [
      { "Type": "기사/특허/인터뷰", "Source": "출처", "Content": "시장 동향 내용 요약" }
    ],
    "LO_Exit_Strategy": {
        "Verified_Signals": ["이미 검증된 시그널 1", "시그널 2 (레퍼런스)"],
        "Expected_LO_Scenarios": [
            { "Category": "구분 (예: 글로벌 제약사 기술이전)", "Probability": "가능성 (상/중/하)", "Comment": "상세 코멘트" }
        ],
        "Valuation_Range": "적정 가치 범위 (보수적 관점에서의 산정 근거)"
    },
    "Export_and_Contract_Stats": {
      "Export_Graph_Data": [["Year", "Value"]],
      "Contract_Count_Graph_Data": [["Year", "Count"]],
      "Sales_Graph_Data": [["Year", "Revenue"]]
    }
  },
  "Technology_and_Pipeline": {
      "Market_Pain_Points": [
          "기존 기술/시장의 한계점 1", 
          "미충족 수요(Unmet Needs) 2"
      ], 
      "Solution_and_Core_Tech": {
          "Technology_Name": "핵심 기술 또는 솔루션 명칭",
          "Key_Features": ["차별화된 기능 1", "기능 2"]
      },
      "Pipeline_Development_Status": {
          "Core_Platform_Details": "핵심 플랫폼 기술의 구동 원리 및 아키텍처 상세 설명",
          "Technical_Risk_Analysis": "개발 과정의 기술적 난관(Hurdle) 및 위험 요소 분석",
          "Technical_Conclusion": "기술 경쟁력 종합 평가 (경쟁우위 및 성공 가능성)"
      }
  },
  "Key_Personnel_and_Organization": {
      "CEO_Qualitative_Assessment": "대표이사의 전문성, 창업/Exit 경험, 리더십, 업계 네트워크 등 VC 관점의 정성적 평가",
      "Key_Manpower_Capabilities": "주요 C-Level 및 핵심 연구 인력의 이력, R&D 조직 구성 및 팀워크 강점 분석"
  },
  "Key_Risks_and_Mitigation": [ 
      { "Risk_Factor": "리스크 요인", "Mitigation_Strategy": "대응 방안" } 
  ],
  "Valuation_and_Judgment": {
      "Valuation_Table": [
          { "Round": "라운드", "Pre_Money": "Pre-Money", "Post_Money": "Post-Money", "Comment": "비고" }
      ],
      "Three_Axis_Assessment": {
          "Technology_Rating": "기술성 평가",
          "Growth_Rating": "성장성 평가",
          "Exit_Rating": "회수 가능성 평가"
      },
      "Suitable_Investor_Type": "적합 투자자 유형"
  },
  "Final_Conclusion": "종합 투자의견 (한 문단으로 상세히)"
}
""".strip()

# =========================================================
# 3. RAG & Data Validation Helpers
# =========================================================
def growth_rag_prompt(base_obj: dict) -> str:
    header = base_obj.get("Report_Header") or {}
    company = header.get("Company_Name", "해당 기업")
    
    fin_status = base_obj.get("Financial_Status") or {}
    existing_history = fin_status.get("Investment_History") or []
    
    return f"""
오직 JSON만 출력.
'{company}'의 부족한 재무 및 투자 정보를 웹 검색을 통해 상세하게 보강하십시오.

[검색 목표]
1. **투자 유치 이력 (상세)**: 설립 이후 모든 투자 라운드(Seed, Series A/B 등), 금액, 참여 투자자(VC) 리스트를 완성하십시오.
2. **최근 실적 및 전망**: 최근 3~5년 매출액, 영업이익 추이와 향후 성장 전망 수치를 검색하여 채우십시오.

[현재 보유 데이터]
투자이력: {json.dumps(existing_history, ensure_ascii=False)}

[Output JSON Schema]
{{
  "Financial_Status": {{
    "Investment_History": [
      {{ "Date": "YYYY.MM", "Round": "라운드", "Amount": "금액", "Investor": "투자자" }}
    ]
  }},
  "Growth_Potential": {{
    "Export_and_Contract_Stats": {{
      "Sales_Graph_Data": [ ["Year", "Revenue"] ]
    }}
  }}
}}
""".strip()

def validate_growth_data(base_obj: dict) -> dict:
    if not isinstance(base_obj, dict): return base_obj
    
    gp = base_obj.get("Growth_Potential") or {}
    stats = gp.get("Export_and_Contract_Stats") or {}
    
    def normalize_chart_list(data_list, header_name):
        default = [["Year", header_name]]
        if not isinstance(data_list, list) or len(data_list) < 2:
            return default
        
        clean_rows = [data_list[0]]
        for row in data_list[1:]:
            if isinstance(row, list) and len(row) >= 2:
                try:
                    y = str(row[0]).replace("년","").strip()
                    val_str = str(row[1]).replace(",","").replace("억","").replace("원","").strip()
                    v = 0 if val_str in ["N/A", "", "-"] else float(val_str)
                    clean_rows.append([y, v])
                except:
                    pass
        return clean_rows if len(clean_rows) > 1 else default

    if isinstance(stats, dict):
        stats["Export_Graph_Data"] = normalize_chart_list(stats.get("Export_Graph_Data"), "Export_Value")
        stats["Contract_Count_Graph_Data"] = normalize_chart_list(stats.get("Contract_Count_Graph_Data"), "Count")
        stats["Revenue_Graph_Data"] = normalize_chart_list(stats.get("Revenue_Graph_Data"), "Revenue")
        gp["Export_and_Contract_Stats"] = stats
        base_obj["Growth_Potential"] = gp

    return base_obj

def merge_growth_info(base_obj: dict, patch: dict) -> dict:
    if not isinstance(base_obj, dict) or not isinstance(patch, dict):
        return base_obj

    base_fin = base_obj.setdefault("Financial_Status", {})
    patch_fin = patch.get("Financial_Status") or {}
    
    if patch_fin.get("Investment_History"):
        curr_hist = base_fin.get("Investment_History") or []
        # RAG 결과가 더 충실하면 교체
        if len(patch_fin["Investment_History"]) > len(curr_hist):
             base_fin["Investment_History"] = patch_fin["Investment_History"]

    base_gp = base_obj.setdefault("Growth_Potential", {})
    patch_gp = patch.get("Growth_Potential") or {}
    
    if patch_gp:
        if patch_gp.get("Target_Market_Trends"):
            if not base_gp.get("Target_Market_Trends"):
                base_gp["Target_Market_Trends"] = patch_gp["Target_Market_Trends"]
            
        if patch_gp.get("Export_and_Contract_Stats"):
            base_stats = base_gp.setdefault("Export_and_Contract_Stats", {})
            patch_stats = patch_gp["Export_and_Contract_Stats"]
            for key in ["Export_Graph_Data", "Contract_Count_Graph_Data", "Sales_Graph_Data"]:
                if patch_stats.get(key) and len(patch_stats[key]) > 1:
                    curr_stats = base_stats.get(key) or []
                    if len(curr_stats) <= 1:
                        base_stats[key] = patch_stats[key]

    return validate_growth_data(base_obj)

# =========================================================
# 4. Main Pipeline Logic
# =========================================================
def refine_pdf_to_json_onecall(
    pdf_path: str,
    json_schema: str = JSON_SCHEMA,
    max_output_tokens: int = 16384,
    retry_max_output_tokens: int = 32768,
    enable_rag: bool = True
) -> dict:

    prompt_pdf = f"""
# [Role]
당신은 'LUCEN Investment Intelligence'의 수석 심사역입니다. 
제공된 IR 자료(PDF)를 심층 분석하여, 투자자가 의사결정을 내릴 수 있는 **상세하고 구체적인** 보고서 데이터를 JSON으로 추출하십시오.

# [Extraction Rules - Very Important]
1. **텍스트 데이터 정밀 추출**: 
   - 문서 내 기술, 시장, 리스크 정보를 요약하십시오.
   - 'Industry_Classification' (산업 대분류)을 반드시 기입하십시오.
   
2. **재무 데이터 "Fact-Check"**:
   - 텍스트 표, 이미지 표, 그래프를 모두 분석하여 'Financial_Status'를 채우십시오.
   - **과거 실적이 없으면 '추정치(Forecast/Plan)'라도 반드시 추출하십시오.** (예: 2025(E))
   - 데이터가 흩어져 있어도 연도별로 모아서 'Income_Statement_Summary'에 기입하십시오.
   - **절대 'null'을 쉽게 반환하지 마십시오.**

3. **비정형/시각 데이터 해석**:
   - **투자 이력**: 문서 내 'History', 'Timeline' 섹션의 다이어그램을 해석하여 투자 라운드, 금액, 투자자를 추출하십시오.
   - **그래프**: 막대 그래프의 높이를 시각적으로 해석하여 대략적인 수치라도 'Graph_Data'에 입력하십시오.

# [Output Schema]
{json_schema}
""".strip()

    r1 = _post_gemini({
        "contents": [{
            "parts": [
                {"text": prompt_pdf},
                _pdf_part_from_path(pdf_path)
            ]
        }],
        "generationConfig": {"temperature": 0.2, "maxOutputTokens": max_output_tokens}
    }, timeout=240)

    if not r1.get("ok"):
        return {"error": r1.get("status"), "message": r1.get("text")}

    text1 = _extract_text(r1["json"])
    base_obj = _safe_json_loads(text1)
    
    if not isinstance(base_obj, dict):
        base_obj = {}

    # RAG Enrichment
    if enable_rag:
        try:
            rag_p = growth_rag_prompt(base_obj)
            r2 = _post_gemini({
                "contents": [{"parts": [{"text": rag_p}]}],
                "tools": [{"google_search": {}}],
                "generationConfig": {"temperature": 0.1, "maxOutputTokens": 8192}
            }, timeout=180)

            if r2.get("ok"):
                text2 = _extract_text(r2["json"])
                patch = _safe_json_loads(text2)
                base_obj = merge_growth_info(base_obj, patch)
        except Exception as e:
            print(f"  [RAG Warning] {e}")

    return base_obj