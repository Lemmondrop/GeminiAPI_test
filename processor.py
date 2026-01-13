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

# 색상 상수
COLOR_RED = "indianred"
COLOR_YELLOW = "khaki"
COLOR_BLUE = "cornflowerblue"

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
# JSON Schema
# =========================================================
JSON_SCHEMA = r"""
{
  "Report_Header": {
    "Company_Name": "기업명",
    "CEO_Name": "대표자명",
    "Industry_Sector": "산업분야",
    "Analyst": "LUCEN Investment Intelligence",
    "Investment_Rating": "강력 매수 / 긍정적 / 관망 / 부정적"
  },
  "Investment_Thesis_Summary": "핵심 투자 하이라이트 (3줄 요약)",
  
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
    "Key_Financial_Commentary": "재무 실적 및 향후 전망 요약",
    "Investment_History": [
      { "Date": "YYYY.MM", "Round": "Series A 등", "Amount": "금액", "Investor": "투자자" }
    ]
  },

  "Growth_Potential": {
    "Target_Market_Trends": [
      { "Type": "기사/특허/인터뷰", "Source": "출처", "Content": "내용" }
    ],
    "Export_and_Contract_Stats": {
      "Export_Graph_Data": [["Year", "Value"]],
      "Contract_Count_Graph_Data": [["Year", "Count"]],
      "Sales_Graph_Data": [["Year", "Revenue"]]
    }
  },
  "Problem_and_Solution": { "Market_Pain_Points": [".."], "Solution_Value_Prop": [".."] },
  "Technology_and_Moat": { "Core_Technology_Name": "..", "Technical_Details": [".."] },
  "Key_Risks_and_Mitigation": [ { "Risk_Factor": "..", "Mitigation_Strategy": ".." } ],
  "Final_Conclusion": "종합 의견"
}
""".strip()

# =========================================================
# ✅ [Fix] Null Safe Access Helpers
# =========================================================
def growth_rag_prompt(base_obj: dict) -> str:
    # .get() 뒤에 or {}를 붙여 None이 반환되어도 빈 딕셔너리로 처리되게 함 (Null Safety)
    header = base_obj.get("Report_Header") or {}
    company = header.get("Company_Name", "해당 기업")
    
    fin_status = base_obj.get("Financial_Status") or {}
    existing_history = fin_status.get("Investment_History") or []
    
    return f"""
오직 JSON만 출력.
'{company}'의 부족한 재무 및 투자 정보를 검색하여 보강하십시오.

[검색 목표]
1. **투자 유치 이력**: 설립 이후 모든 투자 라운드(Seed, Series A 등), 금액, 투자자 정보를 검색하여 리스트를 완성하십시오.
2. **최근 실적 및 전망**: 최근 3년 매출액/영업이익 추이와 향후 성장 전망(수출, 계약 등)을 검색하십시오.

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
    
    # None 체크 강화
    gp = base_obj.get("Growth_Potential") or {}
    if not isinstance(gp, dict): return base_obj
    
    stats = gp.get("Export_and_Contract_Stats") or {}
    if not isinstance(stats, dict): return base_obj

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
                    # 값이 N/A거나 비어있으면 0으로 처리하여 그래프 오류 방지
                    v = 0 if val_str in ["N/A", "", "-"] else float(val_str)
                    clean_rows.append([y, v])
                except:
                    pass
        return clean_rows if len(clean_rows) > 1 else default

    stats["Export_Graph_Data"] = normalize_chart_list(stats.get("Export_Graph_Data"), "Export_Value")
    stats["Contract_Count_Graph_Data"] = normalize_chart_list(stats.get("Contract_Count_Graph_Data"), "Count")
    stats["Revenue_Graph_Data"] = normalize_chart_list(stats.get("Revenue_Graph_Data"), "Revenue")
    
    return base_obj

def merge_growth_info(base_obj: dict, patch: dict) -> dict:
    if not isinstance(base_obj, dict) or not isinstance(patch, dict):
        return base_obj

    # 1. Financial_Status 병합 (Null Safety 강화)
    base_fin = base_obj.get("Financial_Status")
    if not isinstance(base_fin, dict):
        base_fin = {}
        base_obj["Financial_Status"] = base_fin
    
    patch_fin = patch.get("Financial_Status") or {}
    
    # 투자 이력 병합
    if patch_fin.get("Investment_History"):
        curr_hist = base_fin.get("Investment_History") or []
        # RAG 결과가 더 많으면 덮어쓰기
        if len(patch_fin["Investment_History"]) > len(curr_hist):
             base_fin["Investment_History"] = patch_fin["Investment_History"]

    # 2. Growth Potential 병합 (Null Safety 강화)
    base_gp = base_obj.get("Growth_Potential")
    if not isinstance(base_gp, dict):
        base_gp = {}
        base_obj["Growth_Potential"] = base_gp
    
    patch_gp = patch.get("Growth_Potential") or {}
    
    if patch_gp:
        if patch_gp.get("Target_Market_Trends"):
            if not base_gp.get("Target_Market_Trends"):
                base_gp["Target_Market_Trends"] = patch_gp["Target_Market_Trends"]
            
        if patch_gp.get("Export_and_Contract_Stats"):
            base_stats = base_gp.get("Export_and_Contract_Stats")
            if not isinstance(base_stats, dict):
                base_stats = {}
                base_gp["Export_and_Contract_Stats"] = base_stats
            
            patch_stats = patch_gp["Export_and_Contract_Stats"]
            for key in ["Export_Graph_Data", "Contract_Count_Graph_Data", "Sales_Graph_Data"]:
                if patch_stats.get(key) and len(patch_stats[key]) > 1:
                    curr_stats = base_stats.get(key) or []
                    if len(curr_stats) <= 1:
                        base_stats[key] = patch_stats[key]

    return validate_growth_data(base_obj)

# =========================================================
# Main Logic
# =========================================================
def refine_pdf_to_json_onecall(
    pdf_path: str,
    json_schema: str = JSON_SCHEMA,
    max_output_tokens: int = 16384,
    retry_max_output_tokens: int = 32768,
    enable_rag: bool = True
) -> dict:

    # 프롬프트: 텍스트 추출의 정교함 유지 + 비정형/추정 데이터 추출 강화
    prompt_pdf = f"""
# [Role]
당신은 VC 수석 심사역입니다. IR 자료(PDF)를 정밀 분석하여 JSON 데이터를 추출하십시오.

# [Extraction Guidelines]
1. **텍스트 기반 정밀 추출 (Text Data)**:
   - 문서 내에 존재하는 텍스트 테이블(재무상태표, 손익계산서 등)은 수치를 정확하게 옮기십시오.
   - 기술 설명, 시장 문제점 등 텍스트 정보는 요약하여 추출하십시오.

2. **추정 재무제표 필독 (Estimated Financials)**:
   - 과거 실적뿐만 아니라 **'(E)', 'Plan', 'Forecast'** 등이 포함된 **미래 추정 손익계산서/재무상태표**가 있다면 반드시 추출하십시오.
   - 예: 2025(E), 2026(Plan) 등의 컬럼이 있으면 Income_Statement_Summary에 포함시킬 것.

3. **비정형/시각 데이터 해석 (Visual/Unstructured Data)**:
   - 텍스트로 명시되지 않고 **도표, 다이어그램, 그래프**로만 존재하는 정보를 데이터화 하십시오.
   - **투자 이력(Investment History)**: 'History', 'Milestone', 'Funding' 관련 타임라인이나 다이어그램(예: Seed -> Series A 흐름도)을 해석하여 Date, Round, Amount, Investor를 추출하십시오.
   - **그래프(Charts)**: 매출/수출 추이 막대 그래프 등을 발견하면, 각 막대의 높이를 근사치 숫자로 변환하여 'Graph_Data'에 입력하십시오.

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
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": max_output_tokens}
    }, timeout=240)

    if not r1.get("ok"):
        return {"error": r1.get("status"), "message": r1.get("text")}

    text1 = _extract_text(r1["json"])
    base_obj = _safe_json_loads(text1)
    
    # 1차 파싱 실패 시 빈 객체로 진행 (NoneType 방지)
    if not isinstance(base_obj, dict):
        base_obj = {}

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