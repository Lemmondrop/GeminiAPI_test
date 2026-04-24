import os
import time
import json
import re
import csv
import base64
import requests
import random
import io
import pandas as pd
import concurrent.futures
import traceback # 에러 역추적용
from dotenv import load_dotenv

load_dotenv()
api_key = (os.getenv("GEMINI_API_KEY") or "").strip()

API_BASE = "https://generativelanguage.googleapis.com/v1beta"
TARGET_MODEL = "models/gemini-2.0-flash"
HEADERS = {"Content-Type": "application/json"}

# =========================================================
# 1. Helper Functions
# =========================================================
def _strip_code_fences(text: str) -> str:
    if not text: return ""
    text = re.sub(r"^```[a-zA-Z]*\s*", "", text.strip())
    return re.sub(r"\s*```$", "", text)

def safe_json_loads(text: str) -> dict:
    clean_text = _strip_code_fences(text)
    start = clean_text.find("{")
    end = clean_text.rfind("}")
    if start != -1 and end != -1:
        clean_text = clean_text[start:end+1]
    try:
        return json.loads(clean_text)
    except:
        return {}

def pdf_to_base64(pdf_path: str) -> dict:
    with open(pdf_path, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode("utf-8")
    return {"inline_data": {"mime_type": "application/pdf", "data": b64}}

# 🚨 [신규 추가] Pydantic의 $defs와 $ref를 Gemini REST API 규격에 맞게 쫙 펴주는 변환기
def convert_to_gemini_schema(schema_node, defs=None):
    if defs is None:
        defs = schema_node.get("$defs", {})
        
    # $ref(참조)가 있으면 원본 딕셔너리로 교체
    if "$ref" in schema_node:
        ref_name = schema_node["$ref"].split("/")[-1]
        return convert_to_gemini_schema(defs[ref_name], defs)
        
    # Optional 타입(anyOf) 처리
    if "anyOf" in schema_node:
        for sub in schema_node["anyOf"]:
            if sub.get("type") != "null":
                return convert_to_gemini_schema(sub, defs)
        return {"type": "STRING"} 
        
    out = {}
    t = schema_node.get("type", "OBJECT").upper()
    out["type"] = t
    
    if "description" in schema_node:
        out["description"] = schema_node["description"]
        
    if t == "OBJECT" and "properties" in schema_node:
        out["properties"] = {
            k: convert_to_gemini_schema(v, defs) 
            for k, v in schema_node["properties"].items()
        }
        if "required" in schema_node:
            out["required"] = schema_node["required"]
            
    if t == "ARRAY" and "items" in schema_node:
        out["items"] = convert_to_gemini_schema(schema_node["items"], defs)
        
    return out

# 🚨 [핵심 수정] response_schema 파라미터를 추가하여 Pydantic 모델을 수용할 수 있게 만듭니다.
def call_gemini(prompt: str, pdf_path: str = None, tools: list = None, response_schema=None, max_tokens: int = 8192) -> dict:
    if not api_key: raise RuntimeError("GEMINI_API_KEY is missing")
    url = f"{API_BASE}/{TARGET_MODEL}:generateContent?key={api_key}"
    
    parts = [{"text": prompt}]
    if pdf_path: parts.append(pdf_to_base64(pdf_path))
    
    # 기본 Configuration
    generation_config = {
        "temperature": 0.1, 
        "maxOutputTokens": max_tokens
    }
    
    # 🚨 Pydantic 스키마가 전달된 경우, API 페이로드에 JSON Schema 형태로 변환하여 강제 주입합니다.
    if response_schema:
        generation_config["responseMimeType"] = "application/json"
        raw_schema = response_schema.model_json_schema()
        # 🚨 [핵심 수정] 제미나이가 못 읽는 $defs 에러를 원천 차단하기 위해 변환기 적용
        generation_config["responseSchema"] = convert_to_gemini_schema(raw_schema)

    payload = {
        "contents": [{"parts": parts}], 
        "generationConfig": generation_config
    }
    if tools: payload["tools"] = tools
    
    for _ in range(3):
        try:
            resp = requests.post(url, headers=HEADERS, json=payload, timeout=180)
            if resp.status_code == 200:
                try: 
                    return {"ok": True, "text": resp.json()["candidates"][0]["content"]["parts"][0]["text"]}
                except: 
                    return {"ok": False, "error": "Parsing Error"}
            elif resp.status_code == 429: 
                time.sleep(5)
                continue
            else: 
                return {"ok": False, "error": resp.text}
        except Exception as e: 
            time.sleep(3)
            
    return {"ok": False, "error": "Timeout"}

# =========================================================
# 2. Industry Code Logic
# =========================================================
def load_industry_codes(csv_path: str):
    if not os.path.exists(csv_path):
        print(f"⚠️ [Error] 산업코드 파일을 찾을 수 없습니다: {csv_path}")
        return {}
    
    industry_map = {}
    encodings = ['utf-8-sig', 'cp949', 'euc-kr']
    
    for enc in encodings:
        try:
            with open(csv_path, 'r', encoding=enc) as f:
                reader = csv.DictReader(f)
                if reader.fieldnames:
                    reader.fieldnames = [h.strip() for h in reader.fieldnames]
                    
                name_col = next((c for c in reader.fieldnames if '산업내용' in c), None)
                code_col = None
                priority_cols = [c for c in reader.fieldnames if '산업분류코드' in c]
                
                if priority_cols: 
                    code_col = priority_cols[0]
                else: 
                    code_col = next((c for c in reader.fieldnames if '산업코드' in c), None)
                    
                if not name_col or not code_col: 
                    continue
                    
                for row in reader:
                    name = row.get(name_col, '').strip()
                    code = row.get(code_col, '').strip()
                    if name and code:
                        if name not in industry_map:
                            industry_map[name] = []
                        if code not in industry_map[name]:
                            industry_map[name].append(code)
                            
                if industry_map:
                    print(f"   ✅ Industry Codes Loaded: {len(industry_map)} unique industries")
                    break
        except: 
            continue
            
    return industry_map

def get_companies_by_code(target_code: str, csv_path: str):
    if not os.path.exists(csv_path): return []
    matched_companies = []
    encodings = ['utf-8-sig', 'cp949', 'euc-kr']
    target_clean = target_code.strip()
    for enc in encodings:
        try:
            with open(csv_path, 'r', encoding=enc) as f:
                reader = csv.DictReader(f)
                if reader.fieldnames:
                    reader.fieldnames = [h.strip() for h in reader.fieldnames]
                code_col = next((c for c in reader.fieldnames if '산업분류코드' in c), None)
                name_col = next((c for c in reader.fieldnames if '회사명' in c), None)
                if not code_col or not name_col: continue
                for row in reader:
                    row_code = row.get(code_col, '').strip()
                    row_name = row.get(name_col, '').strip()
                    if not row_code: continue
                    if row_code == target_clean: matched_companies.append(row_name)
                    elif len(target_clean) >= 3 and row_code.startswith(target_clean): matched_companies.append(row_name)
                    elif len(target_clean) >= 3 and target_clean in row_code: matched_companies.append(row_name)
                if matched_companies:
                    matched_companies = list(set(matched_companies))
                    break
        except: continue
    return matched_companies

# =========================================================
# 3. Financial Filtering Engine (Debug Mode)
# =========================================================
def get_random_ua():
    uas = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36'
    ]
    return random.choice(uas)

def check_net_income(company_info):
    name = company_info['name']
    code = company_info['code'] 
    
    time.sleep(random.uniform(0.5, 1.0))
    url = f"https://finance.naver.com/item/main.naver?code={code}"
    
    try:
        headers = {'User-Agent': get_random_ua(), 'Referer': 'https://finance.naver.com/'}
        res = requests.get(url, headers=headers, timeout=10)
        
        if res.status_code != 200:
            return (name, False, f"HTTP Error {res.status_code}")

        if len(res.text) < 1000:
            return (name, False, f"HTML 내용 너무 짧음 ({len(res.text)} bytes) - 차단 의심")

        try:
            dfs = pd.read_html(io.StringIO(res.text), attrs={"class": "tb_type1"}, match="매출액")
            if not dfs: 
                return (name, False, "재무 테이블(tb_type1) 없음")
            
            fin_df = dfs[0]
            cols = [str(c).replace(" ", "").replace("'", "").replace("(", "").replace(")", "").replace("\n", "") for c in fin_df.columns]

            target_idx = -1
            for i, c in enumerate(cols):
                if '2024.12' in c and 'E' not in c:
                    target_idx = i
                    break
            if target_idx == -1:
                for i, c in enumerate(cols):
                    if '2024.12' in c:
                        target_idx = i
                        break
            
            if target_idx == -1:
                return (name, False, f"2024년 컬럼 없음. 발견된 최근 컬럼: {cols[-3:]}")

            ni_idx = -1
            for idx, row in fin_df.iterrows():
                label = str(row.iloc[0]).replace(" ", "").strip()
                if '당기순이익(지배)' in label:
                    ni_idx = idx
                    break
                elif label == '당기순이익' and ni_idx == -1:
                    ni_idx = idx

            if ni_idx == -1:
                return (name, False, "당기순이익 행 없음")

            val_raw = fin_df.iloc[ni_idx, target_idx]
            
            def parse_val(v):
                s = str(v).strip()
                if s in ['-', 'nan', '', 'N/A']: return -999999
                try: return float(s.replace(',', ''))
                except: return -999999

            ni_val = parse_val(val_raw)

            if ni_val > 0:
                return (name, True, f"흑자 ({ni_val})")
            else:
                return (name, False, f"적자 ({ni_val})")

        except Exception as e:
            return (name, False, f"파싱 에러: {str(e)[:50]}")

    except Exception as e:
        return (name, False, f"접속 에러: {str(e)}")

def filter_peers_stage2(peer_names, company_csv_path):
    print(f"   📊 [Step 4~8] 재무 정밀 필터링 시작 (Input: {len(peer_names)}개 사)")
    
    dec_candidates = []
    dec_candidate_objs = []
    seen_codes = set()
    
    try:
        encodings = ['utf-8-sig', 'cp949', 'euc-kr']
        for enc in encodings:
            try:
                with open(company_csv_path, 'r', encoding=enc) as f:
                    reader = csv.DictReader(f)
                    if reader.fieldnames: reader.fieldnames = [h.strip() for h in reader.fieldnames]
                    for row in reader:
                        name = row.get('회사명', '').strip()
                        month = row.get('결산월', '').strip()
                        raw_code = row.get('종목코드', '').strip()
                        if name in peer_names:
                            if '12' in month and raw_code:
                                clean_code = raw_code.zfill(6)
                                if clean_code not in seen_codes:
                                    dec_candidates.append(name)
                                    dec_candidate_objs.append({'name': name, 'code': clean_code})
                                    seen_codes.add(clean_code)
                    if dec_candidates: break
            except: continue
    except: 
        return {"dec_passed": peer_names, "profit_passed": peer_names}

    print(f"      👉 12월 결산 & 코드 정제 완료: {len(dec_candidates)}개 사")
    if not dec_candidates: 
        return {"dec_passed": [], "profit_passed": []}

    print("      👉 당기순이익(지배) 흑자 여부 조회 중...")
    
    profit_passed = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        results = list(executor.map(check_net_income, dec_candidate_objs))
    
    for name, passed, reason in results:
        if passed:
            profit_passed.append(name)
    
    print(f"      👉 최종 통과: {len(profit_passed)}개 사")
    
    return {
        "dec_passed": dec_candidates,
        "profit_passed": profit_passed
    }