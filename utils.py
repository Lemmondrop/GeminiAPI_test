import os
import time
import json
import re
import csv
import base64
import requests
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

def call_gemini(prompt: str, pdf_path: str = None, tools: list = None, max_tokens: int = 8192) -> dict:
    if not api_key: raise RuntimeError("GEMINI_API_KEY is missing")
    
    url = f"{API_BASE}/{TARGET_MODEL}:generateContent?key={api_key}"
    parts = [{"text": prompt}]
    if pdf_path: parts.append(pdf_to_base64(pdf_path))
    
    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": max_tokens}
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
# 2. Local Data Loader (CSV) - Enhanced Debugging
# =========================================================
def load_industry_codes(csv_path: str):
    """
    [Step 1] preprocess_corp_code_list(prototype).csv 로드
    - 역할: 산업내용(Industry Name)과 산업분류코드(Code)의 매핑 정보 로드
    - [수정됨] '산업코드'보다 '산업분류코드'를 최우선으로 읽도록 강제
    """
    if not os.path.exists(csv_path):
        print(f"⚠️ [Error] 산업코드 파일을 찾을 수 없습니다: {csv_path}")
        return {}

    industry_map = {}
    encodings = ['utf-8-sig', 'cp949', 'euc-kr']
    
    for enc in encodings:
        try:
            with open(csv_path, 'r', encoding=enc) as f:
                reader = csv.DictReader(f)
                
                # 헤더 공백 제거
                if reader.fieldnames:
                    reader.fieldnames = [h.strip() for h in reader.fieldnames]
                    # print(f"   [Debug] Headers found: {reader.fieldnames}") # 디버깅용

                # 1. '산업내용' 컬럼 찾기
                name_col = next((c for c in reader.fieldnames if '산업내용' in c), None)
                
                # 2. '산업분류코드' 컬럼 찾기 (우선순위 적용)
                # - '산업분류코드'가 포함된 컬럼이 있으면 무조건 그것을 씀
                # - 없으면 Fallback으로 '산업코드'를 씀
                code_col = None
                priority_cols = [c for c in reader.fieldnames if '산업분류코드' in c]
                if priority_cols:
                    code_col = priority_cols[0]
                else:
                    code_col = next((c for c in reader.fieldnames if '산업코드' in c), None)

                if not name_col or not code_col:
                    continue
                
                # print(f"   [Debug] Selected Columns -> Name: {name_col}, Code: {code_col}")

                for row in reader:
                    name = row.get(name_col, '').strip()
                    code = row.get(code_col, '').strip()
                    if name and code:
                        industry_map[name] = code # Key: 산업내용 -> Value: 코드
                
                if industry_map:
                    print(f"   ✅ Industry Codes Loaded: {len(industry_map)} entries (Target Column: {code_col})")
                    break
        except: continue
        
    return industry_map

def get_companies_by_code(target_code: str, csv_path: str):
    """
    [Step 2] company_cord_prototype.csv 로드 및 필터링
    - target_code: 'C204' 또는 '204' 등
    - 매칭 로직을 강화하여 문자/숫자 혼용 처리
    """
    if not os.path.exists(csv_path):
        print(f"⚠️ [Error] 회사 리스트 파일을 찾을 수 없습니다: {csv_path}")
        return []

    matched_companies = []
    encodings = ['utf-8-sig', 'cp949', 'euc-kr']
    
    # 비교를 위해 타겟 코드에서 숫자만 추출할 수도 있음 (상황에 따라)
    target_clean = target_code.strip()

    for enc in encodings:
        try:
            with open(csv_path, 'r', encoding=enc) as f:
                reader = csv.DictReader(f)
                
                if reader.fieldnames:
                    reader.fieldnames = [h.strip() for h in reader.fieldnames]
                
                code_col = next((c for c in reader.fieldnames if '산업분류코드' in c), None)
                name_col = next((c for c in reader.fieldnames if '회사명' in c), None)

                if not code_col or not name_col:
                    continue

                for row in reader:
                    row_code = row.get(code_col, '').strip()
                    row_name = row.get(name_col, '').strip()
                    
                    if not row_code: continue

                    # [매칭 로직]
                    # 1. 완전 일치 (C20499 == C20499)
                    # 2. 전방 일치 (C20499 startswith C204)
                    # 3. 포함 관계 (C20499 contains 204)
                    
                    if row_code == target_clean:
                        matched_companies.append(row_name)
                    elif len(target_clean) >= 3 and row_code.startswith(target_clean):
                        matched_companies.append(row_name)
                    # 혹시 target이 '204'이고 row가 'C20499'인 경우 처리
                    elif len(target_clean) >= 3 and target_clean in row_code:
                        matched_companies.append(row_name)

                if matched_companies:
                    matched_companies = list(set(matched_companies))
                    print(f"   ✅ Found {len(matched_companies)} companies matching code '{target_clean}'")
                    break
        except: continue

    return matched_companies

def find_peers_by_standard_code(standard_code, company_list):
    """
    표준 코드(예: 259)를 사용하여 상장사 코드(예: C00259)를 가진 기업들을 찾음
    Logic: 상장사 코드의 숫자 부분이 표준 코드를 포함하거나 일치하는지 확인
    """
    peers = []
    target_num = re.sub(r'\D', '', str(standard_code)) # 숫자만 추출 (259)
    
    for comp in company_list:
        comp_code_num = re.sub(r'\D', '', comp['code']) # C00259 -> 00259
        
        # 매칭 로직:
        # 1. 표준 코드가 상장사 코드의 뒷부분과 일치 (예: 259 vs 00259 -> match)
        # 2. 또는 상장사 코드가 표준 코드를 포함 (유연한 매칭)
        if comp_code_num.endswith(target_num) or (len(target_num) >= 3 and target_num in comp_code_num):
            peers.append(comp['name'])
            
    return list(set(peers)) # 중복 제거