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

TARGET_MODEL = "models/gemini-2.5-flash"
API_BASE = "https://generativelanguage.googleapis.com/v1beta"
HEADERS = {"Content-Type": "application/json"}

def _strip_code_fences(text: str) -> str:
    """
    ```json ... ``` 또는 ``` ... ``` 형태 제거
    """
    if not text:
        return ""
    s = text.strip()
    if s.startswith("```"):
        # 시작 코드펜스 제거 (```json 또는 ``` 등)
        s = re.sub(r"^```[a-zA-Z]*\s*", "", s)
        # 끝 코드펜스 제거
        s = re.sub(r"\s*```$", "", s)
    return s.strip()

def _extract_first_json_object(text: str) -> str | None:
    """
    텍스트 안에서 첫 번째 JSON 객체({ ... })를 괄호 밸런싱으로 추출
    - 문자열 내부의 중괄호는 무시
    """
    if not text:
        return None
    
    s = _strip_code_fences(text)
    start = s.find("{")
    if start == -1:
        return None
    
    depth = 0
    in_str = False
    esc = False

    for i in range(start, len(s)):
        ch = s[i]

        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        else:
            if ch == '"':
                in_str = True
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return s[start:i+1]
    return None

def _safe_json_loads(text: str) -> dict:
    """
    1) 전체 텍스트를 그대로 JSON 파싱 시도
    2) 실패하면, 첫 번째 JSON 객체만 추출하여 재시도
    """
    raw = _strip_code_fences(text)

    # 1차 : 전체 파싱
    try:
        obj = json.loads(raw)
        return obj if isinstance(obj, dict) else {"_raw_json": obj}
    except:
        pass

    # 2차 : JSON 객체 덩어리 추출 후 파싱
    extracted = _extract_first_json_object(raw)
    if not extracted:
        raise JSONDecodeError("No JSON object found", raw, 0)
    
    obj = json.loads(extracted)
    return obj if isinstance(obj, dict) else {"_raw_json" : obj}

def _debug_dump_response(tag: str, res_json: dict):
    """
    Gemini 원본 응답을 debug_dump 폴더에 저장
    """
    try:
        os.makedirs("debug_dump", exist_ok=True)
        safe_tag = str(tag).replace("\\", "_").replace("/", "_").replace(":", "_")
        path = os.path.join("debug_dump", f"{safe_tag}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(res_json, f, ensure_ascii=False, indent=2)
        print(f"   [DEBUG] Raw Gemini response dumped → {path}")
    except Exception as e:
        print(f"   [DEBUG ERROR] dump 실패: {e}")

def _post_gemini(payload: dict, timeout: int = 120, max_retries: int = 5, base_wait: int = 10) -> dict:
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY가 설정되지 않았습니다.")

    url = f"{API_BASE}/{TARGET_MODEL}:generateContent?key={api_key}"

    for attempt in range(max_retries):
        try:
            resp = requests.post(url, headers=HEADERS, json=payload, timeout=timeout)
        except requests.RequestException:
            wait = 5 + random.uniform(0, 2)
            time.sleep(wait)
            continue

        if resp.status_code == 200:
            return {"ok": True, "json": resp.json()}

        if resp.status_code == 429:
            wait = base_wait * (2 ** attempt) + random.uniform(0, 3)
            time.sleep(wait)
            continue

        if resp.status_code >= 500:
            wait = 5 + random.uniform(0, 2)
            time.sleep(wait)
            continue

        return {"ok": False, "status": resp.status_code, "text": resp.text}

    return {"ok": False, "status": 0, "text": "Max retries exceeded"}

def _extract_text(res_json: dict) -> str:
    """
    Gemini 응답에서 text 추출 (모든 케이스 로깅)
    """
    if not res_json:
        print("   [DEBUG] res_json is None or empty")
        return ""

    candidates = res_json.get("candidates", [])
    if not candidates:
        print("   [DEBUG] candidates 없음")
        return ""

    cand0 = candidates[0]

    # finishReason 확인
    finish_reason = cand0.get("finishReason")
    if finish_reason:
        print(f"   [DEBUG] finishReason = {finish_reason}")

    content = cand0.get("content")
    if not content:
        print("   [DEBUG] content 없음")
        return ""

    parts = content.get("parts", [])
    if not parts:
        print("   [DEBUG] parts 비어있음")
        return ""

    texts = []
    for i, p in enumerate(parts):
        if isinstance(p, dict):
            if "text" in p:
                t = p["text"]
                if isinstance(t, str) and t.strip():
                    texts.append(t.strip())
                else:
                    print(f"   [DEBUG] part[{i}] text는 있으나 비어있음")
            else:
                print(f"   [DEBUG] part[{i}]에 text 키 없음: {list(p.keys())}")
        else:
            print(f"   [DEBUG] part[{i}]가 dict 아님: {type(p)}")

    joined = "\n".join(texts).strip()
    if not joined:
        print("   [DEBUG] 최종 추출 text = EMPTY")

    return joined

def _pdf_part_from_path(pdf_path: str) -> dict:
    """
    Gemini contents.parts 에 넣을 PDF inline_data 파트 생성
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")

    with open(pdf_path, "rb") as f:
        data = f.read()

    b64 = base64.b64encode(data).decode("utf-8")
    return {
        "inline_data": {
            "mime_type": "application/pdf",
            "data": b64
        }
    }

JSON_SCHEMA = r"""
{
  "Report_Header": {
    "Company_Name": "기업명",
    "CEO_Name": "성명",
    "Industry_Sector": "산업분야",
    "Analyst": "LUCEN Investment Intelligence",
    "Investment_Rating": "강력 매수 / 긍정적 / 관망 / 부정적"
  },
  "Investment_Thesis_Summary": "핵심 투자 하이라이트 (3줄 요약)",
  "Problem_and_Solution": {
    "Market_Pain_Points": ["문제점1", "문제점2"],
    "Solution_Value_Prop": ["해결책", "차별점"]
  },
  "Technology_and_Moat": {
    "Core_Technology_Name": "핵심기술명",
    "Technical_Details": ["원리", "특허/인증 등"],
    "Moat_Analysis": ["진입장벽", "Lock-in요소"]
  },
  "Market_Opportunity": {
    "TAM_SAM_SOM_Text": "시장규모 요약",
    "Market_Chart_Data": [
      ["구분", "시장규모(단위 포함)"],
      ["TAM", 0],
      ["SAM", 0],
      ["SOM", 0]
    ]
  },
  "Business_Financial_Status": {
    "Revenue_Model": "수익모델",
    "Current_Status": "사업단계",
    "Financial_Highlights": ["매출/투자 수치", "목표치"]
  },
  "Table_Data_Preview": {
      "Major_Milestones": [["시기", "내용"]],
      "Financial_Table": [
          ["항목", "연도1", "연도2"],
          ["매출액", "값1", "값2"],
          ["영업이익", "값1", "값2"]
      ]
  },
  "Key_Risks_and_Mitigation": [
    { "Risk_Factor": "리스크", "Mitigation_Strategy": "대응책" }
  ],
  "Due_Diligence_Questions": ["질문1", "질문2", "질문3"],
  "Final_Conclusion": "종합 투자의견"
}
""".strip()

def refine_pdf_to_json_onecall(
    pdf_path: str,
    json_schema: str,
    max_output_tokens: int = 16384,          # 기본 상향
    retry_max_output_tokens: int = 32768     # 재시도는 더 크게
) -> dict:

    def call_once(prompt: str, tag: str, tokens: int):
        payload = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    _pdf_part_from_path(pdf_path),
                ]
            }],
            "tools": [{"google_search": {}}],
            "generationConfig": {
                "temperature": 0.1,
                "maxOutputTokens": tokens,
            }
        }
        r = _post_gemini(payload, timeout=240)
        if not r.get("ok"):
            return {"ok": False, "error": f"API Error {r.get('status')}", "message": r.get("text"), "raw": None}

        _debug_dump_response(f"{tag}_{os.path.basename(pdf_path)}", r["json"])
        out_text = _extract_text(r["json"])
        finish = (((r["json"].get("candidates") or [{}])[0]).get("finishReason") or "")
        return {"ok": True, "out_text": out_text, "finish": finish, "raw": r["json"]}

    prompt_normal = f"""
오직 JSON만 출력(설명/마크다운/코드펜스/주석 금지).
JSON은 {{ 로 시작하고 }} 로 끝나야 함.
값 없으면 "N/A" 또는 0.

[재무 데이터 완전성(최우선)]
- PDF 재무표의 '모든 연도 컬럼'을 Financial_Table 헤더에 그대로 생성(축소 금지)
- 매출액/영업이익은 반드시 채우기(없으면 N/A)

[검색]
- PDF에 없는 내용만 google_search로 보완
- 검색 근거는 출처(기관/매체/연도)를 문장에 포함

[Output JSON Schema]
{json_schema}

첨부 PDF 기반으로 JSON 생성.
""".strip()

    r1 = call_once(prompt_normal, "onecall", max_output_tokens)
    if not r1.get("ok"):
        return {"error": r1["error"], "message": r1["message"], "stage": "onecall"}

    if not r1.get("out_text"):
        return {"error": "EmptyText", "message": "응답 텍스트가 비어있습니다.", "stage": "onecall"}

    # MAX_TOKENS 아니면 그대로 파싱
    if r1["finish"] != "MAX_TOKENS":
        try:
            return _safe_json_loads(r1["out_text"])
        except JSONDecodeError:
            return {"error": "JSONDecodeError", "message": "JSON 파싱 실패", "preview": r1["out_text"][:4000], "stage": "onecall"}

    # MAX_TOKENS면: retry_compact + 토큰 더 크게
    prompt_compact = f"""
오직 JSON만 출력. (설명/코드펜스/주석 금지)
JSON은 {{ 로 시작하고 }} 로 끝나야 함.
값 없으면 "N/A" 또는 0.

[초압축]
- 리스트 항목은 각 3개 이내
- 각 문자열은 120자 이내

[재무 최우선]
- Financial_Table 연도 컬럼은 PDF에 있는 모든 연도 포함(축소 금지)
- 표가 길면 '매출액/영업이익' 중심으로라도 연도별 채우기

[검색]
PDF에 없는 내용만 검색, 출처(기관/매체/연도) 포함.

[Output JSON Schema]
{json_schema}
""".strip()

    r2 = call_once(prompt_compact, "onecall_retry_compact", retry_max_output_tokens)
    if not r2.get("ok"):
        return {"error": r2["error"], "message": r2["message"], "stage": "onecall_retry_compact"}

    if not r2.get("out_text"):
        return {"error": "EmptyText", "message": "재시도 응답 텍스트가 비어있습니다.", "stage": "onecall_retry_compact"}

    if r2["finish"] == "MAX_TOKENS":
        return {
            "error": "TruncatedOutput",
            "message": f"재시도도 MAX_TOKENS로 잘렸습니다. maxOutputTokens 상향 필요. (현재 retry={retry_max_output_tokens})",
            "preview": r2["out_text"][:2000],
            "stage": "onecall_retry_compact"
        }

    try:
        return _safe_json_loads(r2["out_text"])
    except JSONDecodeError:
        return {"error": "JSONDecodeError", "message": "재시도 JSON 파싱 실패", "preview": r2["out_text"][:4000], "stage": "onecall_retry_compact"}

# ==========================
# 1단계: PDF -> 핵심 사실/수치/근거 추출(LLM 호출, 비검색)
# ==========================
def stage1_extract_json_from_pdf(pdf_path: str, json_schema: str, max_output_tokens: int = 8192) -> dict:
    """
    1단계: PDF 첨부 + Dynamic Retrieval + JSON mode
    - 과거에 잘 되던 방식: 1단계에서 곧바로 스키마 JSON 생성
    """
    prompt = f"""
# [Role & Context]
당신은 15년 경력의 VC 수석 심사역 'LUCEN Investment Intelligence'입니다.
제공된 IR 자료(PDF)와 Google Search(동적 검색)를 활용하여, 투자심사보고서용 데이터를 JSON으로 추출하십시오.

# [Dynamic Analysis Guidelines (with Web Search)]
1. 산업별 맥락 인식: Bio, AI, Manufacturing 등 산업군을 파악해 KPI를 추출하십시오.
2. Fact-Checking: IR 자료를 우선하되, 시장 규모나 경쟁사 정보가 부족하면 Google Search로 보완하십시오.
3. 재무 데이터 완전성 (Critical):
   - [절대 요약 금지] 문서 내 재무제표(손익계산서/추정치 포함)에 기재된 '모든 연도' 데이터를 빠짐없이 추출하십시오.
   - 예: 2021~2028이 있으면 JSON에도 2021~2028 컬럼이 모두 있어야 합니다. 임의 축소 금지.
   - 표가 여러 개면 가능한 한 모두 반영하고, 최소한 '매출액/영업이익'은 반드시 채우십시오.
4. 시장 규모: TAM/SAM/SOM 수치가 없으면 검색을 통해 추산하고, 근거를 텍스트로 남기십시오.

# [Extraction Rules]
- 반드시 JSON만 출력 (JSON 외 텍스트 금지)
- 숫자는 가능한 한 숫자 형태로(불가하면 문자열)
- 없으면 "N/A" 또는 0 사용(추정 금지)
- Financial_Table 헤더 연도 컬럼은 문서에 등장하는 연도 개수만큼 생성

# [Output JSON Schema]
{json_schema}
""".strip()

    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                _pdf_part_from_path(pdf_path),
            ]
        }],
        # "tools": [{"google_search": {}}], 1단계에서는 검색 없이 수행
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
            "maxOutputTokens": max_output_tokens
        }
    }

    r = _post_gemini(payload, timeout=180)
    if not r.get("ok"):
        return {"error": f"API Error {r.get('status')}", "message": r.get("text")}

    _debug_dump_response(f"stage1_{os.path.basename(pdf_path)}", r["json"])
    
    out_text = _extract_text(r["json"])
    if not out_text:
        return {"error": "EmptyText", "message": "1단계(JSON) 응답 텍스트가 비어있습니다."}

    try:
        return {"ok": True, "json": json.loads(out_text)}
    except JSONDecodeError:
        return {"error": "JSONDecodeError", "message": "1단계 JSON 파싱 실패", "preview": out_text[:800]}


# ==========================
# 2단계: Web RAG 보강(LLM + google_search)
# ==========================
def stage2_enrich_with_web_rag(material_text: str, max_chars: int = 25000) -> dict:
    clipped = (material_text or "").strip()[:max_chars]

    stage2_prompt = f"""
아래 '정제용 원자료'를 바탕으로, IR에서 부족한 정보를 웹 검색으로 보완해 주세요.
특히 다음을 우선 보강:
- 연도별 매출(최근 3~5개 연도), 근거 출처(기사/공시/홈페이지 등)
- 주요기술의 수출/해외 진출 현황 기사(파트너/지역/시점)
- 시장규모(TAM/SAM/SOM에 쓸 수치)와 출처
- 경쟁사/대체재와 비교 포인트(근거)

[필수]
- 수치/연도/근거(가능하면 출처) 중심으로 정리
- 최소 400자 이상 출력
- 확인 불가 항목은 '확인 불가'로 명시(추정 금지)

[정제용 원자료]
{clipped}
""".strip()

    payload = {
        "contents": [{"parts": [{"text": stage2_prompt}]}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 2048},
        "tools": [{"google_search": {}}]
    }

    r = _post_gemini(payload)
    if not r.get("ok"):
        return {"error": f"API Error {r.get('status')}", "message": r.get("text")}

    enriched = _extract_text(r["json"])
    if not enriched:
        return {"error": "EmptyText", "message": "2단계(RAG) 응답 텍스트가 비어있습니다."}

    return {"ok": True, "enriched_text": enriched}


# ==========================
# 3단계: JSON 파일 생성(네가 준 JSON mode 프롬프트 유지)
# ==========================
def stage3_generate_json(enriched_text: str, json_schema: str, max_chars: int = 25000) -> dict:
    clipped = (enriched_text or "").strip()[:max_chars]

    prompt = f"""
반드시 유효한 JSON만 출력하십시오. JSON 외 텍스트 금지.
Financial_Table은 원자료에 등장하는 모든 연도 컬럼을 빠짐없이 포함하십시오.

[Output JSON Schema]
{json_schema}

IR 텍스트:
{clipped}
""".strip()

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1, "responseMimeType": "application/json"}
    }

    r = _post_gemini(payload)
    if not r.get("ok"):
        return {"error": f"API Error {r.get('status')}", "message": r.get("text")}

    out_text = _extract_text(r["json"])
    if not out_text:
        return {"error": "EmptyText", "message": "3단계(JSON) 응답 텍스트가 비어있습니다."}

    try:
        return {"ok": True, "json": json.loads(out_text)}
    except JSONDecodeError:
        return {"error": "JSONDecodeError", "message": "3단계 JSON 파싱 실패", "preview": out_text[:800]}


# ==========================
# 4단계: 검토보고서 작성(JSON 기반)
# ==========================
def stage4_write_report(report_json: dict, max_output_tokens: int = 2048) -> dict:
    j = json.dumps(report_json, ensure_ascii=False)

    prompt = f"""
당신은 15년 경력의 VC 수석 심사역입니다.
아래 JSON(구조화된 투자검토 데이터)을 근거로 기관 투자자용 '검토보고서'를 작성하세요.

[규칙]
- 과장/추정 금지: JSON에 없는 수치/사실은 만들지 말 것
- 재무/시장/기술은 '근거 기반'으로 서술(가능하면 출처/근거를 문장에 포함)
- 아래 목차를 유지:
  1) 투자결론(요약)
  2) 기업/팀 개요
  3) 문제-해결/제품
  4) 기술/차별성(Moat)
  5) 시장/경쟁
  6) 사업/재무(연도별)
  7) 마일스톤/향후 계획
  8) 리스크 & 대응
  9) DD 질문
  10) 종합 의견

[입력 JSON]
{j}

마크다운으로 출력하세요.
""".strip()

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.2, "maxOutputTokens": max_output_tokens}
    }

    r = _post_gemini(payload)
    if not r.get("ok"):
        return {"error": f"API Error {r.get('status')}", "message": r.get("text")}

    report_md = _extract_text(r["json"])
    if not report_md:
        return {"error": "EmptyText", "message": "4단계(보고서) 응답 텍스트가 비어있습니다."}

    return {"ok": True, "report_md": report_md}


# ==========================
# Orchestrator: 1~4단계 실행
# ==========================
def run_pipeline_from_pdf(pdf_path: str, json_schema: str, use_rag: bool = True) -> dict:
    """
    1) (강화) PDF + Dynamic Retrieval로 1차 JSON 생성
    2) (옵션) 부족한 정보만 RAG로 보강(텍스트)
    3) 최종 JSON 재생성(JSON mode)
    4) 보고서 작성(옵션)
    """
    # 1단계: 바로 JSON
    s1 = stage1_extract_json_from_pdf(pdf_path, json_schema=json_schema)
    if not s1.get("ok"):
        return {"ok": False, "stage": 1, **s1}

    stage1_json = s1["json"]

    # use_rag=False면 1단계 결과를 최종으로
    if not use_rag:
        return {"ok": True, "report_json": stage1_json}

    # 2단계: 부족 항목만 보강할 텍스트 생성(검색)
    # (여기서는 stage1_json을 텍스트로 넘겨 "빈칸 채우기" 방식이 더 안정적)
    material_text = json.dumps(stage1_json, ensure_ascii=False)

    s2 = stage2_enrich_with_web_rag(material_text)
    if not s2.get("ok"):
        # 2단계 실패해도 1단계 JSON으로 진행할지 정책 선택 가능
        return {"ok": True, "report_json": stage1_json, "warn": "Stage2Failed", "stage2": s2}

    enriched_text = s2["enriched_text"]

    # 3단계: 보강 텍스트를 반영해 최종 JSON 다시 생성
    s3 = stage3_generate_json(enriched_text, json_schema=json_schema)
    if not s3.get("ok"):
        return {"ok": True, "report_json": stage1_json, "warn": "Stage3Failed", "stage3": s3}

    final_json = s3["json"]
    return {"ok": True, "report_json": final_json}