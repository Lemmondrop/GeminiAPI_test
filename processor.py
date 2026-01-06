import os
import json
import time
import random
import requests
from json import JSONDecodeError
from dotenv import load_dotenv

load_dotenv()
api_key = (os.getenv("GEMINI_API_KEY") or "").strip()

TARGET_MODEL = "models/gemini-2.5-flash"
API_BASE = "https://generativelanguage.googleapis.com/v1beta"
HEADERS = {"Content-Type": "application/json"}


def _post_gemini(payload: dict, timeout: int = 120, max_retries: int = 5, base_wait: int = 10) -> dict:
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY가 설정되지 않았습니다.")

    url = f"{API_BASE}/{TARGET_MODEL}:generateContent?key={api_key}"

    for attempt in range(max_retries):
        try:
            resp = requests.post(url, headers=HEADERS, json=payload, timeout=timeout)
        except requests.RequestException as e:
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
    candidates = res_json.get("candidates", [])
    if not candidates:
        return ""
    parts = (candidates[0].get("content", {}) or {}).get("parts", []) or []
    texts = []
    for p in parts:
        if isinstance(p, dict) and isinstance(p.get("text"), str):
            t = p["text"].strip()
            if t:
                texts.append(t)
    return "\n".join(texts).strip()


def refine_to_json(raw_text: str, max_chars: int = 25000, use_search: bool = True) -> dict:
    """
    main.py에서 from processor import refine_to_json 로 가져다 쓰는 함수
    - use_search=True: 1차(검색) 원자료 → 2차(JSON mode) 정제
    - use_search=False: JSON mode로 바로 생성
    """

    clipped = (raw_text or "").strip()[:max_chars]

    json_schema = r"""
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

    # use_search=False: 단일 호출(JSON mode)
    if not use_search:
        prompt = f"""
반드시 유효한 JSON만 출력하십시오. JSON 외 텍스트 금지.
Financial_Table은 문서에 있는 모든 연도 컬럼을 빠짐없이 포함하십시오.

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
            return {"error": "EmptyText", "message": "응답 텍스트가 비어있습니다."}

        try:
            return json.loads(out_text)
        except JSONDecodeError:
            return {"error": "JSONDecodeError", "message": "JSON 파싱 실패", "preview": out_text[:800]}

    # use_search=True: 2단계 호출
    stage1_prompt = f"""
아래 IR 텍스트를 바탕으로 핵심 사실/수치/근거를 정리하십시오.
- 반드시 텍스트를 200자 이상 출력
- 재무제표에 등장하는 모든 연도 구간/수치 나열
- 시장규모/경쟁사 등은 필요 시 검색으로 보완

IR 텍스트:
{clipped}
""".strip()

    payload1 = {
        "contents": [{"parts": [{"text": stage1_prompt}]}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 2048},
        "tools": [{"google_search": {}}]
    }
    r1 = _post_gemini(payload1)
    if not r1.get("ok"):
        return {"error": f"API Error {r1.get('status')}", "message": r1.get("text")}

    stage1_text = _extract_text(r1["json"])
    if not stage1_text:
        return {"error": "EmptyText", "message": "1차 응답 텍스트가 비어있습니다."}

    stage2_prompt = f"""
반드시 유효한 JSON만 출력하십시오. JSON 외 텍스트 금지.
Financial_Table은 원자료에 등장하는 모든 연도 컬럼을 빠짐없이 포함하십시오.

[Output JSON Schema]
{json_schema}

[정제용 원자료]
{stage1_text}
""".strip()

    payload2 = {
        "contents": [{"parts": [{"text": stage2_prompt}]}],
        "generationConfig": {"temperature": 0.1, "responseMimeType": "application/json"}
    }
    r2 = _post_gemini(payload2)
    if not r2.get("ok"):
        return {"error": f"API Error {r2.get('status')}", "message": r2.get("text")}

    out_text = _extract_text(r2["json"])
    if not out_text:
        return {"error": "EmptyText", "message": "2차(JSON) 응답 텍스트가 비어있습니다."}

    try:
        return json.loads(out_text)
    except JSONDecodeError:
        return {"error": "JSONDecodeError", "message": "2차 JSON 파싱 실패", "preview": out_text[:800]}
