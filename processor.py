import os
import json
import time
import random
import re
import requests
from json import JSONDecodeError
from dotenv import load_dotenv

load_dotenv()
api_key = (os.getenv("GEMINI_API_KEY") or "").strip()

TARGET_MODEL = "models/gemini-2.0-flash"
API_BASE = "https://generativelanguage.googleapis.com/v1beta"
HEADERS = {"Content-Type": "application/json"}

# ---------------------------------
# Rate Limiter (요청 수 제한)
# ---------------------------------
class RateLimeter:
    def __init__(self, rpm: int = 4):
        self.min_interval = 60.0 / max(rpm, 1)
        self._last = 0.0

    def wait(self):
        now = time.time()
        elapsed = now - self._last
        if elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        self._last = time.time()

# 기본 호출(비-grounding) rpm
rate_limiter = RateLimeter(rpm=4)

# grounding은 더 빡빡하므로 별도 limiter (회사 1개당 1회라도 안전하게)
grounding_limiter = RateLimeter(rpm=1)

# -----------------------------
# 1. Low-level helpers
# -----------------------------
def _post_gemini(payload: dict, timeout: int = 120, max_retries: int = 5, base_wait: int = 10, use_grounding: bool = False) -> dict:
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY가 설정되지 않았습니다.")

    url = f"{API_BASE}/{TARGET_MODEL}:generateContent?key={api_key}"

    for attempt in range(max_retries):
        try:
            # 호출 유형별 rate limit
            if use_grounding:
                grounding_limiter.wait()
            else:
                rate_limiter.wait()

            resp = requests.post(url, headers=HEADERS, json=payload, timeout=timeout)

        except requests.RequestException:
            wait = 5 + random.uniform(0, 2)
            time.sleep(wait)
            continue

        if resp.status_code == 200:
            return {"ok": True, "json": resp.json()}

        # 429 Rate Limit
        if resp.status_code == 429:
            retry_after = resp.headers.get("Retry-After")
            wait = None

            if retry_after and str(retry_after).isdigit():
                wait = int(retry_after)

            if wait is None:
                try:
                    body = resp.json()
                    details = (((body.get("error") or {}).get("details")) or [])
                    for d in details:
                        if isinstance(d, dict) and str(d.get("@type", "")).endswith("RetryInfo"):
                            rd = str(d.get("retryDelay", "")).strip()
                            m = re.match(r"(\d+)\s*s", rd)
                            if m:
                                wait = int(m.group(1))
                                break
                except:
                    pass

            if wait is None:
                wait = base_wait * (2 ** attempt)

            wait = float(wait) + random.uniform(0, 2)
            print(f"    [429 Limit] {wait:.1f}초 대기 후 재시도 ({attempt+1}/{max_retries})...")
            time.sleep(wait)
            continue

        # 5xx
        if resp.status_code >= 500:
            print(f"    [Server Error {resp.status_code}] 재시도 중...")
            time.sleep(5)
            continue

        return {"ok": False, "status": resp.status_code, "text": resp.text}

    return {"ok": False, "status": 0, "text": "Max retries exceeded"}


def _extract_text(res_json: dict) -> str:
    candidates = res_json.get("candidates", [])
    if not candidates:
        return ""
    content = candidates[0].get("content", {}) or {}
    parts = content.get("parts", []) or []
    texts = []
    for p in parts:
        if isinstance(p, dict) and isinstance(p.get("text"), str):
            t = p["text"].strip()
            if t:
                texts.append(t)
    return "\n".join(texts).strip()


def _safe_json_loads(maybe_json_text: str):
    if not maybe_json_text:
        raise JSONDecodeError("empty", "", 0)

    text = maybe_json_text.strip()
    try:
        return json.loads(text)
    except JSONDecodeError:
        pass

    text2 = re.sub(r"```(?:json)?", "", text, flags=re.IGNORECASE).replace("```", "").strip()
    s = text2.find("{")
    e = text2.rfind("}")
    if s != -1 and e != -1 and e > s:
        try:
            return json.loads(text2[s:e+1])
        except:
            pass

    return json.loads(text)


def _ensure_dir(path: str):
    if not os.path.exists(path):
        os.makedirs(path)


def _sanitize_filename(name: str) -> str:
    # Windows 파일명 안전화
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:120] if len(name) > 120 else name


def _cache_path_for_company(company_name: str) -> str:
    _ensure_dir("evidence_cache")
    safe = _sanitize_filename(company_name) if company_name else "UNKNOWN"
    return os.path.join("evidence_cache", f"{safe}.json")


def _load_company_cache(company_name: str) -> dict | None:
    path = _cache_path_for_company(company_name)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return None


def _save_company_cache(company_name: str, evidence: dict):
    path = _cache_path_for_company(company_name)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(evidence, f, ensure_ascii=False, indent=2)
    except:
        pass


# -----------------------------
# 2. Schema Definition
# -----------------------------
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
          ["항목", "YYYY(A)", "YYYY+1(E)"],
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


# -----------------------------
# 3. Stage Logic
# -----------------------------
def _detect_missing_slots(draft: dict) -> dict:
    missing = []

    header = draft.get("Report_Header", {}) or {}
    company = str(header.get("Company_Name", "")).strip()

    fin_table = (draft.get("Table_Data_Preview", {}) or {}).get("Financial_Table", [])
    if not fin_table or not isinstance(fin_table, list) or len(fin_table) < 1 or not isinstance(fin_table[0], list) or len(fin_table[0]) < 4:
        missing.append("FINANCIAL_YEARS")

    market = draft.get("Market_Opportunity", {}) or {}
    chart_data = market.get("Market_Chart_Data", [])
    if len(chart_data) < 2:
        missing.append("MARKET_SIZE")
    else:
        # 값이 모두 0이면 결측으로 간주
        try:
            vals = []
            for row in chart_data[1:]:
                if isinstance(row, list) and len(row) > 1:
                    vals.append(str(row[1]).strip())
            if vals and all(v in ("0", "0.0", "", "None") for v in vals):
                missing.append("MARKET_SIZE")
        except:
            missing.append("MARKET_SIZE")

    tech = draft.get("Technology_and_Moat", {}) or {}
    if not tech.get("Technical_Details"):
        missing.append("TECH_DETAILS")

    # 수출/해외계약은 기본적으로 보강 시도
    missing.append("EXPORT_NEWS")

    needs_web = bool(company) and len(set(missing)) > 0
    return {"needs_web": needs_web, "missing": sorted(list(set(missing)))}


def _build_compact_queries(company_name: str, industry_hint: str, missing_slots: list[str]) -> list[str]:
    """
    결측 슬롯이 많아도 쿼리는 최대 2개로 압축:
    - Query 1: 회사 재무/공시/매출
    - Query 2: 회사 해외/수출/계약 + (산업 시장 규모)
    """
    c = company_name.strip()
    ind = (industry_hint or "").strip()

    q1 = f"{c} 매출 영업이익 재무제표 감사보고서"
    # 시장규모/수출/기술을 한 번에 묶기
    # (free tier grounding에서 query 수 늘리면 곧바로 불안정해짐)
    if ind:
        q2 = f"{c} 수출 해외 계약 MOU 수주 {ind} 시장 규모 TAM SAM SOM"
    else:
        q2 = f"{c} 수출 해외 계약 MOU 수주 시장 규모 TAM SAM SOM"

    # missing에 따라 1개만 써도 될 때는 1개로 줄임
    slots = set(missing_slots or [])
    if slots == {"FINANCIAL_YEARS"}:
        return [q1]
    if slots and slots.issubset({"EXPORT_NEWS", "MARKET_SIZE", "TECH_DETAILS"}):
        return [q2]

    return [q1, q2]


def _stage3_web_rag(company_name: str, industry_hint: str, missing_info: dict, draft_json: dict) -> dict:
    # 0) 캐시 우선
    cached = _load_company_cache(company_name)
    if isinstance(cached, dict) and cached.get("Findings"):
        print("   - [Stage 3] 캐시 Evidence 재사용")
        return cached

    missing_slots = missing_info.get("missing", [])
    queries = _build_compact_queries(company_name, industry_hint, missing_slots)

    evidence_schema = r"""
{
  "Findings": [
    {
      "Slot": "MARKET_SIZE | FINANCIAL_YEARS | EXPORT_NEWS | TECH_DETAILS",
      "Summary": "검색 결과 요약",
      "Key_Facts": ["팩트1", "팩트2"],
      "Sources": [{"Title": "제목", "URL": "url"}]
    }
  ],
  "Gaps": ["확인 불가 항목"]
}
""".strip()

    prompt = f"""
당신은 리서처입니다. 아래 쿼리를 사용해 웹 검색을 수행하고, Evidence JSON을 작성하십시오.

[회사명]
{company_name}

[결측 슬롯]
{json.dumps(missing_slots, ensure_ascii=False)}

[검색 쿼리(최대 2개)]
{json.dumps(queries, ensure_ascii=False)}

[작성 규칙]
- 반드시 신뢰 가능한 출처 중심으로 요약
- 숫자/연도/계약/수출 등 팩트는 Key_Facts에 구체적으로
- Sources는 실제 URL 포함
- Markdown 코드 블록 없이 순수 JSON만 출력

[Evidence JSON Schema]
{evidence_schema}
""".strip()

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "tools": [{"google_search": {}}],
        "generationConfig": {"temperature": 0.1}
    }

    # grounding은 더 불안정하므로 호출 전 추가 안전 대기(선택)
    # time.sleep(3)

    r = _post_gemini(payload, use_grounding=True)
    if not r.get("ok"):
        return {"error": f"Web RAG Fail {r.get('status')}", "message": r.get("text")}

    try:
        evidence = _safe_json_loads(_extract_text(r["json"]))
    except:
        evidence = {"Findings": [], "Gaps": missing_slots, "Note": "파싱 실패"}

    # 1) 캐시 저장 (Findings가 없어도 저장해두면 동일 회사에서 무한 재시도를 막음)
    _save_company_cache(company_name, evidence)
    return evidence


def _stage1_ir_to_draft(raw_text: str) -> dict:
    prompt = f"""
당신은 VC 심사역입니다. IR 자료를 분석하여 투자보고서 JSON 초안을 작성하십시오.
외부 지식 금지. 재무 테이블은 모든 연도 포함.

[Output JSON Schema]
{JSON_SCHEMA}

[IR 텍스트]
{raw_text}
""".strip()

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1, "responseMimeType": "application/json"}
    }

    r = _post_gemini(payload, use_grounding=False)
    if not r.get("ok"):
        return {"error": "Stage1 Fail", "msg": r.get("text")}

    return _safe_json_loads(_extract_text(r["json"]))


def _stage4_finalize_json(draft: dict, evidence: dict) -> dict:
    prompt = f"""
Draft와 Evidence를 통합하여 최종 보고서 JSON을 완성하십시오.
Evidence의 Key_Facts로 Draft의 빈칸을 채우되, 과장하지 마십시오.

[Draft]
{json.dumps(draft, ensure_ascii=False)}

[Evidence]
{json.dumps(evidence, ensure_ascii=False)}

[Output JSON Schema]
{JSON_SCHEMA}
""".strip()

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1, "responseMimeType": "application/json"}
    }

    r = _post_gemini(payload, use_grounding=False)
    if not r.get("ok"):
        return draft

    return _safe_json_loads(_extract_text(r["json"]))


# -----------------------------
# 4. Public Function
# -----------------------------
def refine_to_json(raw_text: str, max_chars: int = 25000) -> dict:
    clipped = (raw_text or "").strip()[:max_chars]

    print("   1) [Stage 1] IR 초안 생성...")
    draft = _stage1_ir_to_draft(clipped)
    if isinstance(draft, dict) and "error" in draft:
        return draft

    missing = _detect_missing_slots(draft)

    evidence = {}
    if missing.get("needs_web"):
        print(f"   2) [Stage 2] 결측 감지: {missing.get('missing')}")
        print("   3) [Stage 3] 웹 검색(Grounding) 1회 + 캐시 사용...")
        header = draft.get("Report_Header", {}) or {}
        evidence = _stage3_web_rag(
            str(header.get("Company_Name", "")).strip(),
            str(header.get("Industry_Sector", "")).strip(),
            missing,
            draft
        )
    else:
        print("   INFO: 추가 검색 불필요")

    # evidence가 없거나 실패하면 Stage4 생략(호출 수 감소)
    if isinstance(evidence, dict) and evidence.get("Findings"):
        print("   4) [Stage 4] 최종 병합...")
        return _stage4_finalize_json(draft, evidence)

    return draft
