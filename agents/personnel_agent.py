import json
import re
from pydantic import BaseModel, Field
from typing import List, Optional
from utils import call_gemini, safe_json_loads

# =========================================================================
# 1. [Pydantic 스키마 정의] - 기존 문자열 스키마 3종을 완벽하게 대체
# =========================================================================

# [1-1] Company Signature Schema
class CompanySignatureSchema(BaseModel):
    company_name_kr: str = Field(description="법인명(국문, 문서에서 확인된 그대로. 없으면 빈 문자열)")
    company_name_en: str = Field(description="영문명(있으면)")
    company_domain: str = Field(description="공식 도메인(있으면, 예: example.com)")
    company_address: str = Field(description="본점/주소(있으면)")
    biz_reg_no: str = Field(description="사업자등록번호(있으면, 숫자/하이픈 무관)")
    ticker_or_market: str = Field(description="상장시장/종목코드(있으면)")

# [1-2] CEO Evidence Schema (RAG 검증용)
class EvidenceItem(BaseModel):
    source_type: str = Field(description="공시|등기|홈페이지|언론|기타")
    source_name: str = Field(description="출처명")
    published_date: str = Field(description="YYYY-MM-DD 또는 YYYY 또는 unknown")
    snippet: str = Field(description="대표이사/CEO임을 확인 가능한 짧은 문장/구절")
    url: str = Field(description="가능하면 URL, 없으면 빈 문자열")

class CeoClaim(BaseModel):
    name: str = Field(description="대표이사/CEO 성명")
    title: str = Field(description="직함(대표이사/CEO/공동대표 등)")
    current_or_past: str = Field(description="current|past|unknown")
    as_of_date: str = Field(description="근거 기준일(YYYY-MM-DD 또는 YYYY 또는 unknown)")
    evidence: List[EvidenceItem] = Field(description="대표이사 확인 근거 배열")

class CeoEvidenceSchema(BaseModel):
    company_signature: CompanySignatureSchema
    ceo_claims: List[CeoClaim]

# [1-3] 최종 Personnel Response Schema
class CeoReference(BaseModel):
    Name: str = Field(description="성명")
    Background_and_Education: str = Field(description="학력 및 주요 경력 (연도별 상세)")
    Core_Competency: str = Field(description="핵심 역량 (기술/경영/네트워크)")
    Management_Philosophy: str = Field(description="경영 철학 및 비전")
    VC_Perspective_Evaluation: str = Field(description="VC 관점에서의 정성적 평가 (리더십, 성공 가능성). 🚨근거 부족 시 '검증 실패' 같은 시스템 에러 문구 절대 금지. 정중한 애널리스트 톤으로 작성할 것.")

class TeamCapability(BaseModel):
    Key_Executives: List[str] = Field(description="주요 임원 상세 이력 배열. 데이터가 없으면 빈 배열 []")
    Organization_Strengths: str = Field(description="팀워크, 연구 인력 비중, 조직 문화")
    Advisory_Board: str = Field(description="자문위원단 및 외부 네트워크")

class CapTableItem(BaseModel):
    Shareholder: str = Field(description="주주명 (예: 창업자, VC, 엔젤 등)")
    Shares: str = Field(description="보유 주식수 (있을 경우, 없으면 '-')")
    Ratio: str = Field(description="지분율 (%)")

class KeyPersonnel(BaseModel):
    CEO_Reference: CeoReference
    Team_Capability: TeamCapability
    Cap_Table: List[CapTableItem] = Field(description="주주명부 배열. 없으면 빈 배열 []")

class PersonnelResponseSchema(BaseModel):
    Key_Personnel: KeyPersonnel

# =========================================================================
# 2. [검증 및 로직 함수] - 변수명 및 로직 100% 보존
# =========================================================================
def _norm(s: str) -> str:
    if not s: return ""
    x = str(s).strip().lower().replace("주식회사", "").replace("(주)", "").replace("㈜", "")
    x = re.sub(r"\s+", "", x)
    x = re.sub(r"[^0-9a-z가-힣\.\-:/]", "", x)
    return x

def _extract_company_signature(pdf_path: str, extra_text: str = "") -> dict:
    prompt = f"""
너는 IR 문서에서 회사의 '식별자'만 추출하는 도구다. 추측 금지.
문서에 명시된 값만 추출하여라. 없으면 빈 문자열로 둔다.

[보충 문서 데이터]
{extra_text}
"""
    # 🚨 [최적화] response_schema 주입
    res = call_gemini(prompt, pdf_path=pdf_path, response_schema=CompanySignatureSchema)
    if not res.get("ok"):
        return safe_json_loads('{"company_name_kr":"","company_name_en":"","company_domain":"","company_address":"","biz_reg_no":"","ticker_or_market":""}')
    sig = safe_json_loads(res.get("text", "")) or {}
    for k in ["company_name_kr","company_name_en","company_domain","company_address","biz_reg_no","ticker_or_market"]:
        sig[k] = (sig.get(k) or "").strip()
    return sig

def _build_queries(ceo_name: str, sig: dict) -> list[str]:
    ceo_name = (ceo_name or "").strip()
    name_kr = (sig.get("company_name_kr") or "").strip()
    domain = (sig.get("company_domain") or "").strip()

    queries = []
    if name_kr:
        queries += [f"{name_kr} 대표이사 {ceo_name}", f"{name_kr} 임원현황 {ceo_name}"]
    else:
        queries += [f"{ceo_name} 대표이사 프로필"]
    if domain:
        queries += [f"site:{domain} {ceo_name} 대표이사"]
    return queries

def _extract_ceo_evidence(ceo_name: str, sig: dict) -> dict:
    queries = _build_queries(ceo_name, sig)
    rag_prompt = (
        "아래 각 질의를 검색하여, '대표이사/CEO임을 확인할 수 있는 문장'과 함께 출처를 남기십시오.\n"
        + "\n".join([f"- {q}" for q in queries])
    )
    # RAG 검색 (자유 텍스트 반환이므로 스키마 적용 안함)
    rag_res = call_gemini(rag_prompt, tools=[{"google_search": {}}])
    rag_text = rag_res.get("text", "") if rag_res.get("ok") else ""

    struct_prompt = f"""
너는 VC의 인사 검증 리서처다.
아래 RAG 텍스트에서 '대표이사/CEO' 주장만 추출하여 근거(evidence)로 묶어라.

[company_signature]
{json.dumps(sig, ensure_ascii=False)}

[RAG TEXT]
{rag_text}
"""
    # 🚨 [최적화] 추출된 텍스트를 구조화할 때 response_schema 주입
    sres = call_gemini(struct_prompt, response_schema=CeoEvidenceSchema)
    if not sres.get("ok"): return {"company_signature": sig, "ceo_claims": []}
    out = safe_json_loads(sres.get("text", "")) or {"company_signature": sig, "ceo_claims": []}
    if "company_signature" not in out: out["company_signature"] = sig
    return out

def _company_match_score(sig: dict, evidence_item: dict) -> int:
    blob = " ".join([str(evidence_item.get("snippet","")), str(evidence_item.get("source_name","")), str(evidence_item.get("url",""))])
    b = _norm(blob)
    score = 0
    name_kr = _norm(sig.get("company_name_kr",""))
    domain = _norm(sig.get("company_domain",""))
    biz = re.sub(r"[^0-9]", "", sig.get("biz_reg_no","") or "")
    if name_kr and name_kr in b: score += 3
    if domain and domain in b: score += 3
    if biz and biz in re.sub(r"[^0-9]", "", b): score += 4
    return score

def _validate_ceo_evidence(evd: dict, ceo_name: str, sig: dict) -> tuple[bool, str, dict]:
    ceo_name = (ceo_name or "").strip()
    claims = evd.get("ceo_claims", []) or []

    def fail(reason: str):
        return False, reason, {"company_signature": sig, "ceo_claims": [], "validation_failed": True, "reason": reason}

    if not (sig.get("company_name_kr") or "").strip(): return fail("회사명 미확인")
    if not claims: return fail("CEO 주장 근거 없음")
    
    name_hits = [c for c in claims if (c.get("name") or "").strip() == ceo_name]
    if not name_hits: return fail("동명이인/타인 정보 가능성")

    best = None
    for c in name_hits:
        evidence = c.get("evidence") or []
        if not evidence: continue
        max_match = max((_company_match_score(sig, e) for e in evidence), default=0)
        official = sum(1 for e in evidence if (e.get("source_type") or "").strip() in ("공시","등기","홈페이지") and _company_match_score(sig, e) > 0)
        cand = (max_match, official, len(evidence), c)
        if (best is None) or (cand[:3] > best[:3]): best = cand

    if best is None: return fail("근거가 없음")
    max_match, official, ev_cnt, best_claim = best

    if max_match < 3: return fail("회사 동일성 검증 실패")
    if official < 1: return fail("공식/준공식 근거 부재")

    warning = "근거 1개로 제한됨" if ev_cnt < 2 else ""
    return True, warning, {"company_signature": sig, "ceo_claims": [best_claim], "validation_warning": warning}

# =========================================================================
# 3. [분석 엔진(Agent) 메인 실행부]
# =========================================================================
def analyze(pdf_path: str, ceo_name: str, extra_text: str = "") -> dict:
    print(f"   [Personnel Agent] 경영진 및 조직 역량 분석 중...")

    sig = _extract_company_signature(pdf_path, extra_text)

    if ceo_name:
        # 사실 확인을 위해 RAG 및 자체 검증 수행
        evd = _extract_ceo_evidence(ceo_name, sig)
        ok, warn, cleaned = _validate_ceo_evidence(evd, ceo_name, sig)
        verified_ceo_context = json.dumps(cleaned, ensure_ascii=False, indent=2)
    else:
        verified_ceo_context = json.dumps(
            {"company_signature": sig, "ceo_claims": [], "validation_failed": True, "reason": "ceo_name 미제공"},
            ensure_ascii=False, indent=2
        )

    # 🚨 [최적화] 프롬프트 내 하드코딩된 JSON 양식 제거 및 톤앤매너 지시사항 강조
    prompt = f"""
당신은 벤처캐피탈(VC)의 인사 검증 담당자이자 전문 심사역(Analyst)입니다.
메인 자료(PDF), [보충 문서 데이터], 그리고 '검증된 CEO 근거'를 사용하여 경영진과 조직 역량을 평가하십시오.

[보충 문서 데이터]
{extra_text}

[Verified CEO Evidence (JSON)]
{verified_ceo_context}

[최우선 원칙]
1. Verified CEO Evidence에 없는 개인 정보는 생성 금지.
2. PDF나 보충 문서에 팀 역량 근거가 없으면 N/A 처리.
3. [주주명부(Cap Table) 추출 엄수 - Data Shift 방지 및 가비지 데이터 필터링]: 
   - 텍스트로 뭉개진 표 데이터에서 주주명, 주식수, 지분율을 추출할 때 값이 다른 행(Row)으로 밀리거나 뒤섞이지 않도록 논리적으로 매핑하십시오.
   - '합계(Total)' 또는 '계' 항목은 개별 주주가 아니므로 배열에서 무조건 제외하십시오.
   - [악성 데이터 차단]: 지분율 항목 자리에 '자본금'이라는 단어가 있거나, 주식수 자리에 '백만원', '원' 등 금액 단위가 적혀 있는 행(Row)은 주주 정보가 아니라 표에 잘못 섞여 들어온 '자본 요약 정보'입니다. 이러한 행은 절대 추출하지 말고 무조건 버리십시오(Drop).
   - 검증된 실제 주주 데이터만 숫자 그대로(콤마, % 기호 포함) 정확히 전사하십시오.

[예외 상황 작성 톤앤매너 (매우 중요)]
Verified CEO Evidence가 'validation_failed' 상태이거나 제공된 자료 내에 대표이사/임원진의 구체적 정보가 부족할 경우, "검증 실패", "평가 불가", "정보 없음" 등 기계적이고 단답형의 시스템 에러 같은 표현을 절대 사용하지 마십시오.
대신, 반드시 아래와 같이 전문적인 애널리스트의 어조(Tone)로 정중하게 서술하십시오.
- (잘못된 예) "Verified CEO Evidence 부재로 평가 불가"
- (잘못된 예) "공식/준공식 근거 부재로 검증 실패"
- (올바른 예) "제공된 IR 자료 내에서 대표이사의 핵심 역량 및 과거 레퍼런스를 교차 검증할 수 있는 구체적인 정보가 확인되지 않아, 현 시점에서의 경영 능력 세부 평가는 제한적입니다. 향후 추가적인 인터뷰나 보완 자료를 통한 확인이 필요합니다."
"""
    # [최적화] 최종 PersonnelResponseSchema 주입
    res = call_gemini(prompt, pdf_path=pdf_path, response_schema=PersonnelResponseSchema)
    if res.get("ok"):
        return safe_json_loads(res["text"])
    else:
        print(f"   [Personnel Agent] Error: {res.get('error')}")
        return {}