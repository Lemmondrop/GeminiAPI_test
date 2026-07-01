"""증권신고서 본문 → 비교가치 산정 라벨 3종 추출.

실제 신고서(최초 증권신고서) 구조에 맞춤 — 발행조건확정본엔 표가 없으므로
호출측이 corp_code로 최초 신고서를 먼저 찾아야 함(dart_client.find_initial_registration).

추출 3-타깃:
  1) comparison_ksic — 1차 '업종유사성' 선정기준에 명시된 KSIC 코드들 (Lane C prior 직접 라벨)
  2) final_peers    — 최종 선정 유사회사명 + 적용 PER (ground-truth 검증 + PER 보정)
  3) excluded       — 검토 후 제외된 회사 + 사유 (필터 기준 보정; 예 PER 137 아웃라이어)

본문은 426K자에 달하고 선정기준표(@앞부분)와 최종표(@평가섹션)가 흩어져 있어,
앵커 주변 관련 구간만 모아(collect_valuation_text) Gemini에 투입(토큰 절약 + 정확도).
LLM은 'NL 해석' 역할 — 표 포맷 변이가 커 정규식보다 견고. 회사명 정제는 규약 포팅.
"""
from __future__ import annotations

import json
import re

from config import GEMINI_LLM_MODEL

# ----------------------------- 평가 섹션 구간 수집 -----------------------------
# 선정기준표(KSIC) + 최종 선정표를 가리키는 앵커. 위험고지의 단독 '유사회사'는 제외(오탐).
VALUATION_ANCHORS = (
    "한국표준산업분류", "업종유사성", "유사회사 선정 기준", "비교기업 선정",
    "유사회사 PER 산출", "유사기업 PER", "평균 PER",
    "최종 유사기업", "최종 유사회사", "유사회사 선정 결과", "유사회사 선정결과",
)


def collect_valuation_text(full_text: str, window: int = 1700,
                           max_regions: int = 6, cap: int = 13000) -> str:
    """앵커 주변 구간들을 모아 병합(중복 제거). 평가 관련 텍스트만 압축 추출."""
    text = full_text or ""
    hits: list[int] = []
    for kw in VALUATION_ANCHORS:
        start = 0
        while True:
            i = text.find(kw, start)
            if i < 0:
                break
            hits.append(i)
            start = i + 1
    if not hits:
        return ""
    hits = sorted(set(hits))
    spans: list[list[int]] = []
    for i in hits:
        lo, hi = max(0, i - window // 3), i + window
        if spans and lo <= spans[-1][1]:             # 겹치면 병합
            spans[-1][1] = max(spans[-1][1], hi)
        else:
            spans.append([lo, hi])
    spans = spans[:max_regions]
    return "\n...\n".join(text[lo:hi] for lo, hi in spans)[:cap]


# ----------------------------- 회사명 정제·검증 (discount_pipeline 포팅) -----------------------------
SUMMARY_NAMES = {"평균", "합계", "소계", "계", "구분", "회사명", "상장일", "선정회사",
                 "유사회사", "비교회사", "비교기업", "유사기업", "발행회사", "당사", "동사",
                 "선정", "미선정", "선정여부", "적정"}
_INVALID_TOKENS = (
    "시장지수", "공모주식", "공모가", "희망공모", "확정공모", "청약", "수익률",
    "변동률", "제3자배정", "유상증자", "전환권", "기준", "산정", "충족", "미충족",
    "다음과", "같이", "아래", "코스닥", "코스피", "제조업", "선정여부",
)


def clean_company_name(value: str) -> str:
    v = re.sub(r"\s+", " ", str(value or "")).strip()
    v = re.sub(r"^주\d+\)\s*", "", v)
    v = re.sub(r"^[|:ㆍ·\-\s]+|[|:ㆍ·\-\s]+$", "", v)
    v = re.sub(r"^(?:KOSPI|KOSDAQ|KONEX)\s+", "", v, flags=re.IGNORECASE)
    v = re.sub(r"\s+(?:KOSPI|KOSDAQ|KONEX)$", "", v, flags=re.IGNORECASE)
    v = re.sub(r"\([^)]*\)$", "", v).strip()         # 끝의 (46.12) 등
    return v.strip()


def is_valid_company_name(value: str) -> bool:
    v = (value or "").strip()
    if not v or v in SUMMARY_NAMES:
        return False
    if len(v) < 2 or len(v) > 40:
        return False
    if re.fullmatch(r"[\d,\s.%\-]+", v):
        return False
    if not re.search(r"[가-힣A-Za-z]", v):
        return False
    if any(tok in v for tok in _INVALID_TOKENS):
        return False
    if re.search(r"\d{1,3},\d{3}", v):
        return False
    return True


_KSIC_RE = re.compile(r"[A-U]\d{4,5}")               # C26429 등


def _clean_ksic(v: str) -> str | None:
    m = _KSIC_RE.search(str(v or "").upper().replace(" ", ""))
    return m.group(0) if m else None


# ----------------------------- LLM 라벨 추출 -----------------------------
_PROMPT = """다음은 IPO 증권신고서의 '비교가치 산정 / 유사회사 선정' 관련 발췌다.
아래 항목을 JSON 객체로만 추출하라(설명·코드펜스 금지).

{{
  "issuer_ksic": "발행사 자신의 한국표준산업분류 코드(명시돼 있으면, 예 C26429; 없으면 null)",
  "comparison_ksic": ["1차 업종유사성 선정기준에 명시된 KSIC 코드들(예 C26429, C26410)"],
  "final_peers": [{{"name": "최종 선정된 유사회사명", "per": 최종 적용 PER 숫자 또는 null}}],
  "excluded": [{{"name": "검토했으나 최종 제외된 회사명", "reason": "제외 사유 요약"}}]
}}

규칙:
- final_peers는 '최종 유사회사/최종 유사기업 선정결과'에 포함된 회사만. 후보·탈락은 넣지 말 것.
- 회사명만(코드·수치·시장구분·'등 N개사' 제외). 발행사 본인 제외.
- comparison_ksic는 1차 선정기준의 코드만(최종 피어 코드 아님).
- 값 없으면 빈 배열 또는 null.

[발췌]
{regions}
"""


def _loads_obj(text: str) -> dict:
    t = (text or "").strip()
    if t.startswith("```"):
        t = t.strip("`")
        t = t[4:].strip() if t[:4].lower() == "json" else t
    try:
        obj = json.loads(t)
    except Exception:
        m = re.search(r"\{.*\}", t, re.S)
        obj = json.loads(m.group(0)) if m else {}
    return obj if isinstance(obj, dict) else {}


def _parse_labels(text: str) -> dict:
    obj = _loads_obj(text)
    # final_peers
    peers, seen = [], set()
    for p in (obj.get("final_peers") or []):
        if isinstance(p, dict):
            name, per = clean_company_name(p.get("name", "")), p.get("per")
        else:
            name, per = clean_company_name(str(p)), None
        if is_valid_company_name(name) and name not in seen:
            seen.add(name)
            try:
                per = float(per) if per is not None else None
            except (ValueError, TypeError):
                per = None
            peers.append({"name": name, "per": per})
    # comparison_ksic
    codes, cseen = [], set()
    for c in (obj.get("comparison_ksic") or []):
        code = _clean_ksic(c)
        if code and code not in cseen:
            cseen.add(code)
            codes.append(code)
    # excluded
    excl = []
    for e in (obj.get("excluded") or []):
        if isinstance(e, dict):
            name = clean_company_name(e.get("name", ""))
            if is_valid_company_name(name):
                excl.append({"name": name, "reason": str(e.get("reason", "")).strip()})
    return {
        "issuer_ksic": _clean_ksic(obj.get("issuer_ksic")),
        "comparison_ksic": codes,
        "final_peers": peers,
        "excluded": excl,
    }


def extract_filing_labels(text: str, client=None, model: str = GEMINI_LLM_MODEL) -> dict:
    """신고서 본문 → {issuer_ksic, comparison_ksic, final_peers[{name,per}], excluded}."""
    empty = {"issuer_ksic": None, "comparison_ksic": [], "final_peers": [], "excluded": []}
    regions = collect_valuation_text(text)
    if not regions.strip():
        return empty
    if client is None:
        from google import genai
        client = genai.Client()
    from google.genai import types
    resp = client.models.generate_content(
        model=model,
        contents=_PROMPT.format(regions=regions),
        config=types.GenerateContentConfig(
            temperature=0.0, response_mime_type="application/json"),
    )
    return _parse_labels(resp.text)


def extract_peers(text: str, client=None) -> list[str]:
    """최종 선정 유사회사명 리스트 (cooccur/배치용 단축 인터페이스)."""
    return [p["name"] for p in extract_filing_labels(text, client=client)["final_peers"]]


# ----------------------------- PDF 헬퍼(레거시 경로용) -----------------------------
def extract_pdf_text(pdf_path: str) -> str:
    import pdfplumber
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)