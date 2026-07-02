"""core/ir_parser.py — OCR 추출된 IR .md 텍스트에서 stage_router/valuation 파이프라인이
필요로 하는 구조화 필드(설립연도·투자라운드·IPO 언급·순이익·매출·업종 등)를
Gemini로 자동 추출한다.

배경:
  지금까지 ir_signals.json은 사람이 IR을 읽고 손으로 채웠다. "예상 밸류에이션"
  페이지에서 IR을 업로드하면 바로 결과가 나와야 한다는 원래 목표를 완성하려면,
  이 추출 단계 자체가 자동화되어야 한다. determine_company_stage()는 이미
  이 모듈의 출력(신호 dict)을 받는 소비자로 준비되어 있다.

재사용:
  core/llm_judge.py의 패턴(genai.Client() 지연 임포트, temperature=0,
  response_mime_type="application/json", 429/5xx 재시도, 코드펜스 가드)을
  그대로 따른다 — 같은 실패 모드를 새로 만들지 않기 위해.

캐싱:
  llm_judge.py는 (발행사, 후보) 쌍 단위로 캐싱하지만, 이 모듈은 문서
  단위(파일 내용 해시)로 캐싱한다 — 같은 IR을 다시 돌려도 재호출하지 않는다.
  IR 파일 내용이 바뀌면(개정판 업로드 등) 해시가 달라져 자동 재추출된다.

segment 필드 주의:
  discount_rate_pipeline의 input_limited 모델은 segment_{key} 형태의 원핫
  피처를 기대하므로, key가 아래 SEGMENTS 12개 중 정확히 하나가 아니면
  조용히 무시된다(모델 쪽은 에러 없이 그냥 0으로 처리). 그래서 자유 텍스트가
  아니라 객관식으로 강제하고, 응답이 목록 밖이면 "unknown"으로 떨어뜨린다.
"""
from __future__ import annotations

import hashlib
import json
import re
import time
from pathlib import Path

try:
    import config
    GEMINI_LLM_MODEL = config.GEMINI_LLM_MODEL
    LLM_MAX_RETRIES = config.LLM_MAX_RETRIES
    _CACHE_DIR = config.LLM_JUDGE_CACHE.parent if hasattr(config, "LLM_JUDGE_CACHE") else Path(".cache")
except ImportError:
    GEMINI_LLM_MODEL = "gemini-2.0-flash"
    LLM_MAX_RETRIES = 3
    _CACHE_DIR = Path(".cache")

IR_PARSER_CACHE = _CACHE_DIR / "ir_parser_cache.json"
# v1→v2: net_income_projections/revenue_projections(다년도) 필드 추가.
# v2→v3: business_text(사업설명, Stage3 피어매칭용) 필드 추가.
# v3→v4: business_text 환각(hallucination) 실사례 발견 후 방지 지시 강화
#   (2026-07 케이더블유씨 케이스 — 원문에 없는 "캡슐 커피용 필터"를 지어낸
#   사실 확인됨).
# v4→v5: 위 환각의 실제 원인이 "순수 지어냄"이 아니라 "Gemini 사전학습
#   지식(실제 회사 홈페이지 등)이 문서 내용과 섞여 나온 것"일 가능성이 높다는
#   사용자 지적 반영 — API 호출에 웹검색 tool은 없음(확인됨)이므로 실시간
#   검색이 아니라 사전학습 데이터 혼입으로 추정. "지어내지 마라"만으론 이걸
#   못 막는다(모델 입장에선 지어낸 게 아니라 "아는 사실"이므로) — "문서 밖
#   지식은 전부 무시하라"는 별도 지시 추가.
# v5→v6: 프롬프트 지시만으로는 v5도 뚫림이 실사례로 재확인됨(같은 케이더블유씨
#   케이스 — 이번엔 원문 미확인 고객사명(CJ/대상/광동제약/MALK/Nesquik)까지 새로
#   등장). 프롬프트로 막는 걸 포기하고, business_text_entities(고유명사 목록)를
#   별도로 뽑게 해서 PDF 원문 텍스트와 결정론적으로(LLM 아님) 대조하는
#   verify_entities()를 추가.
# v6→v7: v6 검증 결과 중 일부(Kwangdong/DAESANG)가 실제로는 진짜 사실(원문엔
#   한글 "광동"/"대상")인데 AI가 로마자로 바꿔 써서 검증기가 오탐한 것으로
#   실사례 확인(같은 케이더블유씨, 실제 IR 슬라이드 대조로 확인됨). 반면 같은
#   응답의 "CROWN"은 원문의 실제 파트너명("Shamrock Milk")과 다른 이름으로
#   나와 진짜 문제로 남음 — 표기 오탐과 진짜 오류가 섞여 있으면 사람이 매번
#   구분해야 해서 검증 결과 신뢰도가 떨어진다. 고유명사는 문서 표기 그대로
#   쓰라는 지시(번역·로마자 변환 금지)를 추가해 표기 불일치로 인한 오탐 자체를
#   줄인다(이건 "지어내지 마라"류가 아니라 "단위 환산"류의 기계적 지시라 더
#   안정적으로 지켜질 것으로 기대 — 단, 실측 재검증 필요).
# 캐시 키가 (프롬프트버전, 문서내용)이므로 버전을 올려야 예전 스키마/지시로
# 캐시된 결과를 그대로 재사용하지 않고 다시 추출한다.
IR_PARSER_PROMPT_VERSION = "v7"

# discount_rate_pipeline/model_input_limited.pkl이 기대하는 segment_* 12종 —
# 이 목록 밖 값은 전부 unknown으로 강등한다.
SEGMENTS = ("ai", "consumer", "deeptech", "fintech", "manufacturing", "materials",
            "mobility", "platform", "robotics", "semiconductor", "software", "unknown")

# stage_router.determine_ipo_subtrack()이 general/tech_special을 가르는 기준이
# net_income 부호이므로, latest_round는 자유 텍스트로 받아도 normalize_round()가
# 알아서 처리한다 — 여기서는 형식을 강제하지 않는다.

# 스키마 설명은 텍스트 경로/PDF 경로가 공유한다 — 따로 두면 나중에 한쪽만
# 고치고 잊어버려서 스키마가 갈라지는 사고가 난다(이 프로젝트에서 여러 번
# 겪은 "같은 걸 두 곳에서 각자 판단" 패턴과 동일).
SCHEMA_BLOCK = """{
  "company_name": "문서에 명시된 회사명 (원문 그대로)",
  "business_text": "3~5문장의 사업설명. 다른 기업과 비교(피어 매칭)하는 데 쓰이므로
    다음을 반드시 포함: (1) 한 문장으로 압축한 사업 정체성(예: '~ 전문기업', '~ 솔루션 기업'),
    (2) 구체적 제품/서비스·사업 영역, (3) 핵심 기술이나 차별점(보유 특허·기술·인증 등이
    언급되어 있으면 포함), (4) 규모를 보여주는 사실(설립연도, 주요 고객사, 매출 등 문서에
    실제로 있는 것만).
    ⚠ 절대 규칙 — 문서에 문자 그대로 나오지 않는 제품명·기술명·고객사명을 단 하나도
    지어내지 마라. '이런 사업이면 이런 제품도 있을 법하다'는 추론이 가장 위험한 실수다.
    구체적 명사(제품명, 기술명, 고객사명)는 특히 문서에 있는 표현을 최대한 그대로 옮겨
    적고, 요약/일반화하다가 있지도 않은 디테일을 새로 만들어내지 마라. 문서 내용을
    압축하는 것과 내용을 지어내는 것은 다르다 — 애매하면 그 문장을 통째로 빼라.
    ⚠ 표기 규칙 — 고유명사(회사명·브랜드명 등)는 문서에 적힌 표기를 그대로 써라.
    번역하거나 로마자로 바꾸지 마라. 예: 문서에 "광동"이라고 한글로만 적혀 있으면
    "광동"이라고 써야지 "Kwangdong"으로 바꿔 쓰면 안 된다. 문서에 "대상"이면 "대상"이지
    "DAESANG"이 아니다. 문서에 영문으로 적혀 있으면 영문 그대로. 알고 있는 정식 영문
    사명이 따로 있어도, 이 문서에 적힌 표기가 아니면 그쪽을 쓰지 마라.
    문서에 사업 설명 자체가 없으면 null.",
  "business_text_entities": business_text 안에 사용한 구체적 고유명사만 전부 리스트로
    뽑아라 — 제품명, 기술명, 고객사·브랜드명 등(예: ["PET/PVC/OPS/PP 수축라벨", "CJ",
    "Nesquik", "수성 점·접착 섬유 벽지 제조 기술"]). "포장재"·"친환경 소재" 같은 일반명사는
    제외. 여기도 위 표기 규칙(문서 표기 그대로, 번역·로마자 변환 금지)을 똑같이 지켜라.
    이 목록은 나중에 문서 원문과 자동 대조되니, business_text에 쓴 고유명사를
    빠짐없이 그대로 옮겨 적어라. business_text가 null이면 빈 리스트 [].
  "founding_year": 설립연도(정수) 또는 null,
  "latest_round": "가장 최근에 유치한 투자 라운드 명칭(예: Series A, Seed, Pre-IPO)을 문서 표현 그대로" 또는 null,
  "ipo_mentioned": IR에 IPO 계획·시점이 명시적으로 언급되어 있으면 true, 전혀 언급 없으면 false,
  "ipo_target_year": IPO 목표연도(정수, 언급된 경우만) 또는 null,
  "net_income": 가장 최근 결산(실적, 추정 아님) 기준 순이익(억원 단위 숫자, 적자면 음수) 또는 null,
  "revenue": 가장 최근 결산(실적, 추정 아님) 기준 매출(억원 단위 숫자) 또는 null,
  "net_income_projections": 문서에 "추정 재무제표"·"Financial Projection" 같은 표로 제시된
    미래 여러 연도의 추정 순이익을 전부 {"연도": 순이익억원, ...} 형태로. 표에 2026E/2027E/2028E
    처럼 여러 연도가 있으면 전부 넣어라(하나만 고르지 말 것). 예: {"2026": 1.2, "2027": 4.1, "2028": 8.5}.
    다년도 추정 표 자체가 문서에 없으면 null.
  "revenue_projections": 위와 같은 방식으로 미래 여러 연도의 추정 매출을 전부
    {"연도": 매출억원, ...} 형태로. 예: {"2026": 18.0, "2027": 32.5, "2028": 54.0}. 없으면 null.
  "future_net_income": [구버전 호환용] net_income_projections가 있으면 그 안의 가장 가까운
    연도 값과 동일하게 채워라. net_income_projections를 못 찾았는데 단일 연도짜리 추정
    순이익만 있으면 그 값을 넣어라. 둘 다 없으면 null.
  "ni_year": future_net_income에 대응하는 연도(정수) 또는 null.
  "segment": 아래 12개 중 가장 가까운 업종 하나만 정확히 그대로 선택(다른 단어 사용 금지):
    ai, consumer, deeptech, fintech, manufacturing, materials, mobility, platform,
    robotics, semiconductor, software, unknown
    — 애매하면 unknown을 선택할 것. 절대 목록에 없는 단어를 만들지 말 것.
  "extraction_notes": "추출 근거나 애매했던 부분에 대한 한국어 1~2문장 메모" (없으면 빈 문자열)
}"""

_COMMON_INSTRUCTION = """문서에 명시적으로 나타난 정보만 추출하라. 추측하지 말고, 문서에 없으면
반드시 null로 남겨라 — 없는 값을 지어내는 것이 가장 위험한 실수다. 표·재무제표 안의 숫자는
단위(원/백만원/억원)를 확인해서 반드시 억원 단위로 환산해 넣어라.

⚠ 매우 중요 — 이 회사에 대해 당신이 사전학습 과정에서 이미 알고 있을 수 있는 정보
(홈페이지, 뉴스, 공시, 다른 자료 등에서 접했을 법한 내용)는 전부 무시하라. 오직 지금
첨부/제공된 이 문서 안에 실제로 적힌 내용만 사용해야 한다. 당신이 이 회사에 대해
"알고 있는" 사실이라도, 그게 이 문서에 문자 그대로 나오지 않으면 절대 답에 포함시키지
마라 — 실제로 맞는 정보라도 이 문서 밖에서 가져온 것이면 여기서는 틀린 답이다. 이건
"지어내지 말라"는 지시와는 다른 별개의 규칙이니 둘 다 지켜라.

아래 JSON만 출력하라(설명·마크다운·코드펜스 금지). 스키마:
""" + SCHEMA_BLOCK

# 텍스트(.md 등) 경로용 — OCR 결과 텍스트를 프롬프트에 직접 삽입.
PROMPT_TEMPLATE = """당신은 비상장기업 IR 자료에서 기업가치 산정에 필요한 정형 필드를 추출하는 애널리스트다.
아래 IR 자료 텍스트를 읽고 정보를 추출하라.

[IR 자료 텍스트]
{ir_text}

""" + _COMMON_INSTRUCTION

# PDF 원본 첨부용 — OCR을 거치지 않고 Gemini가 PDF를 직접 읽는다.
# 표·재무제표 숫자는 OCR 오독(자릿수·소수점 오류)이 실제 위험이라, 가능하면
# 이 경로를 우선 쓴다. .md는 이 경로 실패 시의 폴백으로만 남겨둔다.
PDF_PROMPT_TEMPLATE = """당신은 비상장기업 IR 자료에서 기업가치 산정에 필요한 정형 필드를 추출하는 애널리스트다.
첨부된 IR 자료(PDF) 원본을 직접 읽고 정보를 추출하라. 특히 표·재무제표의 숫자는
스캔/OCR을 거치지 않고 원본에서 직접 읽어 정확도를 높여라.

""" + _COMMON_INSTRUCTION


def _file_hash(text: str) -> str:
    return hashlib.sha256(
        f"{IR_PARSER_PROMPT_VERSION}\x00text\x00{text}".encode("utf-8")
    ).hexdigest()[:24]


def _pdf_hash(pdf_bytes: bytes) -> str:
    return hashlib.sha256(
        f"{IR_PARSER_PROMPT_VERSION}\x00pdf\x00".encode("utf-8") + pdf_bytes
    ).hexdigest()[:24]


def _parse_retry_delay(exc: Exception):
    s = str(exc)
    if "RESOURCE_EXHAUSTED" in s or "429" in s:
        m = re.search(r"retry in ([\d.]+)s", s) or re.search(r"'retryDelay':\s*'(\d+)s'", s)
        return float(m.group(1)) if m else 30.0
    if any(t in s for t in ("UNAVAILABLE", "503", "500", "INTERNAL", "DEADLINE_EXCEEDED")):
        return 15.0
    return None


def _coerce_int(v):
    if v is None:
        return None
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def _coerce_float(v):
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _coerce_year_dict(v, min_year: int = 2000, max_year: int = 2100):
    """{"연도": 값, ...} 형태를 {int: float, ...}로 정규화.
    연도나 값이 숫자로 안 바뀌거나 상식적 범위(2000~2100) 밖이면 그 항목만
    조용히 버린다(llm_judge.parse_judgment()과 같은 가드레일 철학 —
    전체를 버리지 말고 이상한 항목만 걸러내기)."""
    if not isinstance(v, dict):
        return None
    out = {}
    for k, val in v.items():
        try:
            year = int(float(k))
            amount = float(val)
        except (TypeError, ValueError):
            continue
        if min_year <= year <= max_year:
            out[year] = amount
    return out or None


def normalize_extraction(raw: dict) -> dict:
    """LLM 원시 JSON → ir_signals.json 스키마로 정규화 + 가드레일.

    llm_judge.parse_judgment()과 같은 철학: 범위를 벗어나거나 목록에 없는
    값은 조용히 넘기지 않고 안전한 기본값(대개 None/unknown)으로 강등한다.
    """
    d = raw or {}
    segment = d.get("segment")
    if segment not in SEGMENTS:
        segment = "unknown"

    entities_raw = d.get("business_text_entities")
    entities = ([str(x).strip() for x in entities_raw if x and str(x).strip()]
               if isinstance(entities_raw, list) else [])

    return {
        "company_name": (d.get("company_name") or "").strip() or None,
        "business_text": (str(d.get("business_text")).strip()[:600]
                          if d.get("business_text") else None),
        "business_text_entities": entities,
        "founding_year": _coerce_int(d.get("founding_year")),
        "latest_round": (d.get("latest_round") or None),
        "ipo_mentioned": bool(d.get("ipo_mentioned", False)),
        "ipo_target_year": _coerce_int(d.get("ipo_target_year")),
        "net_income": _coerce_float(d.get("net_income")),
        "future_net_income": _coerce_float(d.get("future_net_income")),
        "ni_year": _coerce_int(d.get("ni_year")),
        "revenue": _coerce_float(d.get("revenue")),
        "net_income_projections": _coerce_year_dict(d.get("net_income_projections")),
        "revenue_projections": _coerce_year_dict(d.get("revenue_projections")),
        "segment": segment,
        "yoy_growth": 0.0,
        "_extraction_notes": str(d.get("extraction_notes", ""))[:300],
    }


def _parse_llm_json(text: str) -> dict:
    t = (text or "").strip()
    if t.startswith("```"):
        t = t.strip("`")
        t = t[4:].strip() if t[:4].lower() == "json" else t
    try:
        return json.loads(t)
    except Exception:
        m = re.search(r"\{.*\}", t, re.S)
        return json.loads(m.group(0)) if m else {}


def extract_raw_pdf_text(pdf_path: Path) -> str:
    """PDF에서 순수 텍스트만 결정론적으로 추출(LLM 아님) — 고유명사 검증용 코퍼스.
    pdfplumber 실패/미설치 시 빈 문자열 반환(검증만 생략, 크래시는 아님)."""
    try:
        import pdfplumber
    except ImportError:
        return ""
    try:
        texts = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                texts.append(page.extract_text() or "")
        return "\n".join(texts)
    except Exception:
        return ""


def _normalize_for_match(s: str) -> str:
    """공백 전부 제거 + 소문자화 — PDF 텍스트 추출 시 줄바꿈/띄어쓰기가
    레이아웃 때문에 원문과 달라지는 경우가 흔해서, 공백 유무는 무시하고
    문자 시퀀스만 비교한다."""
    return re.sub(r"\s+", "", s).lower()


def _tokenize(s: str) -> list:
    """한글/영문/숫자 토큰 단위로 대략 분리(공백·슬래시·쉼표·괄호·가운뎃점 기준).
    2글자 미만 토큰은 조사·단위 등 잡음일 가능성이 높아 제외한다."""
    return [t for t in re.split(r"[\s/,·()]+", s) if len(t) >= 2]


def verify_entities(entities: list, raw_text: str, overlap_threshold: float = 0.5):
    """business_text_entities 각각이 raw_text에 실제로 근거하는지 확인.

    완전 문자열 일치는 너무 엄격하다 — Gemini가 "수성 수지 사용한 ... 종이
    테이프"를 "수성 수지 종이테이프"로 압축하는 것처럼, 정당한 요약도 어순이
    바뀌거나 조사가 빠진다(실사례로 확인). 그래서 entity를 단어 단위로 쪼개
    각 단어가 raw_text에 있는지 보고, **일정 비율 이상 겹치면 "근거 있음"**으로
    본다 — 완전 날조(원문과 겹치는 단어가 거의/전혀 없음)만 잡아내는 게 목적이지,
    표현 차이까지 전부 잡으려는 게 아니다.

    LLM을 다시 부르는 게 아니라 순수 문자열 매칭 — "지어내지 마라"는 프롬프트
    지시가 반복적으로 뚫린 실사례(2026-07 케이더블유씨) 이후 추가된 안전장치.

    raw_text가 너무 짧으면(pdfplumber 실패·스캔본 등으로 추정) None을 반환한다
    — 이 경우 "전부 원문에 없음"으로 오판하는 것보다 "검증 자체를 못 했다"고
    정직하게 표시하는 게 안전하다."""
    if not entities:
        return {"verified": [], "unverified": [], "skipped": False}
    if len(raw_text.strip()) < 50:
        return None
    norm_raw = _normalize_for_match(raw_text)
    verified, unverified = [], []
    for e in entities:
        norm_e = _normalize_for_match(e)
        if norm_e and norm_e in norm_raw:
            verified.append(e)
            continue
        tokens = _tokenize(e)
        if not tokens:
            unverified.append(e)
            continue
        found = sum(1 for t in tokens if _normalize_for_match(t) in norm_raw)
        ratio = found / len(tokens)
        (verified if ratio >= overlap_threshold else unverified).append(e)
    return {"verified": verified, "unverified": unverified, "skipped": False}


class IrExtractor:
    """단일 IR 문서 텍스트 → 정규화된 신호 dict. 문서 해시 단위로 캐싱."""

    def __init__(self, model: str = GEMINI_LLM_MODEL, client=None,
                 cache_path: Path | None = IR_PARSER_CACHE,
                 max_retries: int = LLM_MAX_RETRIES, temperature: float = 0.0):
        self._client = client
        self.model = model
        self.cache_path = cache_path
        self.max_retries = max_retries
        self.temperature = temperature
        self.cache: dict[str, dict] = {}
        if cache_path and cache_path.exists():
            try:
                self.cache = json.loads(cache_path.read_text(encoding="utf-8"))
            except Exception:
                self.cache = {}

    def _client_(self):
        if self._client is None:
            from google import genai
            self._client = genai.Client()
        return self._client

    def _call(self, ir_text: str) -> dict:
        from google.genai import types
        prompt = PROMPT_TEMPLATE.format(ir_text=ir_text)
        last = None
        for _ in range(self.max_retries):
            try:
                resp = self._client_().models.generate_content(
                    model=self.model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=self.temperature,
                        response_mime_type="application/json",
                    ),
                )
                return _parse_llm_json(resp.text)
            except Exception as e:  # noqa: BLE001
                delay = _parse_retry_delay(e)
                if delay is None:
                    raise
                last = e
                time.sleep(delay + 1.0)
        raise last  # type: ignore[misc]

    def _flush(self):
        if not self.cache_path:
            return
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self.cache_path.write_text(json.dumps(self.cache, ensure_ascii=False), encoding="utf-8")

    def extract(self, ir_text: str, save: bool = True) -> dict:
        k = _file_hash(ir_text)
        if k in self.cache:
            return self.cache[k]
        raw = self._call(ir_text)
        result = normalize_extraction(raw)
        self.cache[k] = result
        if save:
            self._flush()
        return result

    def extract_from_file(self, md_path: str | Path, save: bool = True) -> dict:
        """텍스트(.md/.txt) 경로 — PDF 업로드가 안 될 때의 폴백용."""
        path = Path(md_path)
        if not path.exists():
            raise FileNotFoundError(f"IR 텍스트 파일을 찾을 수 없습니다: {path}")
        text = path.read_text(encoding="utf-8")
        return self.extract(text, save=save)

    def _wait_until_active(self, file_obj, timeout: float = 60.0, interval: float = 2.0):
        """PDF 업로드 직후 상태가 PROCESSING이면 ACTIVE될 때까지 짧게 폴링.
        문서(PDF/이미지)는 보통 즉시 ACTIVE라 대부분 바로 통과한다."""
        elapsed = 0.0
        while getattr(getattr(file_obj, "state", None), "name", None) == "PROCESSING" and elapsed < timeout:
            time.sleep(interval)
            elapsed += interval
            file_obj = self._client_().files.get(name=file_obj.name)
        return file_obj

    def _call_pdf(self, pdf_path: Path) -> dict:
        from google.genai import types
        # google-genai SDK가 업로드 시 파일명을 HTTP 헤더 값으로 그대로 실어
        # 보내는데, httpx가 헤더를 기본 ASCII로 인코딩하려다 한글 파일명
        # (IR 파일이 전부 회사명 기반이라 거의 항상 한글)에서 깨진다
        # (UnicodeEncodeError). SDK 버그라 우리 쪽에서 고칠 수 없으니, 업로드
        # 직전에 ASCII 전용 임시 파일명으로 복사해서 그걸 올리는 방식으로
        # 우회한다. 원본 파일 자체는 건드리지 않는다.
        import shutil
        import tempfile
        import os as _os

        fd, tmp_path_str = tempfile.mkstemp(suffix=".pdf", prefix="ir_upload_")
        _os.close(fd)
        tmp_path = Path(tmp_path_str)
        try:
            shutil.copyfile(pdf_path, tmp_path)

            last = None
            for _ in range(self.max_retries):
                try:
                    uploaded = self._client_().files.upload(file=str(tmp_path))
                    uploaded = self._wait_until_active(uploaded)
                    resp = self._client_().models.generate_content(
                        model=self.model,
                        contents=[uploaded, PDF_PROMPT_TEMPLATE],
                        config=types.GenerateContentConfig(
                            temperature=self.temperature,
                            response_mime_type="application/json",
                        ),
                    )
                    return _parse_llm_json(resp.text)
                except Exception as e:  # noqa: BLE001
                    delay = _parse_retry_delay(e)
                    if delay is None:
                        raise
                    last = e
                    time.sleep(delay + 1.0)
            raise last  # type: ignore[misc]
        finally:
            tmp_path.unlink(missing_ok=True)

    def extract_from_pdf(self, pdf_path: str | Path, save: bool = True,
                         md_path_for_verification: str | Path | None = None) -> dict:
        """PDF 원본을 직접 Gemini에 첨부해 추출 — OCR 오독 경로를 건너뛴다.
        캐시는 파일 내용 해시 기준(같은 PDF 재업로드 방지, 개정판은 자동 재추출).

        md_path_for_verification: 고유명사 검증(_entity_verification)의 코퍼스만
        보강하는 용도. 추출 자체(숫자·필드)는 여전히 PDF 원본에서만 하지만,
        디자인된 IR 슬라이드는 텍스트가 이미지로 렌더링돼 pdfplumber가 못 읽는
        구간이 실사례로 확인됨(2026-07 케이더블유씨 — "광동"/"Shamrock Milk"가
        실제 슬라이드엔 있는데 pdfplumber 추출 텍스트엔 없었음). OCR 결과(.md)는
        이미지 렌더링 텍스트도 읽어내니, 검증 코퍼스에만 보조로 더한다 — 숫자
        추출에 OCR을 다시 쓰는 게 아니라 "이 문자열이 문서 어딘가에 있냐"는
        존재 여부 확인에만 쓰는 것이라 역할이 다르다."""
        path = Path(pdf_path)
        if not path.exists():
            raise FileNotFoundError(f"IR PDF 파일을 찾을 수 없습니다: {path}")
        pdf_bytes = path.read_bytes()
        k = _pdf_hash(pdf_bytes)
        if k in self.cache:
            return self.cache[k]
        raw = self._call_pdf(path)
        result = normalize_extraction(raw)

        # business_text에 쓴 고유명사가 원문에 실제로 있는지 결정론적으로
        # 검증(LLM 재질의 아님, 순수 문자열 매칭). 검증 결과는 신호 스키마가
        # 아니라 감사용 메타데이터라 밑줄(_) 접두사로 구분.
        entities = result.get("business_text_entities") or []
        if entities:
            corpus_parts = [extract_raw_pdf_text(path)]
            md_used = False
            if md_path_for_verification:
                md_p = Path(md_path_for_verification)
                if md_p.exists():
                    corpus_parts.append(md_p.read_text(encoding="utf-8", errors="replace"))
                    md_used = True
            raw_text = "\n".join(corpus_parts)
            verification = verify_entities(entities, raw_text)
            if verification is not None:
                verification["corpus_included_ocr_md"] = md_used
            result["_entity_verification"] = verification
        else:
            result["_entity_verification"] = None

        self.cache[k] = result
        if save:
            self._flush()
        return result

    def extract_from_document(self, pdf_path: str | Path | None = None,
                              md_path: str | Path | None = None,
                              save: bool = True) -> dict:
        """PDF 우선, 실패하면 .md 텍스트로 폴백. 최소 하나는 있어야 한다.
        PDF 추출이 성공해도, md_path가 같이 주어졌으면 고유명사 검증 코퍼스에는
        같이 포함시킨다(디자인 슬라이드의 이미지 렌더링 텍스트를 OCR로 보강)."""
        if pdf_path is None and md_path is None:
            raise ValueError("pdf_path 또는 md_path 중 하나는 있어야 합니다.")
        if pdf_path is not None:
            try:
                return self.extract_from_pdf(pdf_path, save=save,
                                             md_path_for_verification=md_path)
            except Exception as e:  # noqa: BLE001
                if md_path is None:
                    raise
                print(f"[ir_parser] PDF 추출 실패({type(e).__name__}: {e}) "
                      f"— .md 텍스트로 폴백합니다: {md_path}")
                return self.extract_from_file(md_path, save=save)
        return self.extract_from_file(md_path, save=save)