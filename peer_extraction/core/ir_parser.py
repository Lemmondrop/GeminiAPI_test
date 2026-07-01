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
IR_PARSER_PROMPT_VERSION = "v1"

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
  "founding_year": 설립연도(정수) 또는 null,
  "latest_round": "가장 최근에 유치한 투자 라운드 명칭(예: Series A, Seed, Pre-IPO)을 문서 표현 그대로" 또는 null,
  "ipo_mentioned": IR에 IPO 계획·시점이 명시적으로 언급되어 있으면 true, 전혀 언급 없으면 false,
  "ipo_target_year": IPO 목표연도(정수, 언급된 경우만) 또는 null,
  "net_income": 가장 최근 결산/추정 기준 순이익(억원 단위 숫자, 적자면 음수) 또는 null,
  "future_net_income": IR에 명시된 미래 추정 순이익(억원, 특정 미래연도 기준) 또는 null,
  "ni_year": future_net_income이 어느 연도 추정치인지(정수) 또는 null,
  "revenue": 가장 최근 매출(억원 단위 숫자) 또는 null,
  "segment": 아래 12개 중 가장 가까운 업종 하나만 정확히 그대로 선택(다른 단어 사용 금지):
    ai, consumer, deeptech, fintech, manufacturing, materials, mobility, platform,
    robotics, semiconductor, software, unknown
    — 애매하면 unknown을 선택할 것. 절대 목록에 없는 단어를 만들지 말 것.
  "extraction_notes": "추출 근거나 애매했던 부분에 대한 한국어 1~2문장 메모" (없으면 빈 문자열)
}"""

_COMMON_INSTRUCTION = """문서에 명시적으로 나타난 정보만 추출하라. 추측하지 말고, 문서에 없으면
반드시 null로 남겨라 — 없는 값을 지어내는 것이 가장 위험한 실수다. 표·재무제표 안의 숫자는
단위(원/백만원/억원)를 확인해서 반드시 억원 단위로 환산해 넣어라.

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


def normalize_extraction(raw: dict) -> dict:
    """LLM 원시 JSON → ir_signals.json 스키마로 정규화 + 가드레일.

    llm_judge.parse_judgment()과 같은 철학: 범위를 벗어나거나 목록에 없는
    값은 조용히 넘기지 않고 안전한 기본값(대개 None/unknown)으로 강등한다.
    """
    d = raw or {}
    segment = d.get("segment")
    if segment not in SEGMENTS:
        segment = "unknown"

    return {
        "company_name": (d.get("company_name") or "").strip() or None,
        "founding_year": _coerce_int(d.get("founding_year")),
        "latest_round": (d.get("latest_round") or None),
        "ipo_mentioned": bool(d.get("ipo_mentioned", False)),
        "ipo_target_year": _coerce_int(d.get("ipo_target_year")),
        "net_income": _coerce_float(d.get("net_income")),
        "future_net_income": _coerce_float(d.get("future_net_income")),
        "ni_year": _coerce_int(d.get("ni_year")),
        "revenue": _coerce_float(d.get("revenue")),
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

    def extract_from_pdf(self, pdf_path: str | Path, save: bool = True) -> dict:
        """PDF 원본을 직접 Gemini에 첨부해 추출 — OCR 오독 경로를 건너뛴다.
        캐시는 파일 내용 해시 기준(같은 PDF 재업로드 방지, 개정판은 자동 재추출)."""
        path = Path(pdf_path)
        if not path.exists():
            raise FileNotFoundError(f"IR PDF 파일을 찾을 수 없습니다: {path}")
        pdf_bytes = path.read_bytes()
        k = _pdf_hash(pdf_bytes)
        if k in self.cache:
            return self.cache[k]
        raw = self._call_pdf(path)
        result = normalize_extraction(raw)
        self.cache[k] = result
        if save:
            self._flush()
        return result

    def extract_from_document(self, pdf_path: str | Path | None = None,
                              md_path: str | Path | None = None,
                              save: bool = True) -> dict:
        """PDF 우선, 실패하면 .md 텍스트로 폴백. 최소 하나는 있어야 한다."""
        if pdf_path is None and md_path is None:
            raise ValueError("pdf_path 또는 md_path 중 하나는 있어야 합니다.")
        if pdf_path is not None:
            try:
                return self.extract_from_pdf(pdf_path, save=save)
            except Exception as e:  # noqa: BLE001
                if md_path is None:
                    raise
                print(f"[ir_parser] PDF 추출 실패({type(e).__name__}: {e}) "
                      f"— .md 텍스트로 폴백합니다: {md_path}")
                return self.extract_from_file(md_path, save=save)
        return self.extract_from_file(md_path, save=save)