"""Gemini 생성 LLM 사업유사성 판정기 (Stage 3).

각 (발행사, 후보) 쌍을 사업모델 유사성으로 판정 → {tier, score, rationale} 구조화 JSON.

재현성 설계 (프로젝트 핵심 우려 해소):
  - 판정 결과를 (model, prompt_version, 발행사텍스트, 후보텍스트) 해시 키로 영구 캐싱.
  - 한 번 판정된 쌍은 캐시에서 읽음 → 재실행 시 동일 결과. LLM 비결정성 제거.
  - temperature=0 으로 첫 실행 분산도 최소화.
  - prompt_version을 키에 포함 → 프롬프트 개정 시 자동 재판정.

다수결 앙상블 (경계 구간 안정화):
  - temperature=0이어도 score 40~65 부근의 애매한 케이스는 호출마다
    tier가 뒤집히는 현상이 실측 확인됨(변동성 검증 스크립트 결과).
  - score가 ambiguous_range 안이면 추가 confirm_calls회 더 호출 후
    다수결 tier로 확정 → 1회 판정 시 캐시에 저장되므로 이후 재실행은
    여전히 100% 동일 결과(추가 호출은 최초 1회만 발생).
  - 명확한 케이스(대다수)는 1콜로 종료되어 비용 증가 미미.

client 인자: 테스트에서 가짜 클라이언트 주입(미지정 시 genai.Client()).
"""
from __future__ import annotations
import hashlib
import json
import re
import time
from collections import Counter

from config import (GEMINI_LLM_MODEL, LLM_JUDGE_CACHE, LLM_MAX_RETRIES,
                    LLM_PROMPT_VERSION)

TIERS = ("높음", "중간", "낮음")

# ── 다수결 앙상블 기본 설정 ──────────────────────────────────
# 변동성 검증(llm_judge_variance_test.py) 실측 기준:
#   케이사인 score 30~65 사이, 이니텍 score 20~50 사이에서 tier 전환 발생
#   → 여유를 둬 35~70 구간을 "애매" 판정 기준으로 설정
AMBIGUOUS_SCORE_RANGE = (35, 70)
CONFIRM_CALLS = 2          # 애매 시 추가 호출 수 (최초 1회 + 추가 2회 = 총 3회 다수결)

PROMPT_TEMPLATE = """당신은 IPO 공모가 산정의 비교기업(peer) 선정 전문가다.
아래 '발행사'와 '후보기업'의 사업 모델 유사성을 평가하라.
판단 기준: 제품/서비스의 본질, 전방산업(주요 고객), 밸류체인 단계, 수익 모델.
단순 표준산업분류 코드가 아니라 실제 사업 내용으로 판단한다.

[발행사]
{issuer}

[후보기업]
{candidate}

아래 JSON만 출력하라(설명·마크다운·코드펜스 금지):
{{"tier": "높음 또는 중간 또는 낮음", "score": 0부터 100 사이 정수, "rationale": "한국어 1~2문장 근거"}}
높음=핵심 사업이 직접 경쟁/대체 관계, 중간=일부 사업·밸류체인 중첩, 낮음=업종만 비슷하고 사업 상이.
"""


def _key(model: str, issuer: str, candidate: str) -> str:
    raw = f"{model}\x00{LLM_PROMPT_VERSION}\x00{issuer}\x00{candidate}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def _parse_retry_delay(exc: Exception) -> float | None:
    s = str(exc)
    # 429(한도초과): 명시된 retryDelay 우선
    if "RESOURCE_EXHAUSTED" in s or "429" in s:
        m = re.search(r"retry in ([\d.]+)s", s) or re.search(r"'retryDelay':\s*'(\d+)s'", s)
        return float(m.group(1)) if m else 30.0
    # 일시적 서버오류(과부하/내부오류): 고정 백오프 후 재시도
    if any(t in s for t in ("UNAVAILABLE", "503", "500", "INTERNAL", "DEADLINE_EXCEEDED")):
        return 15.0
    return None


def parse_judgment(text: str) -> dict:
    """LLM 원시 텍스트 → 정규화 dict. 코드펜스·필드오류·범위 가드."""
    t = (text or "").strip()
    if t.startswith("```"):                      # ```json ... ``` 제거
        t = t.strip("`")
        t = t[4:].strip() if t[:4].lower() == "json" else t
    try:
        d = json.loads(t)
    except Exception:
        m = re.search(r"\{.*\}", t, re.S)        # 본문 중 JSON 블록 구제
        d = json.loads(m.group(0)) if m else {}
    tier = d.get("tier", "낮음")
    if tier not in TIERS:
        tier = "낮음"
    try:
        score = int(float(d.get("score", 0)))
    except (TypeError, ValueError):
        score = 0
    score = max(0, min(100, score))
    return {"tier": tier, "score": score, "rationale": str(d.get("rationale", ""))[:300]}


class LlmJudge:
    def __init__(self, model: str = GEMINI_LLM_MODEL, client=None,
                 cache_path=LLM_JUDGE_CACHE, max_retries: int = LLM_MAX_RETRIES,
                 temperature: float = 0.0,
                 ambiguous_range: tuple[int, int] = AMBIGUOUS_SCORE_RANGE,
                 confirm_calls: int = CONFIRM_CALLS):
        self._client = client                    # lazy (genai 임포트 지연)
        self.model = model
        self.cache_path = cache_path
        self.max_retries = max_retries
        self.temperature = temperature
        self.ambiguous_range = ambiguous_range
        self.confirm_calls = confirm_calls
        self.cache: dict[str, dict] = {}
        if cache_path and cache_path.exists():
            try:
                self.cache = json.loads(cache_path.read_text(encoding="utf-8"))
            except Exception:
                self.cache = {}

    # ----------------------------- 내부 -----------------------------
    def _client_(self):
        if self._client is None:
            from google import genai
            self._client = genai.Client()        # GEMINI_API_KEY / GOOGLE_API_KEY
        return self._client

    def _call(self, issuer: str, candidate: str) -> dict:
        from google.genai import types
        prompt = PROMPT_TEMPLATE.format(issuer=issuer, candidate=candidate)
        last: Exception | None = None
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
                return parse_judgment(resp.text)
            except Exception as e:               # noqa: BLE001
                delay = _parse_retry_delay(e)
                if delay is None:
                    raise
                last = e
                time.sleep(delay + 1.0)
        raise last  # type: ignore[misc]

    def _call_with_consistency(self, issuer: str, candidate: str) -> dict:
        """1차 판정 → score가 애매 구간이면 추가 호출 후 다수결로 확정.

        - 명확한 케이스(대다수): 1콜로 즉시 반환, 비용 증가 없음.
        - 애매 구간(실측: score 35~70 부근에서 tier 전환 빈발): 추가
          confirm_calls회 더 호출(temperature=0이어도 비결정 발생 확인됨)
          → 다수결 tier 중 해당 tier의 첫 판정(rationale 포함) 반환.
        - 동점이면 score 평균이 가장 가까운 결과를 채택(안정적 단일값 보장).
        """
        first = self._call(issuer, candidate)
        lo, hi = self.ambiguous_range
        if not (lo <= first["score"] <= hi) or self.confirm_calls <= 0:
            return first

        results = [first]
        for _ in range(self.confirm_calls):
            results.append(self._call(issuer, candidate))

        tiers = [r["tier"] for r in results]
        counts = Counter(tiers)
        top_count = counts.most_common(1)[0][1]
        majority_tiers = [t for t, c in counts.items() if c == top_count]

        if len(majority_tiers) == 1:
            majority_tier = majority_tiers[0]
        else:
            # 완전 동률(예: 1-1-1) — TIERS 순위(높음>중간>낮음) 중 더 보수적인
            # (낮은) 등급 쪽 다수표가 있는 결과를 우선해 과대포함을 방지.
            tier_rank = {"높음": 2, "중간": 1, "낮음": 0}
            majority_tier = min(majority_tiers, key=lambda t: tier_rank[t])

        same_tier_results = [r for r in results if r["tier"] == majority_tier]
        scores = [r["score"] for r in same_tier_results]
        avg_score = sum(scores) / len(scores)
        # 평균 score에 가장 가까운 rationale을 대표값으로 채택
        final = min(same_tier_results, key=lambda r: abs(r["score"] - avg_score))
        final = dict(final)
        final["score"] = round(avg_score)
        final["_consistency_votes"] = tiers          # 디버깅/감사용 (캐시에도 저장됨)
        return final

    def _flush(self):
        if not self.cache_path:
            return
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self.cache_path.write_text(json.dumps(self.cache, ensure_ascii=False), encoding="utf-8")

    # ----------------------------- 공개 -----------------------------
    def judge(self, issuer_text: str, candidate_text: str, save: bool = True) -> dict:
        k = _key(self.model, issuer_text, candidate_text)
        if k in self.cache:
            return self.cache[k]
        result = self._call_with_consistency(issuer_text, candidate_text)
        self.cache[k] = result
        if save:
            self._flush()
        return result

    def judge_many(self, issuer_text: str,
                   candidates: list[tuple[str, str]],
                   progress: bool = True) -> dict[str, dict]:
        """candidates: [(id, business_text)] → {id: 판정}. 캐시 미스만 호출.

        진행 표시: 신규 LLM 호출은 느리므로(건당 수초, 애매 구간은 최대 3콜)
        처리 중인 건수를 갱신 출력.
        중단 대비: 신규 5건마다 캐시 저장 → Ctrl+C로 끊겨도 진행분 보존·재개.
        """
        import sys
        out: dict[str, dict] = {}
        n = len(candidates)
        new = 0
        consistency_triggered = 0
        for i, (cid, ctext) in enumerate(candidates, 1):
            k = _key(self.model, issuer_text, ctext)
            if k in self.cache:
                out[cid] = self.cache[k]
            else:
                if progress:
                    sys.stdout.write(f"\r  ...LLM 사업유사성 판정 {i}/{n} (신규 호출 중)")
                    sys.stdout.flush()
                out[cid] = self._call_with_consistency(issuer_text, ctext)
                if "_consistency_votes" in out[cid]:
                    consistency_triggered += 1
                self.cache[k] = out[cid]
                new += 1
                if new % 5 == 0 and self.cache_path is not None:
                    self._flush()                      # 주기적 저장(재개 안전)
        if progress:
            if new:
                extra = (f", 경계구간 다수결 {consistency_triggered}건"
                         if consistency_triggered else "")
                sys.stdout.write(f"\r  LLM 판정 완료: {n}건 (신규 {new}, 캐시 {n - new}{extra})        \n")
            else:
                sys.stdout.write(f"\r  LLM 판정: {n}건 전부 캐시 적중 (추가 호출 0)        \n")
            sys.stdout.flush()
        if new and self.cache_path is not None:
            self._flush()
        return out