"""임베더.

운영 : GeminiEmbedder  — gemini-embedding-001, 1024차원, task_type 비대칭 인코딩.
                         분당 쿼터(paid 3000/min) 대비 슬라이딩윈도우 페이싱 + 429 자동 재시도.
개발 : HashingEmbedder — API 없이 전체 파이프라인을 돌리는 결정론 어휘 임베더(캐시 호환).

공통 인터페이스
  .encode(texts, task_type) -> np.ndarray[float32]   (L2 정규화됨 → 코사인 = 내적)
  .encode_corpus(texts) / .encode_query(texts)        편의 래퍼

재현성 메모
  - 같은 모델 버전 + 같은 입력 → 같은 벡터. 버전을 config.EMBED_MODEL에 핀.
  - MRL로 3072 미만 차원으로 자르면 unit-norm이 아니므로 코사인 위해 반드시 재정규화한다.
"""
from __future__ import annotations
import re
import sys
import time
from collections import deque
from typing import Literal, Sequence

import numpy as np

from config import (EMBED_BATCH, EMBED_MAX_PER_MIN, EMBED_MAX_RETRIES,
                    EMBED_MODEL, OUTPUT_DIM)

TaskType = Literal["RETRIEVAL_DOCUMENT", "RETRIEVAL_QUERY"]


def _l2_normalize(x: np.ndarray) -> np.ndarray:
    n = np.linalg.norm(x, axis=1, keepdims=True)
    n[n == 0] = 1.0
    return (x / n).astype(np.float32)


class _RateLimiter:
    """슬라이딩 윈도우 레이트리밋(콘텐츠 단위). 분당 한도 아래로 사전 페이싱."""

    def __init__(self, max_per_min: int):
        self.max = max_per_min
        self.win: deque[float] = deque()

    def acquire(self, n: int) -> None:
        now = time.monotonic()
        while self.win and now - self.win[0] > 60:
            self.win.popleft()
        # 이번 n건을 넣으면 분당 한도를 넘는다면, 가장 오래된 기록이 60s를 벗어날 때까지 대기
        while self.win and len(self.win) + n > self.max:
            sleep_for = 60 - (now - self.win[0]) + 0.2
            if sleep_for > 0:
                time.sleep(sleep_for)
            now = time.monotonic()
            while self.win and now - self.win[0] > 60:
                self.win.popleft()
        self.win.extend([now] * n)


def _parse_retry_delay(exc: Exception) -> float | None:
    """레이트리밋(429/RESOURCE_EXHAUSTED)이면 권장 대기초 반환, 아니면 None(재시도 안 함)."""
    s = str(exc)
    if "RESOURCE_EXHAUSTED" not in s and "429" not in s:
        return None
    m = re.search(r"retry in ([\d.]+)s", s) or re.search(r"'retryDelay':\s*'(\d+)s'", s)
    return float(m.group(1)) if m else 30.0


class GeminiEmbedder:
    """운영 권장. 코퍼스 벡터는 VectorCache가 1회만 생성·캐시하므로 매 추론 호출 안 함.

    client 인자: 테스트에서 가짜 클라이언트를 주입하기 위한 시드(미지정 시 genai.Client()).
    """

    def __init__(self, model: str = EMBED_MODEL, dim: int = OUTPUT_DIM,
                 batch: int = EMBED_BATCH, max_per_min: int = EMBED_MAX_PER_MIN,
                 max_retries: int = EMBED_MAX_RETRIES, client=None,
                 progress: bool = True):
        if client is None:
            from google import genai
            client = genai.Client()   # 환경변수 GEMINI_API_KEY(또는 GOOGLE_API_KEY)
        self.client = client
        self.model, self.dim, self.batch = model, dim, batch
        self.max_retries = max_retries
        self.progress = progress
        self._limiter = _RateLimiter(max_per_min)

    def fit(self, corpus: Sequence[str]) -> None:  # 인터페이스 호환용 no-op
        return None

    def _embed_once(self, chunk: list[str], task_type: TaskType) -> list[list[float]]:
        from google.genai import types
        last: Exception | None = None
        for _ in range(self.max_retries):
            self._limiter.acquire(len(chunk))            # 사전 페이싱
            try:
                resp = self.client.models.embed_content(
                    model=self.model,
                    contents=chunk,
                    config=types.EmbedContentConfig(
                        task_type=task_type,
                        output_dimensionality=self.dim,
                    ),
                )
                return [e.values for e in resp.embeddings]
            except Exception as e:                        # noqa: BLE001
                delay = _parse_retry_delay(e)
                if delay is None:                         # 레이트리밋 아니면 즉시 전파
                    raise
                last = e
                if self.progress:
                    print(f"\n  [429] 쿼터 한도 — {delay:.0f}s 대기 후 재시도", flush=True)
                time.sleep(delay + 1.0)
        raise last  # type: ignore[misc]

    def encode(self, texts: Sequence[str],
               task_type: TaskType = "RETRIEVAL_DOCUMENT") -> np.ndarray:
        texts = [("" if t is None else str(t)) for t in texts]
        texts = [t if t.strip() else " " for t in texts]  # 빈 Part 방지 (API 400 가드)
        out: list[list[float]] = []
        n = len(texts)
        for i in range(0, n, self.batch):
            out.extend(self._embed_once(texts[i:i + self.batch], task_type))
            done = min(i + self.batch, n)
            if self.progress and n > self.batch:
                pct = done * 100 // n
                sys.stdout.write(f"\r  ...임베딩 {done}/{n} ({pct}%)")
                sys.stdout.flush()
        if self.progress and n > self.batch:
            sys.stdout.write("\n")
        return _l2_normalize(np.asarray(out, dtype=np.float32))

    def encode_corpus(self, texts: Sequence[str]) -> np.ndarray:
        return self.encode(texts, "RETRIEVAL_DOCUMENT")

    def encode_query(self, texts: Sequence[str]) -> np.ndarray:
        return self.encode(texts, "RETRIEVAL_QUERY")


class HashingEmbedder:
    """오프라인 개발/CI용 결정론 어휘 임베더.

    char n-gram HashingVectorizer를 OUTPUT_DIM(1024) 고정 차원으로 투영한다.
      - 캐시 호환(차원=OUTPUT_DIM), fit 불필요, 같은 입력 → 같은 벡터(해시 고정).
      - API 키 없이 전체 파이프라인을 돌려볼 수 있어 CI/회귀테스트에 적합.

    한계: '어휘 중첩'만 포착한다. 키워드가 겹치지 않는 순수 의미 유사
          (예: 알루미늄 ↔ 2차전지의 '경량화/EV' 테마)는 못 잡는다.
          실제 의미검색은 GeminiEmbedder로. (해싱은 그 하한선 역할)
    """

    def __init__(self, dim: int = OUTPUT_DIM):
        from sklearn.feature_extraction.text import HashingVectorizer
        self.vec = HashingVectorizer(
            analyzer="char_wb", ngram_range=(2, 4),
            n_features=dim, alternate_sign=False, norm=None,
        )
        self.dim = dim

    def fit(self, corpus: Sequence[str]) -> None:  # 인터페이스 호환용 no-op
        return None

    def encode(self, texts: Sequence[str], task_type: TaskType | None = None) -> np.ndarray:
        m = self.vec.transform([str(t) for t in texts]).toarray().astype(np.float32)
        return _l2_normalize(m)

    def encode_corpus(self, texts: Sequence[str]) -> np.ndarray:
        return self.encode(texts)

    def encode_query(self, texts: Sequence[str]) -> np.ndarray:
        return self.encode(texts)