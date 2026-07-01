"""코퍼스 임베딩 캐시 + FAISS 검색.

설계
  1) 상장 유니버스 텍스트를 1회 임베딩 → corpus_embeddings.npy 저장 후 재사용.
  2) corpus_meta.json에 {model, dim, corpus_hash, n_rows, embed_date} 동봉.
     - 모델/차원/코퍼스 중 하나라도 바뀌면 캐시 무효 → 자동 재생성.
  3) 정규화 벡터 → FAISS IndexFlatIP(내적 = 코사인)로 top-K 검색.

코퍼스 텍스트는 CORPUS_TEXT_COLS(주 신호)를 조인하되, 빈 행은 CORPUS_TEXT_FALLBACK_COLS로
코얼레스 폴백한다(빈 문자열은 Gemini가 'empty Part' 400으로 거부하므로 절대 비우지 않음).

재현성의 80%가 이 모듈에서 나온다: 코퍼스가 고정 캐시면 매 추론은 '쿼리 1건 임베딩 +
고정 행렬 코사인'이라 API를 쓰더라도 결과가 흔들릴 표면적이 거의 없다.
"""
from __future__ import annotations
import datetime as dt
import hashlib
import json
from pathlib import Path

import numpy as np
import pandas as pd

from config import (CORPUS_EMB_NPY, CORPUS_META_JSON, CORPUS_TEXT_COLS,
                    CORPUS_TEXT_FALLBACK_COLS, CSV_ENCODING, EMBED_MODEL,
                    LISTED_UNIVERSE_CSV, OUTPUT_DIM)


def _read_csv(path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(path, encoding=CSV_ENCODING)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="utf-8-sig")


def _clean(v) -> str:
    s = "" if v is None else str(v).strip()
    return "" if s.lower() == "nan" else s


def _corpus_texts(df: pd.DataFrame) -> list[str]:
    """행별 텍스트 구성: 주 컬럼 조인 → 비면 폴백 코얼레스 → 그래도 비면 공백(빈 Part 방지)."""
    primary = [c for c in CORPUS_TEXT_COLS if c in df.columns]
    fallback = [c for c in CORPUS_TEXT_FALLBACK_COLS if c in df.columns]

    def build(row) -> str:
        parts = [p for p in (_clean(row[c]) for c in primary) if p]
        if parts:
            return " | ".join(parts)
        for c in fallback:                  # 주 신호 빈 행만 폴백
            v = _clean(row[c])
            if v:
                return v
        return " "                          # 최후 폴백 (절대 빈 문자열 금지)

    return df.apply(build, axis=1).tolist()


def _hash_texts(texts: list[str]) -> str:
    h = hashlib.sha256()
    for t in texts:
        h.update(t.encode("utf-8"))
        h.update(b"\x00")
    return h.hexdigest()[:16]


class VectorCache:
    def __init__(self, embedder, csv_path: Path = LISTED_UNIVERSE_CSV):
        self.embedder = embedder
        self.df = _read_csv(csv_path).reset_index(drop=True)
        self.texts = _corpus_texts(self.df)
        self.corpus_hash = _hash_texts(self.texts)
        self.vectors: np.ndarray | None = None
        self.index = None

    # ----------------------------- 무결성 -----------------------------
    def _meta_valid(self) -> bool:
        if not (CORPUS_EMB_NPY.exists() and CORPUS_META_JSON.exists()):
            return False
        meta = json.loads(CORPUS_META_JSON.read_text(encoding="utf-8"))
        return (
            meta.get("model") == EMBED_MODEL
            and meta.get("dim") == OUTPUT_DIM
            and meta.get("corpus_hash") == self.corpus_hash
            and meta.get("n_rows") == len(self.df)
        )

    def _write_meta(self) -> None:
        meta = {
            "model": EMBED_MODEL,
            "dim": OUTPUT_DIM,
            "corpus_hash": self.corpus_hash,
            "n_rows": len(self.df),
            "embed_date": dt.datetime.now().isoformat(timespec="seconds"),
        }
        CORPUS_META_JSON.write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    # ----------------------------- 빌드/로드 -----------------------------
    def build_or_load(self, force: bool = False) -> np.ndarray:
        if not force and self._meta_valid():
            self.vectors = np.load(CORPUS_EMB_NPY)
        else:
            self.vectors = self.embedder.encode_corpus(self.texts)
            if self.vectors.shape[1] != OUTPUT_DIM:
                raise ValueError(
                    f"임베딩 차원 불일치: {self.vectors.shape[1]} != OUTPUT_DIM({OUTPUT_DIM})"
                )
            np.save(CORPUS_EMB_NPY, self.vectors)
            self._write_meta()
        self._build_index()
        return self.vectors

    def _build_index(self) -> None:
        import faiss
        d = self.vectors.shape[1]
        self.index = faiss.IndexFlatIP(d)   # 정규화 벡터 → 내적 = 코사인
        self.index.add(self.vectors)

    # ----------------------------- 검색 -----------------------------
    def search(self, query_vec: np.ndarray, top_k: int = 200):
        q = np.atleast_2d(query_vec).astype(np.float32)
        scores, idx = self.index.search(q, top_k)
        rows = self.df.iloc[idx[0]].copy()
        rows["embed_score"] = scores[0]
        return scores[0], rows