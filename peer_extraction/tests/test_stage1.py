"""Stage 1 union 로직 결정론 테스트 (API 불필요).

목 임베더로 '어느 회사가 어느 레인에 잡혀야 하는가'를 사전 확정하고 단언으로 검증한다.
임베딩 의미품질이 아니라 union 메커니즘(코드/임베딩/prior/cap/provenance)의 정확성을 본다.

실행: peer_extraction 루트에서  python tests/test_stage1.py
"""
from __future__ import annotations
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import numpy as np
import pandas as pd

import config
from core.vector_cache import VectorCache
from stages.stage1_population import StartupQuery, build_population

# ----------------------------- 목 임베더 -----------------------------
# 테마 → 1024차원 기저벡터. 같은 테마=완전동일 벡터(코사인 1.0), 다른 테마=직교(0.0).
THEME_IDX = {"ALU": 0, "ACT": 1, "FOOD": 2}


def theme_of(text: str) -> str:
    t = str(text)
    if any(k in t for k in ["알루미늄", "경량", "2차전지", "배터리"]):
        return "ALU"
    if "액추에이터" in t:
        return "ACT"
    return "FOOD"


class MockEmbedder:
    def __init__(self, dim: int = config.OUTPUT_DIM):
        self.dim = dim

    def _vec(self, text: str) -> np.ndarray:
        v = np.zeros(self.dim, dtype="float32")
        v[THEME_IDX[theme_of(text)]] = 1.0
        return v

    def encode_corpus(self, texts):
        return np.vstack([self._vec(t) for t in texts])

    def encode_query(self, texts):
        return np.vstack([self._vec(t) for t in texts])


PRIOR = {"303": {"242": 1.0, "281": 1.0}}  # 한라식 확장 선례: 303 발행사 → 242·281로 확대


import tempfile

_TMP_CSV = Path(tempfile.gettempdir()) / "stage1_test_universe.csv"


def make_vcache(rows: list[dict]) -> VectorCache:
    df = pd.DataFrame(rows)
    # 테스트 격리: 실데이터 경로(config.LISTED_UNIVERSE_CSV)를 절대 건드리지 않고 임시 CSV 사용.
    df.to_csv(_TMP_CSV, index=False, encoding="cp949")
    vc = VectorCache(MockEmbedder(), csv_path=_TMP_CSV)
    vc.build_or_load(force=True)
    return vc


def flags(pool: pd.DataFrame, name: str) -> dict:
    r = pool[pool[config.NAME_COL] == name]
    if r.empty:
        return {}
    r = r.iloc[0]
    return {c: bool(r[c]) for c in
            ["lane_code_exact", "lane_code_mid", "lane_cooccur", "lane_embed", "in_pool"]}


# ----------------------------- 검증 -----------------------------
results: list[tuple[str, bool, str]] = []


def check(name: str, cond: bool, detail: str = ""):
    results.append((name, bool(cond), detail))


def run_base_case():
    """기본 유니버스: 코드/임베딩/prior 각 레인이 독립적으로 동작하는지."""
    rows = [
        # 삼현: 발행사와 같은 소분류(303) → code_exact, 사업은 액추에이터(임베딩 낮음)
        {"회사명": "삼현", "시장구분": "KOSDAQ", "산업분류코드": "C00303",
         "업종": "자동차부품", "주요제품": "스마트 액추에이터"},
        # 알루코·조일·파워로직스: 코드 다름(242/281) → 임베딩으로만 + prior로 cooccur
        {"회사명": "알루코", "시장구분": "KOSDAQ", "산업분류코드": "C00242",
         "업종": "비철금속", "주요제품": "알루미늄 압출재 경량화"},
        {"회사명": "조일알미늄", "시장구분": "KOSDAQ", "산업분류코드": "C00242",
         "업종": "비철금속", "주요제품": "알루미늄판"},
        {"회사명": "파워로직스", "시장구분": "KOSDAQ", "산업분류코드": "C00281",
         "업종": "전자부품", "주요제품": "2차전지 보호회로 배터리팩"},
        # 미드사: 같은 중분류(30) 다른 소분류(301) → mid fallback, 임베딩 낮음
        {"회사명": "미드사", "시장구분": "KOSDAQ", "산업분류코드": "C00301",
         "업종": "기타", "주요제품": "라면 스낵"},
        # 프라이어사: 242(prior에 있음) + 사업 무관(FOOD) → cooccur 단독 포착
        {"회사명": "프라이어사", "시장구분": "KOSDAQ", "산업분류코드": "C00242",
         "업종": "음식료", "주요제품": "과자"},
        # 음식료: 코드·임베딩·prior 어디에도 안 걸림 → 풀에서 제외돼야 함
        {"회사명": "음식료", "시장구분": "KOSDAQ", "산업분류코드": "C00999",
         "업종": "음식료", "주요제품": "라면"},
    ]
    vc = make_vcache(rows)
    q = StartupQuery(seed_code="C00303", business_text="자동차 경량화 알루미늄 다이캐스팅 부품")
    # ALU 회사 3개 → top_k=3이면 정확히 그 3개만 임베딩 히트
    cfg = config.Stage1Config(top_k_embed=3, use_mid_fallback=True, recall_cap=250)
    pool = build_population(q, vc, cooccur_prior=PRIOR, cfg=cfg)

    f_sam = flags(pool, "삼현")
    check("A. code_exact: 삼현(303) exact 플래그", f_sam.get("lane_code_exact") is True, str(f_sam))
    check("A. code_exact: 삼현은 mid 아님(배타)", f_sam.get("lane_code_mid") is False)

    f_mid = flags(pool, "미드사")
    check("B. code_mid: 미드사(301) mid fallback", f_mid.get("lane_code_mid") is True, str(f_mid))
    check("B. code_mid: 미드사 exact 아님", f_mid.get("lane_code_exact") is False)

    f_alu = flags(pool, "알루코")
    check("C. embed: 알루코 임베딩 히트", f_alu.get("lane_embed") is True, str(f_alu))
    check("C. embed: 삼현(ACT) 임베딩 미히트", f_sam.get("lane_embed") is False)
    check("C. embed: 음식료(FOOD) 임베딩 미히트",
          flags(pool, "음식료").get("lane_embed") in (False, None))

    check("D. cooccur: 알루코(242) prior 포착", f_alu.get("lane_cooccur") is True)
    check("D. cooccur: 파워로직스(281) prior 포착",
          flags(pool, "파워로직스").get("lane_cooccur") is True)
    check("D. cooccur: 미드사(301) prior 미포착", f_mid.get("lane_cooccur") is False)

    f_pri = flags(pool, "프라이어사")
    check("E. cooccur 단독: 프라이어사 in_pool", f_pri.get("in_pool") is True, str(f_pri))
    check("E. cooccur 단독: 코드·임베딩 전부 F",
          f_pri.get("lane_code_exact") is False and f_pri.get("lane_code_mid") is False
          and f_pri.get("lane_embed") is False and f_pri.get("lane_cooccur") is True)

    check("F. 노이즈 제외: 음식료 풀에서 빠짐", flags(pool, "음식료") == {})

    # 교차업종 핵심 명제: 진짜 피어(알루코)가 발행사 코드 밖에서 임베딩으로 잡힘
    check("G. 교차업종: 알루코 code_exact=F 이지만 in_pool=T(임베딩)",
          f_alu.get("lane_code_exact") is False and f_alu.get("in_pool") is True)

    # 출력 정렬: embed_score 내림차순
    scores = pool["embed_score"].tolist()
    check("I. 출력 embed_score 내림차순", scores == sorted(scores, reverse=True))
    check("I. 출력 컬럼 구비",
          set(["회사명", "시장구분", "산업분류코드", "embed_score",
               "lane_code_exact", "lane_code_mid", "lane_cooccur", "lane_embed", "in_pool"])
          <= set(pool.columns))

    return pool


def run_cap_case():
    """recall_cap: 결정론 히트는 전원 유지, 임베딩-only는 점수순 cap까지만."""
    rows = [
        {"회사명": "삼현", "시장구분": "KOSDAQ", "산업분류코드": "C00303",
         "업종": "자동차부품", "주요제품": "스마트 액추에이터"},        # exact (det)
        {"회사명": "미드사", "시장구분": "KOSDAQ", "산업분류코드": "C00301",
         "업종": "기타", "주요제품": "라면"},                          # mid (det)
        {"회사명": "프라이어사", "시장구분": "KOSDAQ", "산업분류코드": "C00242",
         "업종": "음식료", "주요제품": "과자"},                        # cooccur (det)
    ]
    # 임베딩-only ALU 회사 다수: 코드 255(중분류25, prior에 없음) → det 아님, embed만
    for i in range(8):
        rows.append({"회사명": f"임베드ALU{i}", "시장구분": "KOSDAQ", "산업분류코드": "C00255",
                     "업종": "비철금속", "주요제품": "알루미늄 경량 소재"})
    vc = make_vcache(rows)
    q = StartupQuery(seed_code="C00303", business_text="알루미늄 경량화")
    # det 3개 + 임베딩 8개 후보 중, cap=5 → det 3 전원 + embed-only 2개만
    cfg = config.Stage1Config(top_k_embed=8, use_mid_fallback=True, recall_cap=5)
    pool = build_population(q, vc, cooccur_prior=PRIOR, cfg=cfg)

    det_names = {"삼현", "미드사", "프라이어사"}
    in_names = set(pool["회사명"])
    check("H. cap: 풀 크기 == recall_cap(5)", len(pool) == 5, f"len={len(pool)}")
    check("H. cap: 결정론 히트 3종 전원 생존", det_names <= in_names, str(in_names))
    embed_kept = [n for n in in_names if n.startswith("임베드ALU")]
    check("H. cap: 임베딩-only는 cap 잔여(2개)만", len(embed_kept) == 2, str(embed_kept))


def main():
    run_base_case()
    run_cap_case()

    # 정리 (실데이터 CSV는 건드리지 않음 — 캐시와 임시파일만)
    for p in (config.CORPUS_EMB_NPY, config.CORPUS_META_JSON, _TMP_CSV):
        Path(p).unlink(missing_ok=True)

    print("=" * 64)
    passed = sum(1 for _, ok, _ in results if ok)
    for name, ok, detail in results:
        mark = "PASS" if ok else "FAIL"
        line = f"[{mark}] {name}"
        if not ok and detail:
            line += f"   <- {detail}"
        print(line)
    print("=" * 64)
    print(f"{passed}/{len(results)} passed")
    sys.exit(0 if passed == len(results) else 1)


if __name__ == "__main__":
    main()
