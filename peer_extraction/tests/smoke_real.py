"""실환경 스모크 검증 — 실제 Gemini(gemini-embedding-001) + 실제 listed_universe.csv.

목적: union 로직(이미 test_stage1.py로 통과)이 아니라, Lane B 임베딩이 실제로
      '교차업종 피어'를 의미적으로 잘 끌어오는지를 육안으로 확인한다.

선행
  1) pip install -r requirements.txt
  2) GEMINI_API_KEY는 상위 'OCR Sample/.env'에서 자동 로드됨 (config._load_env).
     별도 export 불필요 — .env 파일에 GEMINI_API_KEY=... 만 있으면 됨.
  3) data/listed_universe.csv 배치  (컬럼: 회사명/시장구분/산업분류코드/업종/주요제품)

실행 (peer_extraction 루트)
  python tests/smoke_real.py

처음 1회는 코퍼스 전체를 임베딩하므로 수십 초~수 분 소요(이후 캐시 적중 → 즉시).
"""
from __future__ import annotations
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import config
from core.embedder import GeminiEmbedder
from core.vector_cache import VectorCache
from stages.stage1_population import StartupQuery, build_population, save_pool

import numpy as np

# ===== 캘리브레이션 모드 =====
# CALIBRATION_FIRM을 지정하면, 그 상장사를 '가상 스타트업'으로 삼아(코드·텍스트를 쿼리로)
# 코퍼스에 둔 채 자기 자신은 순위에서 제외하고, 알려진 정답 피어(KNOWN_PEERS)가
# 임베딩만으로 몇 위까지 올라오는지 게이트 1-B를 채점한다.
CALIBRATION_FIRM = "한라캐스트"
KNOWN_PEERS = ["알루코", "조일알미늄", "파워로직스", "삼현"]  # 한라캐스트 신고서 최종 피어
# 어휘(해싱) 기준선 — 전체 3,281행 기준 전역 순위. Gemini가 이보다 위로 올려야 의미검색이 일한 것:
LEXICAL_BASELINE = {"알루코": 135, "조일알미늄": 119, "파워로직스": 1439, "삼현": 103}

# 자유 쿼리 모드 (CALIBRATION_FIRM=None일 때 사용)
SEED_CODE = "C00303"
BUSINESS_TEXT = "자동차 차체용 알루미늄 다이캐스팅 경량화 부품. 전기차(EV) 경량 구조부품."

# 게이트는 임베딩 단독 순위(진단표)로 Lane B 개선폭을 보고, 풀은 prior 앵커까지 켜서
# production을 대표하게 한다. (한라식 확장 선례: 303 발행사 → 242·281 피어로 확대)
COOCCUR_PRIOR = {"303": {"242": 1.0, "281": 1.0}}
TOP_K_EMBED = 30  # lane_embed가 'top-30 진입'을 뜻하도록 좁게


def resolve_query(vc):
    if CALIBRATION_FIRM:
        r = vc.df[vc.df[config.NAME_COL] == CALIBRATION_FIRM]
        if r.empty:
            sys.exit(f"[중단] CALIBRATION_FIRM '{CALIBRATION_FIRM}' 가 CSV에 없음")
        r = r.iloc[0]
        text = f"{r['업종']} | {r['주요제품']}"
        print(f"[캘리브레이션] {CALIBRATION_FIRM} (코드 {r[config.CODE_COL]}) 를 가상 스타트업으로 사용")
        return StartupQuery(seed_code=str(r[config.CODE_COL]), business_text=text)
    return StartupQuery(seed_code=SEED_CODE, business_text=BUSINESS_TEXT)


def gate_diagnostic(vc, q, pool):
    """알려진 피어의 임베딩 전역 순위 + lane_embed 포착 여부 채점."""
    if not (CALIBRATION_FIRM and KNOWN_PEERS):
        return
    qv = vc.embedder.encode_query([q.business_text])[0]
    scores = vc.vectors @ qv
    names = vc.df[config.NAME_COL].tolist()
    score_by_name = dict(zip(names, scores))
    # 자기 자신(가상 스타트업)은 순위에서 제외
    order = [i for i in np.argsort(-scores) if names[i] != CALIBRATION_FIRM]
    rank_by_name = {names[idx]: r + 1 for r, idx in enumerate(order)}
    in_pool = set(pool[config.NAME_COL])

    print("\n=== 게이트 1-B 진단: 임베딩만으로 진짜 피어가 상위에 오는가 ===")
    print(f"{'피어':12s} {'임베딩순위':>9s} {'어휘기준선':>9s} {'score':>7s}  lane_embed")
    hits = 0
    for nm in KNOWN_PEERS:
        rk = rank_by_name.get(nm)
        base = LEXICAL_BASELINE.get(nm, "-")
        sc = score_by_name.get(nm, float("nan"))
        row = pool[pool[config.NAME_COL] == nm]
        le = bool(row.iloc[0]["lane_embed"]) if not row.empty else False
        hits += int(le)
        arrow = ""
        if isinstance(base, int) and rk is not None:
            arrow = " ↑" if rk < base else " ↓"
        print(f"{nm:12s} {str(rk):>9s} {str(base):>9s} {sc:7.3f}  {le}{arrow}")
    print(f"\n합격선: 알루코·조일·파워로직스가 어휘기준선보다 위로 + 가능하면 top-{TOP_K_EMBED} 진입.")
    print(f"임베딩 top-{TOP_K_EMBED} 포착: {hits}/{len(KNOWN_PEERS)}  "
          f"(특히 파워로직스=키워드 무겹침 케이스가 핵심)")


def main():
    if not config.LISTED_UNIVERSE_CSV.exists():
        sys.exit(f"[중단] CSV 없음: {config.LISTED_UNIVERSE_CSV}")

    print(f"모델={config.EMBED_MODEL}  차원={config.OUTPUT_DIM}  텍스트컬럼={config.CORPUS_TEXT_COLS}")
    vc = VectorCache(GeminiEmbedder())
    print(f"코퍼스 {len(vc.df)}행 임베딩/로드 중... (최초 1회만 API 호출)")
    vc.build_or_load()
    print("캐시 준비 완료.\n")

    q = resolve_query(vc)
    cfg = config.Stage1Config(top_k_embed=TOP_K_EMBED)
    pool = build_population(q, vc, cooccur_prior=COOCCUR_PRIOR, cfg=cfg)

    print(f"\n모집단 {len(pool)}행 생성. 임베딩 상위 20:\n")
    show = pool.head(20).copy()
    show["embed_score"] = show["embed_score"].round(3)
    with_pd_opts(lambda: print(show.to_string(index=False)))

    gate_diagnostic(vc, q, pool)

    out = save_pool(pool, ROOT / "data" / "stage1_pool.csv")
    print(f"\n전체 모집단 저장: {out}")


def with_pd_opts(fn):
    import pandas as pd
    with pd.option_context("display.max_columns", None, "display.width", 200,
                           "display.unicode.east_asian_width", True):
        fn()


if __name__ == "__main__":
    main()