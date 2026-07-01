"""Stage 3/4 정답 손실 진단 — 재무 통과한 정답이 사업유사성·PER에서 왜 떨어지나.

한 발행사를 끝까지 돌려 정답 피어가 Stage 3(tier)·Stage 4(drop_reason) 어디서 죽는지 추적.
특히 Stage 4 하드컷(시총과대/PER과소/과대)이 중·대형 정답 피어를 쳐내는지 확인.

실행: python scripts\\diag_stage34_loss.py --issuer 노을
"""
from __future__ import annotations
import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from core.cooccur import _norm_name, load_prior
from stages.stage1_population import StartupQuery, build_population
from stages.stage2_financial import apply_stage2, load_financials, survivors as s2_surv
from stages.stage3_business import apply_stage3, final_peers as s3_final
from stages.stage4_general import apply_stage4

_ANCHORS = ("사업의 개요", "회사의 개요", "사업의 내용", "주요 제품")


def overview(rcept):
    fp = config.DART_DOC_CACHE_DIR / f"{rcept}.txt"
    if not rcept or not fp.exists():
        return ""
    t = fp.read_text(encoding="utf-8")
    for kw in _ANCHORS:
        i = t.find(kw)
        if i >= 0:
            return re.sub(r"\s+", " ", t[i:i + 2500])
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--issuer", required=True)
    a = ap.parse_args()
    ldir = Path("data/labels")

    peers = pd.read_csv(ldir / "peer_labels.csv", encoding="utf-8-sig", dtype=str)
    iksic = pd.read_csv(ldir / "issuer_ksic.csv", encoding="utf-8-sig", dtype=str)
    seed = {r["corp_code"]: (str(r.get("induty_code") or "").strip()
                             or str(r.get("issuer_sub") or "").strip())
            for _, r in iksic.iterrows()}
    cache = json.loads(config.FILING_LABELS_CACHE.read_text(encoding="utf-8"))
    rcept = {c: r.get("initial_rcept", "") for c, r in cache.items()}

    g = peers[peers["issuer"] == a.issuer]
    if not len(g):
        sys.exit(f"발행사 '{a.issuer}' 없음")
    corp = g["corp_code"].iloc[0]
    gt_names = {_norm_name(p) for p in g["peer"]}
    print(f"발행사 {a.issuer} (corp {corp}) / 정답 피어: {', '.join(g['peer'])}\n")

    from core.embedder import GeminiEmbedder
    from core.vector_cache import VectorCache
    from core.llm_judge import LlmJudge
    vc = VectorCache(GeminiEmbedder()); vc.build_or_load()
    fin = load_financials()
    prior = load_prior() or None

    btext = overview(rcept.get(corp, ""))
    q = StartupQuery(seed_code=seed[corp], business_text=btext)
    pool = build_population(q, vc, cooccur_prior=prior, cfg=config.Stage1Config())
    pool = pool[pool[config.NAME_COL].map(_norm_name) != _norm_name(a.issuer)]
    s2 = apply_stage2(pool, fin, "general")
    surv = s2_surv(s2)
    s3 = apply_stage3(surv, btext, LlmJudge(), config.Stage3Config())
    s4 = apply_stage4(s3_final(s3), config.Stage4Config())

    # 정답 피어 추적
    print(f"{'정답피어':<14} {'모집단':>5} {'재무':>5} {'S3tier':>7} {'S4통과':>6} {'S4사유':<20}")
    print("-" * 70)
    s2["_n"] = s2[config.NAME_COL].map(_norm_name)
    s3["_n"] = s3[config.NAME_COL].map(_norm_name)
    s4["_n"] = s4[config.NAME_COL].map(_norm_name)
    for _, r in g.iterrows():
        n = _norm_name(r["peer"])
        in_pool = n in set(pool[config.NAME_COL].map(_norm_name))
        s2row = s2[s2["_n"] == n]
        in_s2 = bool(len(s2row)) and bool(s2row["passed_stage2"].iloc[0])
        s3row = s3[s3["_n"] == n]
        tier = s3row["tier"].iloc[0] if len(s3row) else "-"
        s4row = s4[s4["_n"] == n]
        s4pass = bool(s4row["passed_stage4"].iloc[0]) if len(s4row) else False
        s4reason = (s4row["drop_reason"].iloc[0] if len(s4row) else "(S3서 탈락)") or ""
        print(f"{str(r['peer'])[:14]:<14} {'O' if in_pool else 'X':>5} "
              f"{'O' if in_s2 else 'X':>5} {str(tier):>7} {'O' if s4pass else 'X':>6} {str(s4reason)[:20]:<20}")

    # Stage4 전체 컷 내역
    from collections import Counter
    cuts = Counter(r for r in s4["drop_reason"] if r and "max_peers" not in r)
    print("\nStage4 전체 컷 내역:", "  ".join(f"{k}×{v}" for k, v in cuts.most_common()))
    print(f"최종 비교군: {', '.join(s4[s4['passed_stage4']][config.NAME_COL].tolist())}")


if __name__ == "__main__":
    main()
