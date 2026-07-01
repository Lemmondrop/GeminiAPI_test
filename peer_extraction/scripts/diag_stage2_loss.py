"""Stage 2 정답 손실 진단 — 모집단엔 든 정답 피어가 '왜' 재무필터에서 잘렸나.

모집단 통과했지만 재무 탈락한 정답 피어들의 drop_reason을 집계.
  - '적자/PER≤0' 위주 → 시점 불일치(IPO 당시 흑자→현재 적자). 검증 방법 보정 대상.
  - '결산월/감사의견/매출' 위주 → 필터 과도. Stage 2 기준 완화 대상.

실행: python scripts\\diag_stage2_loss.py --limit 30
"""
from __future__ import annotations
import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from core.cooccur import _norm_name, load_prior
from stages.stage1_population import StartupQuery, build_population
from stages.stage2_financial import apply_stage2, load_financials

_ANCHORS = ("사업의 개요", "회사의 개요", "사업의 내용", "주요 제품")


def overview(rcept):
    if not rcept:
        return ""
    fp = config.DART_DOC_CACHE_DIR / f"{rcept}.txt"
    if not fp.exists():
        return ""
    t = fp.read_text(encoding="utf-8")
    for kw in _ANCHORS:
        i = t.find(kw)
        if i >= 0:
            return re.sub(r"\s+", " ", t[i:i + 2500])
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=30)
    a = ap.parse_args()
    ldir = Path("data/labels")

    peers = pd.read_csv(ldir / "peer_labels.csv", encoding="utf-8-sig", dtype=str)
    iksic = pd.read_csv(ldir / "issuer_ksic.csv", encoding="utf-8-sig", dtype=str)
    seed = {r["corp_code"]: (str(r.get("induty_code") or "").strip()
                             or str(r.get("issuer_sub") or "").strip())
            for _, r in iksic.iterrows()}
    cache = json.loads(config.FILING_LABELS_CACHE.read_text(encoding="utf-8")) \
        if config.FILING_LABELS_CACHE.exists() else {}
    rcept = {c: r.get("initial_rcept", "") for c, r in cache.items()}

    uni = pd.read_csv(config.LISTED_UNIVERSE_CSV, encoding=config.CSV_ENCODING, dtype=str)
    uni_norm = {_norm_name(n) for n in uni[config.NAME_COL].dropna()}

    from core.embedder import GeminiEmbedder
    from core.vector_cache import VectorCache
    vc = VectorCache(GeminiEmbedder()); vc.build_or_load()
    fin = load_financials()
    prior = load_prior() or None

    gt = {}
    for issuer, g in peers.groupby("issuer"):
        corp = g["corp_code"].iloc[0]
        avail = {_norm_name(p) for p in g["peer"] if _norm_name(p) in uni_norm}
        if len(avail) >= 2 and corp in seed:
            gt[issuer] = (corp, avail)

    reason_counter = Counter()
    lost_total = 0
    samples = []
    for k, (issuer, (corp, avail)) in enumerate(list(gt.items())[:a.limit], 1):
        q = StartupQuery(seed_code=seed[corp], business_text=overview(rcept.get(corp, "")))
        pool = build_population(q, vc, cooccur_prior=prior, cfg=config.Stage1Config())
        pool = pool[pool[config.NAME_COL].map(_norm_name) != _norm_name(issuer)]
        s2 = apply_stage2(pool, fin, "general")
        s2["_n"] = s2[config.NAME_COL].map(_norm_name)
        # 모집단엔 든 정답 중 재무 탈락분
        in_pool = s2[s2["_n"].isin(avail)]
        lost = in_pool[~in_pool["passed_stage2"]]
        for _, r in lost.iterrows():
            reason_counter[str(r.get("drop_reason", "?"))] += 1
            lost_total += 1
            if len(samples) < 25:
                samples.append((issuer, r[config.NAME_COL], r.get("drop_reason", "?")))

    print(f"\n검증 {min(a.limit, len(gt))}개 발행사 / 재무탈락 정답 피어 {lost_total}개\n")
    print("탈락 사유 집계:")
    for reason, cnt in reason_counter.most_common():
        print(f"  {cnt:>3}  {reason}")
    print("\n샘플(발행사 / 피어 / 사유):")
    for iss, peer, rs in samples:
        print(f"  {iss[:14]:<14} {str(peer)[:14]:<14} {rs}")


if __name__ == "__main__":
    main()
