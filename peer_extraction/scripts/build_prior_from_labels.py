"""통합 prior 빌더 — 신고서 라벨 3종 + 발행사 KSIC → Lane C prior.

입력(data/labels):
  issuer_ksic.csv      : corp_code → issuer_sub (발행사 소분류, DART 기업개황)
  comparison_ksic.csv  : 발행사별 1차 업종유사성 KSIC 코드들 (주 신호)
  peer_labels.csv      : 발행사별 최종 선정 피어 회사명 (보강 신호)

각 IPO를 (issuer_sub, [comparison_ksic 코드... + 피어 회사명...])로 만들어
core.cooccur.build_prior에 투입 → {발행사_소분류: {피어_소분류: weight}}.
발행사 자기 소분류는 자동 제외(Lane A 담당).

실행: python scripts\\build_prior_from_labels.py --min-count 2
"""
from __future__ import annotations
import argparse
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

from core.cooccur import build_prior, name_to_sub_map, resolve_sub, save_prior


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--labels-dir", default="data/labels")
    ap.add_argument("--min-count", type=int, default=2)
    ap.add_argument("--no-normalize", action="store_true")
    ap.add_argument("--peers-only", action="store_true", help="comparison_ksic 빼고 피어명만")
    ap.add_argument("--ksic-only", action="store_true", help="피어명 빼고 comparison_ksic만")
    a = ap.parse_args()

    ldir = Path(a.labels_dir)
    iksic = pd.read_csv(ldir / "issuer_ksic.csv", encoding="utf-8-sig", dtype=str)
    corp2sub = {r["corp_code"]: str(r["issuer_sub"]) for _, r in iksic.iterrows()
                if str(r.get("issuer_sub", "")).strip() and len(str(r["issuer_sub"])) == 3}
    print(f"발행사 소분류 사전: {len(corp2sub)}개")

    comp = pd.read_csv(ldir / "comparison_ksic.csv", encoding="utf-8-sig", dtype=str) \
        if (ldir / "comparison_ksic.csv").exists() else pd.DataFrame()
    peers = pd.read_csv(ldir / "peer_labels.csv", encoding="utf-8-sig", dtype=str)

    # 발행사(corp_code)별 신호 모으기
    comp_by_corp = defaultdict(list)
    if not a.peers_only and len(comp):
        for _, r in comp.iterrows():
            comp_by_corp[r["corp_code"]].append(r["ksic"])
    peer_by_corp = defaultdict(list)
    if not a.ksic_only:
        for _, r in peers.iterrows():
            if str(r.get("peer", "")).strip():
                peer_by_corp[r["corp_code"]].append(r["peer"])

    # IPO 레코드: (issuer_sub, [comparison 코드 + 피어명])
    ipos = []
    used = 0
    for corp, isub in corp2sub.items():
        signals = comp_by_corp.get(corp, []) + peer_by_corp.get(corp, [])
        if signals:
            ipos.append((isub, signals))      # issuer_sub를 식별자로(코드라 resolve_sub가 그대로 처리)
            used += 1
    print(f"prior 입력 IPO: {used}개 (발행사 소분류 매칭됨)")

    name2sub = name_to_sub_map()
    prior = build_prior(ipos, name2sub, min_count=a.min_count,
                        normalize=not a.no_normalize)
    meta = prior.pop("_meta", {})
    out = save_prior(prior)

    n_sub = len(prior)
    covered = sorted(prior.keys())
    print(f"\nprior 발행사 소분류: {n_sub}개")
    print(f"커버된 소분류: {', '.join(covered[:40])}{' ...' if n_sub > 40 else ''}")
    shown = 0
    for isub, row in sorted(prior.items()):
        top = list(row.items())[:6]
        print(f"  {isub} → " + ", ".join(f"{p}({w})" for p, w in top))
        shown += 1
        if shown >= 15:
            break
    print(f"\n저장: {out}")
    print("pipeline 자동 로드 → Stage1 Lane C 활성.")


if __name__ == "__main__":
    main()
