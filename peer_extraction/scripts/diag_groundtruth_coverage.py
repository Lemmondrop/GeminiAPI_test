"""ground-truth 검증 가능성 진단 — 정답 피어가 우리 모집단에 얼마나 존재하나.

검증의 분모(우리가 뽑을 수 있었던 정답)를 먼저 측정.
peer_labels의 피어 회사명을 유니버스(company_cord_prototype)와 대조:
  - 모집단에 존재 → 검증 가능한 정답
  - 미존재(비상장/외국/신규/표기차이) → 구조적으로 못 뽑음(분모에서 제외)

실행: python scripts\\diag_groundtruth_coverage.py
"""
from __future__ import annotations
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from core.cooccur import _norm_name, name_to_sub_map


def main():
    ldir = Path("data/labels")
    peers = pd.read_csv(ldir / "peer_labels.csv", encoding="utf-8-sig", dtype=str)

    # 유니버스 회사명 집합(정규화)
    uni = pd.read_csv(config.LISTED_UNIVERSE_CSV, encoding=config.CSV_ENCODING, dtype=str)
    uni_names = {_norm_name(n) for n in uni[config.NAME_COL].dropna()}
    print(f"유니버스 회사: {len(uni_names):,}개")
    print(f"피어 라벨 엣지: {len(peers)}개 / 발행사 {peers['issuer'].nunique()}개\n")

    # 피어별 모집단 존재 여부
    peers = peers[peers["peer"].astype(str).str.strip().ne("")].copy()
    peers["in_uni"] = peers["peer"].map(lambda x: _norm_name(x) in uni_names)
    total = len(peers)
    in_uni = int(peers["in_uni"].sum())
    print(f"피어 엣지 중 모집단 존재: {in_uni}/{total} ({in_uni/total*100:.1f}%)")
    print(f"미존재(비상장/외국/표기차이): {total - in_uni}\n")

    # 발행사별 검증 가능 정답 수 분포
    by_issuer = peers.groupby("issuer")["in_uni"].agg(["sum", "count"])
    by_issuer.columns = ["available", "total"]
    verifiable = by_issuer[by_issuer["available"] >= 2]   # 정답 2개+ 있어야 recall 의미
    print(f"검증 가능 발행사(모집단 정답 2개+): {len(verifiable)}/{len(by_issuer)}")
    print(f"  정답 평균 {by_issuer['available'].mean():.1f}개 / "
          f"검증가능 발행사 정답 평균 {verifiable['available'].mean():.1f}개\n")

    # 미존재 피어 샘플(표기차이 vs 진짜 비상장 구분 참고)
    missing = peers[~peers["in_uni"]]["peer"].value_counts().head(20)
    print("모집단 미존재 피어 상위 20 (표기차이면 정규화 보강 여지):")
    for name, cnt in missing.items():
        print(f"  {name}  ×{cnt}")


if __name__ == "__main__":
    main()
