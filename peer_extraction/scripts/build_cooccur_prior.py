"""과거 IPO 피어그룹 → 동시출현 prior 행렬 생성 (Stage1 Lane C 데이터화).

입력 CSV(엣지리스트): 한 행 = (issuer, peer) 한 쌍. 회사명 또는 KSIC 코드 모두 허용.
  issuer,peer
  나라스페이스,쎄트렉아이
  나라스페이스,AP위성
  ...
피어가 회사명이면 상장 유니버스 CSV로 KSIC 소분류 자동 해소.

실행 (peer_extraction 루트):
  python scripts/build_cooccur_prior.py --input data/peer_labels.csv --min-count 2

출력: data/cache/cooccur_prior.json  → pipeline이 자동 로드.
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from core.cooccur import (build_prior, edges_from_csv, name_to_sub_map,
                          resolve_sub, save_prior)


def main():
    ap = argparse.ArgumentParser(description="동시출현 prior 빌드")
    ap.add_argument("--input", required=True, help="엣지리스트 CSV (issuer,peer)")
    ap.add_argument("--issuer-col", default="issuer")
    ap.add_argument("--peer-col", default="peer")
    ap.add_argument("--min-count", type=int, default=2,
                    help="이 IPO 수 이상 등장한 (발행사소분류→피어소분류)만 유지 (노이즈 제거)")
    ap.add_argument("--no-normalize", action="store_true", help="정규화 끄고 원시 카운트 저장")
    args = ap.parse_args()

    if not Path(args.input).exists():
        sys.exit(f"[중단] 입력 CSV 없음: {args.input}")
    if not config.LISTED_UNIVERSE_CSV.exists():
        sys.exit(f"[중단] 유니버스 CSV 없음: {config.LISTED_UNIVERSE_CSV}")

    print(f"유니버스 로드 → 회사명→소분류 사전 구성...")
    name2sub = name_to_sub_map()
    print(f"  사전 {len(name2sub):,}개")

    ipo_peers = edges_from_csv(args.input, args.issuer_col, args.peer_col)
    print(f"입력: {len(ipo_peers)}개 IPO(발행사) / "
          f"{sum(len(p) for _, p in ipo_peers)}개 피어 엣지")

    # 미해소 진단(상위 몇 개)
    unresolved = []
    for iss, peers in ipo_peers:
        for tok in [iss, *peers]:
            if resolve_sub(tok, name2sub) is None and tok not in unresolved:
                unresolved.append(tok)
    if unresolved:
        print(f"  ⚠ 코드 미해소 {len(unresolved)}개 (회사명 표기 불일치 가능): "
              f"{', '.join(unresolved[:10])}{' ...' if len(unresolved) > 10 else ''}")

    prior = build_prior(ipo_peers, name2sub,
                        min_count=args.min_count, normalize=not args.no_normalize)
    meta = prior.get("_meta", {})
    out = save_prior(prior)

    print(f"\nprior 생성: 발행사 소분류 {meta.get('n_issuer_subs', 0)}개 "
          f"(min_count={args.min_count})")
    # 샘플 출력
    shown = 0
    for isub, row in prior.items():
        if isub == "_meta":
            continue
        top = list(row.items())[:5]
        print(f"  {isub} → " + ", ".join(f"{p}({w})" for p, w in top))
        shown += 1
        if shown >= 8:
            break
    print(f"\n저장: {out}")
    print("pipeline이 자동 로드 → Stage1 Lane C 활성. (run.py 재실행 시 반영)")


if __name__ == "__main__":
    main()
