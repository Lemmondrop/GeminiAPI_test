"""발행사 KSIC 수집 — peer_labels.csv의 발행사 corp_code → DART 기업개황 induty_code → 소분류.

발행사는 신생 상장사라 유니버스 사전에 없음 → DART company.json으로 표준산업분류 직접 조회.
재개 가능(corp_code별 캐시).

실행: python scripts\\build_issuer_ksic.py
출력: data/labels/issuer_ksic.csv  (issuer, corp_code, induty_code, issuer_sub)
"""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from core.dart_client import DartClient
from core.ksic import parse_ksic

CACHE = config.CACHE_DIR / "issuer_ksic_cache.json"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--labels-dir", default="data/labels")
    ap.add_argument("--refresh", action="store_true")
    a = ap.parse_args()

    ldir = Path(a.labels_dir)
    peer_csv = ldir / "peer_labels.csv"
    if not peer_csv.exists():
        sys.exit(f"[중단] {peer_csv} 없음 — 먼저 extract_peers_from_filings 실행")

    df = pd.read_csv(peer_csv, encoding="utf-8-sig", dtype={"corp_code": str})
    # comparison_ksic의 발행사도 포함(피어 0이어도 발행사 KSIC는 필요)
    comp_csv = ldir / "comparison_ksic.csv"
    if comp_csv.exists():
        cdf = pd.read_csv(comp_csv, encoding="utf-8-sig", dtype={"corp_code": str})
        df = pd.concat([df[["issuer", "corp_code"]], cdf[["issuer", "corp_code"]]],
                       ignore_index=True)
    pairs = df.dropna(subset=["corp_code"]).drop_duplicates("corp_code")
    pairs = pairs[pairs["corp_code"].astype(str).str.strip().ne("")]

    cache = {} if a.refresh else (json.loads(CACHE.read_text(encoding="utf-8"))
                                  if CACHE.exists() else {})
    dc = DartClient()
    todo = [(r["corp_code"], r["issuer"]) for _, r in pairs.iterrows()
            if a.refresh or r["corp_code"] not in cache]
    print(f"발행사 {len(pairs)}개 / 조회 대상 {len(todo)}개 (캐시 {len(pairs) - len(todo)})")

    for i, (corp, name) in enumerate(todo, 1):
        rec = {"issuer": name, "corp_code": corp, "induty_code": "", "issuer_sub": ""}
        try:
            prof = dc.company_profile(corp)
            ind = str(prof.get("induty_code") or "").strip()
            rec["induty_code"] = ind
            rec["issuer_sub"] = parse_ksic(ind)[2] if ind else ""
        except Exception as e:
            rec["induty_code"] = f"error:{type(e).__name__}"
        cache[corp] = rec
        if i % 20 == 0 or i == len(todo):
            CACHE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
            print(f"  {i}/{len(todo)}")

    CACHE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    out = pd.DataFrame(cache.values())
    out_path = ldir / "issuer_ksic.csv"
    out.to_csv(out_path, index=False, encoding="utf-8-sig")
    ok = (out["issuer_sub"].astype(str).str.len() == 3).sum()
    print(f"\n저장: {out_path}")
    print(f"발행사 소분류 확보: {ok}/{len(out)}")


if __name__ == "__main__":
    main()
