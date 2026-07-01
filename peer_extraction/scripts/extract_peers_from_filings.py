"""신고서 배치 라벨 추출 — enriched.csv(200건) → 피어/KSIC/탈락 라벨 CSV.

흐름(건별):
  발행조건확정 rcept → corp_code로 최초 증권신고서 재탐색(find_initial_registration)
  → fetch_document(본문) → extract_filing_labels(Gemini) → 캐시
재개 가능: 캐시(corp_code별)에 있으면 스킵. 실패는 로그 남기고 계속.

실행 (peer_extraction 루트):
  python scripts\\extract_peers_from_filings.py --input "...\\peer_group_features_enriched_partial.csv" --limit 10
  python scripts\\extract_peers_from_filings.py --input "...csv"            # 전체
출력(data/labels): peer_labels.csv, comparison_ksic.csv, excluded.csv, extract_log.csv
"""
from __future__ import annotations
import argparse
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from core.dart_client import DartClient
from core.filing_labels import build_label_rows, summarize
from core.peer_extract import extract_filing_labels


def _load_cache() -> dict:
    p = config.FILING_LABELS_CACHE
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {}


def _save_cache(cache: dict) -> None:
    config.FILING_LABELS_CACHE.parent.mkdir(parents=True, exist_ok=True)
    config.FILING_LABELS_CACHE.write_text(
        json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")


def _write_outputs(cache: dict, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    peer_rows, ksic_rows, excl_rows = build_label_rows(cache)
    pd.DataFrame(peer_rows, columns=["issuer", "corp_code", "peer", "per", "initial_rcept"]) \
        .to_csv(out_dir / "peer_labels.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(ksic_rows, columns=["issuer", "corp_code", "ksic", "ksic_sub"]) \
        .to_csv(out_dir / "comparison_ksic.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(excl_rows, columns=["issuer", "corp_code", "peer", "reason"]) \
        .to_csv(out_dir / "excluded.csv", index=False, encoding="utf-8-sig")
    log = [{"corp_code": c, "company_name": r.get("company_name", ""),
            "initial_rcept": r.get("initial_rcept", ""), "status": r.get("status", ""),
            "error": r.get("error", ""),
            "n_peers": len((r.get("labels") or {}).get("final_peers", []) or [])}
           for c, r in cache.items()]
    pd.DataFrame(log).to_csv(out_dir / "extract_log.csv", index=False, encoding="utf-8-sig")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="enriched.csv 경로")
    ap.add_argument("--out-dir", default="data/labels")
    ap.add_argument("--limit", type=int, default=None, help="처리 건수 상한(파일럿용)")
    ap.add_argument("--refresh", action="store_true", help="캐시 무시하고 재추출")
    ap.add_argument("--flush-every", type=int, default=5)
    args = ap.parse_args()

    if not Path(args.input).exists():
        sys.exit(f"[중단] 입력 없음: {args.input}")
    df = pd.read_csv(args.input, encoding="utf-8-sig",
                     dtype={"corp_code": str, "rcept_no": str})
    df["corp_code"] = df["corp_code"].astype(str).str.strip()
    df.loc[df["corp_code"].isin(["", "nan", "None", "NaN"]), "corp_code"] = ""
    df.loc[df["corp_code"] != "", "corp_code"] = df.loc[df["corp_code"] != "", "corp_code"].str.zfill(8)
    if args.limit:
        df = df.head(args.limit)

    out_dir = Path(args.out_dir)
    cache = {} if args.refresh else _load_cache()
    dc = DartClient()

    # corp_code 결측 → 회사명으로 역매핑 복구
    missing = df[df["corp_code"] == ""]
    name2corp = {}
    if len(missing):
        print(f"corp_code 결측 {len(missing)}건 → 회사명 역매핑 시도...")
        name2corp = dc.corp_code_by_name()

    def resolve_corp(row) -> str:
        cc = row["corp_code"]
        if cc:
            return cc
        return name2corp.get(dc._norm_name(row.get("company_name", "")), "")

    # 처리 대상: 미캐시 또는 에러(재시도). ok/no_initial/no_corp_code는 스킵.
    def needs(key):
        if args.refresh or key not in cache:
            return True
        return str(cache[key].get("status", "")).startswith("error")

    todo = []
    for _, r in df.iterrows():
        corp = resolve_corp(r)
        key = corp or f"name:{r.get('company_name','')}"   # 충돌 방지(결측도 고유키)
        if needs(key):
            todo.append((key, corp, r))
    print(f"입력 {len(df)}건 / 처리 대상 {len(todo)}건 "
          f"(캐시 적중 {len(df) - len(todo)})")

    done = 0
    for key, corp, r in todo:
        name = r.get("company_name", "")
        rec = {"company_name": name, "corp_code": corp,
               "initial_rcept": "", "status": "", "error": "", "labels": {}}
        try:
            if not corp:
                rec["status"] = "no_corp_code"
            else:
                init = dc.find_initial_registration(corp)
                if not init:
                    rec["status"] = "no_initial"
                else:
                    rec["initial_rcept"] = init["rcept_no"]
                    text = dc.fetch_document(init["rcept_no"])
                    rec["labels"] = extract_filing_labels(text)
                    rec["status"] = "ok"
        except Exception as e:
            rec["status"] = f"error:{type(e).__name__}"
            rec["error"] = str(e)[:200]
        cache[key] = rec
        done += 1
        npeers = len((rec.get("labels") or {}).get("final_peers", []) or [])
        tail = f"  ({rec['error']})" if rec["error"] else ""
        print(f"  [{done}/{len(todo)}] {name[:16]:<16} {rec['status']:<14} 피어 {npeers}{tail}")
        if done % args.flush_every == 0:
            _save_cache(cache)
            _write_outputs(cache, out_dir)

    _save_cache(cache)
    _write_outputs(cache, out_dir)

    s = summarize(cache)
    print("\n" + "=" * 50)
    print(f"전체 {s['total']} / 성공 {s['ok']} / 피어엣지 {s['peer_edges']} / "
          f"피어없음 {s['peers_empty']} / 최초신고서없음 {s['no_initial']} / "
          f"corp없음 {s['no_corp_code']} / 오류 {s['error']}")
    print(f"출력: {out_dir}/  (peer_labels.csv, comparison_ksic.csv, excluded.csv, extract_log.csv)")
    if s["peers_empty"]:
        print(f"  ※ 피어없음 {s['peers_empty']}건 = 기술특례(PER 피어 미존재) 추정.")
    if s["error"] or s["no_corp_code"]:
        print(f"  ※ 오류·corp없음은 extract_log.csv의 error 컬럼 확인. 재실행 시 오류건 자동 재시도.")


if __name__ == "__main__":
    main()