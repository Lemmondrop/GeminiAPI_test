"""신고서 라벨 배치 추출 — 순수 로직(네트워크 분리, 테스트 가능).

캐시(dict: corp_code → 레코드)에서 3종 출력 행을 조립.
레코드 형태:
  {company_name, corp_code, initial_rcept, status,
   labels: {issuer_ksic, comparison_ksic[], final_peers[{name,per}], excluded[{name,reason}]}}
"""
from __future__ import annotations

from core.ksic import parse_ksic


def build_label_rows(cache: dict) -> tuple[list[dict], list[dict], list[dict]]:
    """캐시 → (peer_rows, ksic_rows, excluded_rows). 각자 CSV로 저장될 평면 행."""
    peer_rows, ksic_rows, excl_rows = [], [], []
    for corp_code, rec in cache.items():
        if rec.get("status") != "ok":
            continue
        issuer = rec.get("company_name", "")
        lab = rec.get("labels", {}) or {}
        for p in lab.get("final_peers", []) or []:
            peer_rows.append({
                "issuer": issuer, "corp_code": corp_code,
                "peer": p.get("name", ""), "per": p.get("per"),
                "initial_rcept": rec.get("initial_rcept", ""),
            })
        for code in lab.get("comparison_ksic", []) or []:
            sub = parse_ksic(code)[2] if code else ""
            ksic_rows.append({
                "issuer": issuer, "corp_code": corp_code,
                "ksic": code, "ksic_sub": sub,
            })
        for e in lab.get("excluded", []) or []:
            excl_rows.append({
                "issuer": issuer, "corp_code": corp_code,
                "peer": e.get("name", ""), "reason": e.get("reason", ""),
            })
    return peer_rows, ksic_rows, excl_rows


def summarize(cache: dict) -> dict:
    """추출 현황 요약."""
    n = len(cache)
    ok = sum(1 for r in cache.values() if r.get("status") == "ok")
    peers = sum(len((r.get("labels") or {}).get("final_peers", []) or [])
                for r in cache.values() if r.get("status") == "ok")
    empty = sum(1 for r in cache.values()
                if r.get("status") == "ok"
                and not (r.get("labels") or {}).get("final_peers"))
    no_init = sum(1 for r in cache.values() if r.get("status") == "no_initial")
    no_corp = sum(1 for r in cache.values() if r.get("status") == "no_corp_code")
    err = sum(1 for r in cache.values() if str(r.get("status", "")).startswith("error"))
    return {"total": n, "ok": ok, "peer_edges": peers,
            "peers_empty": empty, "no_initial": no_init,
            "no_corp_code": no_corp, "error": err}