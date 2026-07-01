"""신고서 라벨 배치 조립 단위테스트 (네트워크 불필요)."""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.filing_labels import build_label_rows, summarize

results = []
def check(n, c, d=""): results.append((n, bool(c), d))

CACHE = {
    "01140837": {
        "company_name": "이지트로닉스", "corp_code": "01140837",
        "initial_rcept": "20211222000428", "status": "ok",
        "labels": {
            "issuer_ksic": None,
            "comparison_ksic": ["C26429", "C26410", "C28121"],
            "final_peers": [{"name": "이노와이어리스", "per": 46.12},
                            {"name": "우리산업", "per": 28.16}],
            "excluded": [{"name": "영화테크", "reason": "PER 137 아웃라이어"}],
        },
    },
    "01022902": {   # 기술특례 추정 — 피어 없음
        "company_name": "애드바이오텍", "corp_code": "01022902",
        "initial_rcept": "20211230000111", "status": "ok",
        "labels": {"issuer_ksic": None, "comparison_ksic": [],
                   "final_peers": [], "excluded": []},
    },
    "09999999": {   # 최초신고서 못 찾음
        "company_name": "실패사", "corp_code": "09999999",
        "initial_rcept": "", "status": "no_initial", "labels": {},
    },
}


def run():
    peer_rows, ksic_rows, excl_rows = build_label_rows(CACHE)
    check("피어행 2개(이지트로닉스만)", len(peer_rows) == 2, str(len(peer_rows)))
    check("피어 PER 보존", peer_rows[0]["per"] == 46.12)
    check("피어 발행사 태깅", peer_rows[0]["issuer"] == "이지트로닉스")
    check("KSIC행 3개", len(ksic_rows) == 3, str(len(ksic_rows)))
    check("KSIC 소분류 파생", ksic_rows[0]["ksic_sub"] == "264", str(ksic_rows[0]))
    check("탈락행 1개", len(excl_rows) == 1 and excl_rows[0]["peer"] == "영화테크")
    check("실패사 제외", all(r["issuer"] != "실패사" for r in peer_rows + ksic_rows))

    s = summarize(CACHE)
    check("요약 total=3", s["total"] == 3)
    check("요약 ok=2", s["ok"] == 2)
    check("요약 피어엣지=2", s["peer_edges"] == 2)
    check("요약 피어없음=1(애드바이오텍)", s["peers_empty"] == 1)
    check("요약 no_initial=1", s["no_initial"] == 1)


def main():
    run()
    p = sum(1 for _, ok, _ in results if ok)
    print("=" * 60)
    for n, ok, d in results:
        print(f"[{'PASS' if ok else 'FAIL'}] {n}" + ("" if ok or not d else f"   <- {d}"))
    print("=" * 60); print(f"{p}/{len(results)} passed")
    sys.exit(0 if p == len(results) else 1)


if __name__ == "__main__":
    main()
