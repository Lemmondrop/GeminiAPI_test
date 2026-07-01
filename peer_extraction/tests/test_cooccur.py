"""동시출현 prior 빌더 단위테스트 (네트워크 불필요).

회사명/코드 해소 · IPO단위 카운트 · min_count 임계 · 자기소분류 제외 · 정규화 · Stage1 연동.
실행: peer_extraction 루트에서  python tests\\test_cooccur.py
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.cooccur import build_prior, resolve_sub

results: list[tuple[str, bool, str]] = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))


# 가짜 회사명→소분류 사전 (유니버스 CSV 대체)
N2S = {
    "쎄트렉아이": "263", "ap위성": "263", "인텔리안테크": "263",   # 위성/통신장비
    "한화시스템": "264", "lig넥스원": "264",                        # 방산전자
    "알루코": "242",                                               # 알루미늄
}


def run():
    # ----- resolve_sub: 회사명 / 코드 / 숫자코드 -----
    check("회사명 해소", resolve_sub("쎄트렉아이", N2S) == "263")
    check("회사명 정규화(주식회사)", resolve_sub("주식회사 AP위성", N2S) == "263")
    check("코드 직접파싱 C00263", resolve_sub("C00263", N2S) == "263")
    check("숫자코드 263", resolve_sub("263", N2S) == "263")
    check("미해소 None", resolve_sub("없는회사", N2S) is None)

    # ----- build_prior: IPO단위 카운트 + 자기소분류 제외 -----
    ipos = [
        ("쎄트렉아이", ["AP위성", "인텔리안테크", "한화시스템"]),   # 발행사 263 → 피어 263,263,264
        ("AP위성", ["쎄트렉아이", "한화시스템"]),                   # 263 → 263, 264
        ("인텔리안테크", ["한화시스템", "LIG넥스원"]),              # 263 → 264, 264
    ]
    # 263 발행사 기준: 264가 3개 IPO 전부 등장(피어 264), 263은 자기소분류라 제외
    prior = build_prior(ipos, N2S, min_count=2, normalize=False)
    check("자기소분류(263) 제외", "263" not in prior.get("263", {}), str(prior.get("263")))
    check("264 링크 카운트=3", prior.get("263", {}).get("264") == 3.0, str(prior.get("263")))

    # min_count 임계: min_count=3이면 3회 미만 링크 제거
    p3 = build_prior(ipos, N2S, min_count=3, normalize=False)
    check("min_count=3 → 264 유지(3회)", p3.get("263", {}).get("264") == 3.0)

    # 단발(1회) 링크는 min_count=2에서 제거되는지: 별도 케이스
    single = [("쎄트렉아이", ["알루코"])]                          # 263→242 단 1회
    p1 = build_prior(single, N2S, min_count=2, normalize=False)
    check("단발 링크 제거(min_count=2)", "263" not in p1 or "242" not in p1.get("263", {}),
          str(p1.get("263")))

    # 정규화: 행 합 = 1
    pn = build_prior(ipos, N2S, min_count=1, normalize=True)
    row = pn.get("263", {})
    check("정규화 행합≈1", abs(sum(row.values()) - 1.0) < 1e-6, str(row))

    # _meta 동봉
    check("_meta 포함", "_meta" in prior and prior["_meta"]["n_ipos"] == 3)

    # ----- Stage1 연동: load_prior 형태가 Lane C 기대형({sub:{sub:w}}) -----
    from core.cooccur import save_prior, load_prior
    import tempfile, os
    tmp = Path(tempfile.gettempdir()) / "_test_prior.json"
    save_prior(pn, tmp)
    loaded = load_prior(tmp)
    check("load_prior _meta 제외", "_meta" not in loaded and "263" in loaded)
    # Lane C 사용방식 재현: prior.get(s_sub, {}).keys()
    prior_subs = set(loaded.get("263", {}).keys())
    check("Lane C 키집합 사용가능", "264" in prior_subs, str(prior_subs))
    os.remove(tmp)


def main():
    run()
    passed = sum(1 for _, ok, _ in results if ok)
    print("=" * 60)
    for name, ok, detail in results:
        mark = "PASS" if ok else "FAIL"
        line = f"[{mark}] {name}"
        if not ok and detail:
            line += f"   <- {detail}"
        print(line)
    print("=" * 60)
    print(f"{passed}/{len(results)} passed")
    sys.exit(0 if passed == len(results) else 1)


if __name__ == "__main__":
    main()
