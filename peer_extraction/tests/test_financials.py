"""재무 추출 로직 단위테스트 (DART 네트워크 불필요).

compute_metrics / _pick(CFS 우선) / _to_num / extract_audit 의 정확성을 합성 응답으로 검증.
실행: peer_extraction 루트에서  python tests/test_financials.py
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.financials import _to_num, compute_metrics, extract_audit, _pick

results: list[tuple[str, bool, str]] = []


def check(name: str, cond: bool, detail: str = ""):
    results.append((name, bool(cond), detail))


# ---- 합성 사업보고서(연간): CFS와 OFS 혼재, BS 행 포함 ----
ANNUAL = [
    {"fs_div": "CFS", "sj_div": "IS", "account_nm": "매출액", "thstrm_amount": "1,000", "frmtrm_amount": "900"},
    {"fs_div": "CFS", "sj_div": "IS", "account_nm": "영업이익", "thstrm_amount": "100", "frmtrm_amount": "80"},
    {"fs_div": "CFS", "sj_div": "IS", "account_nm": "당기순이익", "thstrm_amount": "70", "frmtrm_amount": "60"},
    {"fs_div": "OFS", "sj_div": "IS", "account_nm": "당기순이익", "thstrm_amount": "50", "frmtrm_amount": "40"},
    {"fs_div": "CFS", "sj_div": "BS", "account_nm": "자산총계", "thstrm_amount": "5,000", "frmtrm_amount": "4,800"},
]
# 1분기 보고서: thstrm=당기 YTD, frmtrm=전기 동기 YTD
QUARTER = [
    {"fs_div": "CFS", "sj_div": "IS", "account_nm": "영업이익", "thstrm_amount": "30", "frmtrm_amount": "20"},
    {"fs_div": "CFS", "sj_div": "IS", "account_nm": "당기순이익", "thstrm_amount": "25", "frmtrm_amount": "15"},
    {"fs_div": "CFS", "sj_div": "IS", "account_nm": "매출액", "thstrm_amount": "280", "frmtrm_amount": "250"},
]


def run():
    m = compute_metrics(ANNUAL, QUARTER)

    # CFS 우선 (별도 OFS 50이 아니라 연결 70)
    check("CFS 우선: net_income_annual=70", m["net_income_annual"] == 70, str(m["net_income_annual"]))
    check("FS 표기 = CFS", m["net_income_fs"] == "CFS", str(m["net_income_fs"]))

    # 온기
    check("op_income_annual=100", m["op_income_annual"] == 100)
    check("revenue_annual=1000", m["revenue_annual"] == 1000)

    # LTM = 온기 - 전기동기 + 당기동기
    check("op LTM = 100-20+30 = 110", m["op_income_ltm"] == 110, str(m["op_income_ltm"]))
    check("net LTM = 70-15+25 = 80", m["net_income_ltm"] == 80, str(m["net_income_ltm"]))
    check("rev LTM = 1000-250+280 = 1030", m["revenue_ltm"] == 1030, str(m["revenue_ltm"]))

    # 분기 없으면 LTM = 온기
    m2 = compute_metrics(ANNUAL, [])
    check("분기 없음 → LTM=온기", m2["net_income_ltm"] == 70 and m2["op_income_ltm"] == 100)

    # BS 행은 손익 지표로 안 잡힘 (자산총계가 어디에도 안 섞임)
    check("BS 무시: revenue_annual≠5000", m["revenue_annual"] != 5000)

    # OFS 폴백: 연결이 아예 없는 회사
    ofs_only = [
        {"fs_div": "OFS", "sj_div": "IS", "account_nm": "당기순이익", "thstrm_amount": "33", "frmtrm_amount": "30"},
    ]
    cur, prev, fs = _pick(ofs_only, ("당기순이익",))
    check("연결 없으면 OFS 폴백", cur == 33 and fs == "OFS", f"{cur},{fs}")

    # 숫자 파싱
    check("_to_num '(1,234)' = -1234", _to_num("(1,234)") == -1234)
    check("_to_num '△500' = -500", _to_num("△500") == -500)
    check("_to_num '1,000' = 1000", _to_num("1,000") == 1000)
    check("_to_num '' = None", _to_num("") is None)
    check("_to_num '-' = None", _to_num("-") is None)

    # 감사의견
    op, ad = extract_audit([{"adt_opinion": "적정", "adtor": "삼일회계법인"}])
    check("감사의견 적정 추출", op == "적정" and ad == "삼일회계법인")
    check("빈 감사응답 → (None,None)", extract_audit([]) == (None, None))


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
