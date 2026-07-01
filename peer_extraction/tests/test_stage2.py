"""Stage 2 재무 필터 단위테스트 (네트워크 불필요).

트랙별 조건(general/tech_special) · 감사의견 '적정의견' 매칭 · 결산월 · 재무결측 탈락 · drop사유 검증.
실행: peer_extraction 루트에서  python tests/test_stage2.py
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from stages.stage2_financial import apply_stage2, survivors, _audit_ok

results: list[tuple[str, bool, str]] = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))


# 모집단 (stage1 출력 형태; 종목코드 포함)
POOL = pd.DataFrame({
    "회사명": ["흑자사", "적자LTM사", "온기적자사", "결산3월사", "한정의견사", "재무없음사"],
    "종목코드": ["000001", "000002", "000003", "000004", "000005", "000006"],
    "산업분류코드": ["C00303"] * 6,
})

# 재무 (universe_financials 형태)
FIN = pd.DataFrame({
    "종목코드": ["000001", "000002", "000003", "000004", "000005"],  # 000006은 재무 없음
    "결산월": ["12월", "12월", "12월", "3월", "12월"],   # 실데이터 포맷 (문자열 '12월')
    "op_income_annual": [100, 50, -30, 100, 100],
    "op_income_ltm": [110, 40, 20, 110, 110],
    "net_income_annual": [70, 30, -50, 70, 70],
    "net_income_ltm": [80, -10, 15, 80, 80],     # 적자LTM사 net_ltm<0; 온기적자사는 LTM 흑자(15)
    "audit_opinion": ["적정의견", "적정의견", "적정의견", "적정의견", "한정의견"],
    "has_financials": [True, True, True, True, True],
})


def run():
    # ---- general: net_income_annual>0 (연간 흑자, 신고서 2차 기준) ----
    g = apply_stage2(POOL, FIN, "general")
    gpass = set(survivors(g)["회사명"])
    check("general: 흑자사 통과", "흑자사" in gpass, str(gpass))
    check("general: 적자LTM사 통과(연간 흑자면 OK)", "적자LTM사" in gpass,
          "general은 net_income_annual을 보므로 LTM 적자 무관")
    check("general: 온기적자사 탈락(연간 순익<0)", "온기적자사" not in gpass)
    check("general: 결산3월사 탈락", "결산3월사" not in gpass)
    check("general: 한정의견사 탈락", "한정의견사" not in gpass)
    check("general: 재무없음사 탈락", "재무없음사" not in gpass)

    # drop 사유 확인
    rmap = dict(zip(g["회사명"], g["drop_reason"]))
    check("사유: 온기적자사 = 미충족:net_income_annual>0", "net_income_annual" in rmap["온기적자사"], rmap["온기적자사"])
    check("사유: 결산3월사 = 결산월≠12", "결산월" in rmap["결산3월사"], rmap["결산3월사"])
    check("사유: 한정의견사 = 감사의견부적정", "감사" in rmap["한정의견사"], rmap["한정의견사"])
    check("사유: 재무없음사 = 재무없음", rmap["재무없음사"] == "재무없음", rmap["재무없음사"])

    # ---- tech_special: 영업+순이익 온기·LTM 모두 양수 ----
    t = apply_stage2(POOL, FIN, "tech_special")
    tpass = set(survivors(t)["회사명"])
    check("tech_special: 흑자사 통과(전부 양수)", "흑자사" in tpass, str(tpass))
    check("tech_special: 온기적자사 탈락(온기 영업/순익 음수)", "온기적자사" not in tpass)
    check("tech_special: 적자LTM사 탈락(net_ltm<0)", "적자LTM사" not in tpass)

    # general과 tech_special 차이: 적자LTM사(연간 흑자→general통과, LTM 음수→tech탈락)
    check("트랙 차이: 적자LTM사 general통과 & tech탈락",
          "적자LTM사" in gpass and "적자LTM사" not in tpass)

    # 감사의견 매칭
    check("_audit_ok('적정의견')=True", _audit_ok("적정의견") is True)
    check("_audit_ok('적정')=True", _audit_ok("적정") is True)
    check("_audit_ok('한정의견')=False", _audit_ok("한정의견") is False)
    check("_audit_ok('의견거절')=False", _audit_ok("의견거절") is False)
    check("_audit_ok(None)=False", _audit_ok(None) is False)


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