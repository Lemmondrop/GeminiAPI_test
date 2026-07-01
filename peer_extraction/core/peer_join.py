"""core/peer_join.py — Stage1~3 피어 후보와 재무 데이터 조인.

배경:
  Stage1~3을 거친 피어 후보(stage3 DataFrame: 회사명·종목코드·tier·
  sim_score·rationale)와, 재무 수집 결과(config.UNIVERSE_FIN_CSV:
  market_cap·revenue·net_income 등)는 지금 별도 테이블이다.

  Module 2(밸류에이션, valuation.py/valuation_psr.py)가 요구하는
  PeerItem 형태(market_cap_억원·revenue_억원·net_income_억원·per)를
  만들려면 둘을 종목코드(config.STOCK_COL) 기준으로 조인해야 한다.

  본 모듈은 그 조인 한 단계만 담당한다 — 계산 로직(PER/PSR 산출)은
  core/valuation.py·core/valuation_psr.py에 둔다(관심사 분리).

단위 처리:
  universe_financials.csv는 원(₩) 단위로 저장되어 있음(실측 확인:
  엔비알모션 market_cap=92,703,810,000원). 본 모듈은 조인 시점에
  즉시 억원으로 변환해 반환하므로, 호출측은 단위 변환을 신경 쓸
  필요가 없다.

재무 우선순위:
  매출·순이익 모두 LTM(최근 4분기 합산) 우선, 결측 시 연간(Annual)
  값으로 폴백 — config.py의 ANNUAL_YEAR/LTM_QUARTER 설계 의도와 일치.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

try:
    import config
    _DEFAULT_FIN_CSV = config.UNIVERSE_FIN_CSV
    _DEFAULT_STOCK_COL = config.STOCK_COL
except ImportError:
    # config.py 미발견 시(단독 테스트 등) 안전한 기본값으로 폴백
    _DEFAULT_FIN_CSV = None
    _DEFAULT_STOCK_COL = "종목코드"


def _won_to_억원(value) -> float | None:
    """원(₩) 값을 억원으로 변환. None/NaN/결측이면 None 반환."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return float(value) / 1e8


def _coalesce(*values):
    """첫 번째로 유효한(None/NaN 아닌) 값을 반환. 전부 무효면 None."""
    for v in values:
        if v is None:
            continue
        try:
            if pd.isna(v):
                continue
        except (TypeError, ValueError):
            pass
        return v
    return None


def load_financials_lookup(
    fin_csv_path: str | Path | None = None,
    stock_col: str | None = None,
) -> dict:
    """universe_financials.csv를 1회 로드해 {종목코드: row_dict} 딕셔너리로 반환.

    여러 피어를 반복 조인할 때 CSV를 매번 다시 읽지 않도록, 호출측에서
    이 함수로 한 번 로드해 enrich_peers_with_financials()의
    fin_lookup 인자로 재사용하는 것을 권장한다.
    """
    stock_col = stock_col or _DEFAULT_STOCK_COL
    fin_path = Path(fin_csv_path) if fin_csv_path else _DEFAULT_FIN_CSV
    if fin_path is None:
        raise ValueError(
            "fin_csv_path를 지정하거나 config.UNIVERSE_FIN_CSV가 "
            "설정되어 있어야 합니다 (config.py import 실패 시 "
            "fin_csv_path를 명시적으로 전달하세요)."
        )
    fin_path = Path(fin_path)
    if not fin_path.exists():
        raise FileNotFoundError(f"재무 데이터 파일을 찾을 수 없습니다: {fin_path}")

    fin_df = pd.read_csv(fin_path, dtype={stock_col: str})
    fin_df[stock_col] = fin_df[stock_col].astype(str).str.strip().str.zfill(6)

    # ── 중복 종목코드 처리 ───────────────────────────────────
    # company_cord_prototype.csv에서 이미 확인된 것처럼 원본 CSV에
    # 중복 행이 섞여 있을 수 있음(같은 종목코드가 여러 번 등장).
    # has_financials=True인 행을 우선 보존하고, 그래도 중복이면
    # 마지막 행을 채택(가장 최근에 수집/갱신됐을 가능성이 높음).
    dup_count = int(fin_df[stock_col].duplicated().sum())
    if dup_count > 0:
        print(f"[peer_join] 경고: {fin_path.name}에 중복 종목코드 "
              f"{dup_count}건 발견 — has_financials 우선, 그 외 "
              "마지막 행으로 정리합니다.")
        if "has_financials" in fin_df.columns:
            fin_df = fin_df.sort_values(
                "has_financials", ascending=True, kind="stable")
        fin_df = fin_df.drop_duplicates(subset=stock_col, keep="last")

    return fin_df.set_index(stock_col).to_dict("index")


def _normalize_name(name: str) -> str:
    """회사명 느슨한 비교용 정규화 — 공백·법인격 표기 차이 흡수.

    "주식회사 케이더블유씨" / "㈜케이더블유씨" / "케이더블유씨(주)" 등
    표기가 갈리는 경우를 같은 회사로 매칭하기 위함. 종목코드가 없을 때만
    쓰는 차선책이므로 완벽한 정규화보다 흔한 패턴만 제거한다.
    """
    if not name:
        return ""
    n = name.strip()
    for tok in ("주식회사", "㈜", "(주)", "（주）"):
        n = n.replace(tok, "")
    return n.replace(" ", "").strip()


def build_name_lookup(fin_lookup: dict, name_col: str = "회사명") -> dict:
    """종목코드 기준 fin_lookup으로부터 {정규화된 회사명: row} 보조 인덱스 생성.

    enrich_peers_with_financials()가 종목코드로 못 찾을 때 폴백으로 사용.
    """
    name_idx = {}
    for row in fin_lookup.values():
        nm = row.get(name_col)
        if nm:
            name_idx[_normalize_name(nm)] = row
    return name_idx


def enrich_peers_with_financials(
    peers: list[dict],
    fin_lookup: dict | None = None,
    fin_csv_path: str | Path | None = None,
    stock_col: str | None = None,
    name_col: str = "회사명",
    name_key_in_peer: str = "name",
) -> list[dict]:
    """피어 후보 리스트에 재무 데이터(매출·순이익·시총·PER·PSR)를 조인한다.

    Args:
        peers: Stage1~3 결과. 각 dict는 stock_col 키(있으면 우선) 또는
               name_key_in_peer 키(폴백)를 가져야 한다.
               예) stage3_df.to_dict("records") — 종목코드 있음(우선)
               예) poc_runner.py 결과 final_peers — 종목코드 없음,
                   "name" 키만 있음 → 자동으로 회사명 폴백 조인.
        fin_lookup: load_financials_lookup() 반환값을 미리 넘기면
               CSV를 다시 읽지 않음(여러 케이스 반복 처리 시 효율적).
               None이면 내부에서 1회 로드.
        fin_csv_path, stock_col: load_financials_lookup()과 동일.
        name_col: fin CSV 내 회사명 컬럼명 (기본 "회사명").
        name_key_in_peer: peer dict 내 회사명 키 (기본 "name" —
               poc_runner.py 출력 스키마 기준. Stage3 dataframe이면
               "회사명"으로 바꿔서 호출).

    Returns:
        각 피어 dict에 다음 필드가 추가된 새 리스트(원본 미변형):
          - has_financials: bool, 재무 조인 성공 여부
          - join_method: "stock_code" | "name" | None — 어떤 방식으로
            매칭됐는지(감사 추적용. name 매칭은 동명이인 위험이 있어
            결과 신뢰도가 stock_code보다 낮음을 인지하고 사용할 것)
          - market_cap_억원, revenue_억원, net_income_억원: float|None
          - per: float|None  (순이익 흑자일 때만 산출, 적자/결측이면 None)
          - psr: float|None  (매출 양수일 때만 산출)
          - audit_opinion: str|None
    """
    stock_col = stock_col or _DEFAULT_STOCK_COL
    if fin_lookup is None:
        fin_lookup = load_financials_lookup(fin_csv_path, stock_col)
    name_lookup = build_name_lookup(fin_lookup, name_col)

    enriched: list[dict] = []
    matched = 0
    matched_by_name = 0

    for p in peers:
        p = dict(p)  # 원본 dict 변형 방지
        row = None
        join_method = None

        sc = str(p.get(stock_col, "")).strip()
        if sc:
            # 종목코드는 항상 순수 숫자가 아님(스팩 등은 영숫자 혼합,
            # 예: "0004V0") — 형식 제약 없이 그대로 조회 시도.
            row = fin_lookup.get(sc.zfill(6)) or fin_lookup.get(sc)
            if row is not None:
                join_method = "stock_code"

        if row is None:
            nm = p.get(name_key_in_peer) or p.get(name_col)
            if nm:
                row = name_lookup.get(_normalize_name(nm))
                if row is not None:
                    join_method = "name"
                    matched_by_name += 1

        if row is None:
            p.update({
                "has_financials": False,
                "join_method": None,
                "market_cap_억원": None,
                "revenue_억원": None,
                "net_income_억원": None,
                "per": None,
                "psr": None,
                "audit_opinion": None,
            })
            enriched.append(p)
            continue

        market_cap_억 = _won_to_억원(row.get("market_cap"))

        revenue_won = _coalesce(row.get("revenue_ltm"), row.get("revenue_annual"))
        revenue_억 = _won_to_억원(revenue_won)

        net_income_won = _coalesce(row.get("net_income_ltm"), row.get("net_income_annual"))
        net_income_억 = _won_to_억원(net_income_won)

        per = None
        if (net_income_억 is not None and net_income_억 > 0
                and market_cap_억 is not None):
            per = round(market_cap_억 / net_income_억, 2)

        psr = None
        if (revenue_억 is not None and revenue_억 > 0
                and market_cap_억 is not None):
            psr = round(market_cap_억 / revenue_억, 3)

        p.update({
            "has_financials": bool(_coalesce(row.get("has_financials"), True)),
            "join_method": join_method,
            "market_cap_억원": market_cap_억,
            "revenue_억원": revenue_억,
            "net_income_억원": net_income_억,
            "per": per,
            "psr": psr,
            "audit_opinion": row.get("audit_opinion"),
        })
        matched += 1
        enriched.append(p)

    if peers:
        print(f"[peer_join] 재무 조인: {matched}/{len(peers)}개 성공 "
              f"({matched/len(peers)*100:.0f}%)"
              + (f"  [회사명 폴백 매칭 {matched_by_name}건 — 동명이인 위험 "
                 "있으니 결과 검토 권장]" if matched_by_name else ""))
    else:
        print("[peer_join] 입력 피어 0개 — 조인 생략")

    return enriched


# ══════════════════════════════════════════════════════════
# 데모 — 실제 universe_financials.csv 스키마와 동일한 가상 데이터로 검증
# ══════════════════════════════════════════════════════════
if __name__ == "__main__":
    import io

    print("=== peer_join 데모 (실제 CSV 스키마 기준 가상 데이터) ===\n")

    # 실제로 확인된 universe_financials.csv 행 형태 재현
    demo_csv = """종목코드,회사명,결산월,corp_code,market_cap,revenue_annual,revenue_ltm,revenue_fs,op_income_annual,op_income_ltm,op_income_fs,net_income_annual,net_income_ltm,net_income_fs,audit_opinion,auditor,has_financials
0004V0,엔비알모션,12월,924931,9.270381e+10,6.309446e+10,6.250655e+10,CFS,-3.928308e+09,-5.158675e+09,CFS,-1.224866e+10,-1.361167e+10,CFS,적정의견,회계법인 더올,True
493330,지에프아이,12월,1450938,8.642414e+10,4.495571e+10,5.824928e+10,CFS,6.754673e+09,8.277231e+09,CFS,-3.583187e+09,-1.995718e+09,CFS,적정의견,대성회계법인,True
012210,삼미금속,12월,125664,1.890759e+11,7.379693e+10,7.438604e+10,OFS,5.812312e+09,6.429920e+09,OFS,1.690159e+09,2.656519e+09,OFS,적정의견,한미회계법인,True
"""
    demo_fin_df = pd.read_csv(io.StringIO(demo_csv), dtype={"종목코드": str})
    demo_fin_df["종목코드"] = demo_fin_df["종목코드"].str.zfill(6)
    fin_lookup_demo = demo_fin_df.set_index("종목코드").to_dict("index")

    # Stage3 결과 형태로 가상 피어 입력 (이름·종목코드·tier·sim_score만 보유)
    demo_peers = [
        {"종목코드": "0004V0", "회사명": "엔비알모션", "tier": "중간", "sim_score": 70},
        {"종목코드": "493330", "회사명": "지에프아이", "tier": "중간", "sim_score": 60},
        {"종목코드": "012210", "회사명": "삼미금속",   "tier": "높음", "sim_score": 85},
        {"종목코드": "999999", "회사명": "조인실패예시", "tier": "낮음", "sim_score": 40},
    ]

    result = enrich_peers_with_financials(demo_peers, fin_lookup=fin_lookup_demo)

    print()
    for p in result:
        print(f"  {p['회사명']:10s} has_fin={p['has_financials']!s:5s} "
              f"시총={p['market_cap_억원']!s:>8s}억 "
              f"매출={p['revenue_억원']!s:>8s}억 "
              f"순이익={p['net_income_억원']!s:>8s}억 "
              f"PER={p['per']!s:>6s} PSR={p['psr']!s:>6s}")

    print("\n[검증 포인트]")
    print("  엔비알모션: 순이익 적자(-136.1억) → PER=None, 그러나 매출 있어 PSR 산출됨")
    print("  지에프아이: 순이익 적자(-20.0억)  → PER=None, PSR은 산출됨")
    print("  삼미금속  : 순이익 흑자(26.6억)   → PER·PSR 둘 다 산출됨")
    print("  조인실패예시: 재무 미존재 종목코드 → has_financials=False, 전부 None")