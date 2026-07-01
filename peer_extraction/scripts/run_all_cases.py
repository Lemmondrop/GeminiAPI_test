"""
scripts/run_all_cases.py
────────────────────────────────────────────────────────────
poc_results/*.json (단, _summary.json 제외) 전체를 순회하면서,
ir_signals.json에 등록된 IR 신호로 run_full_valuation.run_one()을
호출해 트랙 자동판별 + 가치산출까지 한 번에 돌리고 표로 요약한다.

ir_signals.json에 케이스가 없으면(=아직 IR 신호를 안 채워넣은 신규
케이스) 빈 신호로 실행한다 — stage_router가 review_needed=True로
안전하게 처리하고, 이 배치 스크립트가 "IR 신호 미입력" 사유를 별도로
표시해준다. 트랙을 사람이 고르는 게 아니라, "아직 원재료가 없다"는
걸 구분해서 보여주는 것이 목적.

실행 위치: peer_extraction/scripts/
사용법:
  python run_all_cases.py
  python run_all_cases.py --quiet   # 케이스별 상세 로그 생략, 요약표만
────────────────────────────────────────────────────────────
"""
import sys
import json
import argparse
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
ROOT_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(ROOT_DIR))

from run_full_valuation import run_one

POC_RESULTS_DIR = ROOT_DIR / "poc_results"
IR_SIGNALS_PATH = ROOT_DIR / "ir_signals.json"


def load_ir_signals() -> dict:
    if not IR_SIGNALS_PATH.exists():
        print(f"[경고] {IR_SIGNALS_PATH} 없음 — 모든 케이스를 빈 신호로 실행합니다.")
        return {}
    data = json.loads(IR_SIGNALS_PATH.read_text(encoding="utf-8"))
    data.pop("_설명", None)
    return data


def list_cases() -> list:
    return sorted(
        p.stem for p in POC_RESULTS_DIR.glob("*.json")
        if p.stem != "_summary"
    )


def format_value(result: dict) -> str:
    # cv_r2가 음수에 가까운(=예측력 사실상 없음) 프로토타입 모델이라,
    # ML 조정값(공모시총)을 메인으로 내세우면 오히려 오도할 수 있다.
    # 근거가 분명한 평가시총(피어 PER 기반)을 메인으로, ML 조정은 참고용 부제로만.
    base = None
    if result["equity_value_억원"] is not None:
        base = f"{result['equity_value_억원']:,.1f}억원 (평가시총)"
    elif result["value_range_억원"] is not None:
        lo, hi = result["value_range_억원"]
        base = f"{lo:,.0f}억 ~ {hi:,.0f}억원"

    if base is None:
        return f"산출불가 ({result['warning']})" if result["warning"] else "산출불가"

    offering = result.get("offering_market_cap_억원")
    if offering is not None:
        pct = result.get("offering_discount_pct")
        return f"{base}  |  참고(낮은신뢰도) 공모조정: {offering:,.1f}억원 (할인 {pct:.1f}%)"
    return base


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quiet", action="store_true", help="케이스별 상세 로그 생략, 요약표만 출력")
    args = ap.parse_args()

    ir_signals = load_ir_signals()
    cases = list_cases()

    if not cases:
        sys.exit(f"[ERROR] {POC_RESULTS_DIR} 에 케이스 파일이 없습니다.")

    print(f"총 {len(cases)}개 케이스 실행: {cases}\n")

    summary_rows = []
    for case in cases:
        signals = ir_signals.get(case)
        missing_signals = signals is None
        signals = signals or {}

        if not args.quiet:
            print(f"\n{'#'*65}\n# {case}" + ("  [⚠ IR 신호 미입력 — 빈 신호로 실행]" if missing_signals else "") + f"\n{'#'*65}")

        result = run_one(case, signals, verbose=not args.quiet)
        sd = result["stage_decision"]

        summary_rows.append({
            "case": case,
            "track": sd.track,
            "sub_track": sd.sub_track or "-",
            "confidence": sd.confidence,
            "review_needed": sd.review_needed,
            "missing_signals": missing_signals,
            "value": format_value(result),
        })

    # ── 요약표 ──
    print(f"\n\n{'='*100}")
    print("전체 요약")
    print(f"{'='*100}")
    header = f"{'케이스':<28} {'트랙':<18} {'서브트랙':<14} {'신뢰도':<8} {'확인필요':<10} {'가치':<28}"
    print(header)
    print("-" * len(header))
    for row in summary_rows:
        flag = "⚠ IR미입력" if row["missing_signals"] else ("⚠ 확인필요" if row["review_needed"] else "")
        print(f"{row['case']:<28} {row['track']:<18} {row['sub_track']:<14} "
              f"{row['confidence']:<8.2f} {flag:<10} {row['value']:<28}")

    n_missing = sum(1 for r in summary_rows if r["missing_signals"])
    if n_missing:
        print(f"\n⚠ {n_missing}개 케이스는 ir_signals.json에 IR 신호가 없어 빈 값으로 실행됐습니다.")
        print(f"  해당 케이스의 설립연도·최근 라운드·IPO 언급 여부 등을 ir_signals.json에 추가해주세요:")
        for r in summary_rows:
            if r["missing_signals"]:
                print(f"    - {r['case']}")


if __name__ == "__main__":
    main()