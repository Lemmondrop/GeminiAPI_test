"""
scripts/update_ir_signals.py
────────────────────────────────────────────────────────────
OCR로 이미 추출된 IR .md 파일을 읽어 core.ir_parser.IrExtractor로 구조화
필드를 뽑고, ir_signals.json의 해당 케이스 항목을 자동으로 채우거나
갱신한다. run_all_cases.py가 그 다음에 바로 이 값을 쓸 수 있다.

--case는 poc_results/{case}.json 파일명과 반드시 일치해야 한다(예:
"주식회사 케이더블유씨 (KWC)"). IR 문서 안의 회사명 표기("(주)케이더블유씨"
등)와는 다를 수 있으므로 자동 매칭하지 않고 명시적으로 받는다 — 잘못된
케이스에 잘못 병합되는 사고를 방지.

사용법:
  python scripts\\update_ir_signals.py --case "주식회사 케이더블유씨 (KWC)" \\
      --pdf-path "C:\\...\\(주)케이더블유씨 IR 자료(20250905).pdf"

  PDF 업로드가 실패할 때를 대비해 .md 폴백도 같이 줄 수 있습니다:
  python scripts\\update_ir_signals.py --case "..." \\
      --pdf-path "...\\IR.pdf" --md-path "..._전체추출결과.md"

  여러 필드를 자동추출 후에도 수동으로 덮어쓰고 싶으면 --override 사용:
  python scripts\\update_ir_signals.py --case "..." --pdf-path "..." \\
      --override ipo_target_year=2028
────────────────────────────────────────────────────────────
"""
import sys
import json
import argparse
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
ROOT_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(ROOT_DIR))

from core.ir_parser import IrExtractor

IR_SIGNALS_PATH = ROOT_DIR / "ir_signals.json"

# ir_signals.json에 실제로 쓰는 필드만 골라서 저장 — extractor 출력엔
# company_name/_extraction_notes처럼 감사용 필드도 섞여 있으므로 분리.
SIGNAL_FIELDS = ("founding_year", "latest_round", "ipo_mentioned", "ipo_target_year",
                 "net_income", "future_net_income", "ni_year", "revenue",
                 "segment", "yoy_growth")


def load_ir_signals() -> dict:
    if not IR_SIGNALS_PATH.exists():
        return {"_설명": "각 키는 poc_results/{키}.json 파일명과 일치해야 합니다."}
    return json.loads(IR_SIGNALS_PATH.read_text(encoding="utf-8"))


def save_ir_signals(data: dict):
    IR_SIGNALS_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def print_review(case: str, extracted: dict):
    """추출 결과를 사람이 바로 훑어볼 수 있게 출력 — null인 필드를 눈에
    띄게 표시해서, LLM이 못 찾은 건지 원래 문서에 없는 건지 확인을
    유도한다(추측해서 채우지 않는 게 원칙이므로 null 자체는 정상일 수 있음)."""
    print(f"\n=== 추출 결과: {case} ===")
    print(f"  문서상 회사명: {extracted.get('company_name') or '(추출 안 됨)'}")
    for f in SIGNAL_FIELDS:
        v = extracted.get(f)
        flag = "  ⚠ null" if v is None else ""
        print(f"  {f:<18}: {v}{flag}")
    notes = extracted.get("_extraction_notes")
    if notes:
        print(f"  메모: {notes}")
    print(f"\n  ※ null 필드는 문서에 없어서 안 채운 것일 수도, LLM이 못 찾은 것일 "
          f"수도 있습니다 — 위 IR 원본과 대조해서 확인해주세요.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", required=True,
                    help="poc_results/{case}.json 파일명과 정확히 일치해야 함")
    ap.add_argument("--pdf-path", default=None, help="IR PDF 원본 경로 (우선 사용)")
    ap.add_argument("--md-path", default=None,
                    help="OCR 추출된 .md 경로 (PDF 실패 시 폴백, 또는 PDF 없을 때 단독 사용)")
    ap.add_argument("--override", action="append", default=[],
                    help="key=value 형태로 자동추출 결과를 덮어씀 (여러 번 지정 가능)")
    ap.add_argument("--no-cache", action="store_true",
                    help="문서 해시 캐시를 무시하고 강제로 재호출")
    args = ap.parse_args()

    if not args.pdf_path and not args.md_path:
        sys.exit("[ERROR] --pdf-path 또는 --md-path 중 하나는 지정해야 합니다.")

    extractor = IrExtractor()
    if args.no_cache:
        extractor.cache = {}

    source = args.pdf_path or args.md_path
    print(f"[1] IR 문서 읽는 중: {source}"
          + ("  (PDF 우선, 실패 시 .md 폴백)" if args.pdf_path and args.md_path else ""))
    extracted = extractor.extract_from_document(pdf_path=args.pdf_path, md_path=args.md_path)

    for kv in args.override:
        if "=" not in kv:
            sys.exit(f"[ERROR] --override는 key=value 형태여야 합니다: {kv!r}")
        k, v = kv.split("=", 1)
        if k not in SIGNAL_FIELDS:
            sys.exit(f"[ERROR] 알 수 없는 필드: {k!r} (허용: {SIGNAL_FIELDS})")
        # 최소한의 타입 추정 — 문자열 그대로 두면 안 되는 필드만 캐스팅
        if k in ("founding_year", "ipo_target_year", "ni_year"):
            v = int(v)
        elif k in ("net_income", "future_net_income", "revenue", "yoy_growth"):
            v = float(v)
        elif k == "ipo_mentioned":
            v = v.lower() in ("true", "1", "yes")
        extracted[k] = v
        print(f"    override 적용: {k} = {v}")

    print_review(args.case, extracted)

    signal_entry = {f: extracted.get(f) for f in SIGNAL_FIELDS}

    all_signals = load_ir_signals()
    is_update = args.case in all_signals
    all_signals[args.case] = signal_entry
    save_ir_signals(all_signals)

    action = "갱신" if is_update else "신규 추가"
    print(f"\n[2] ir_signals.json {action} 완료: \"{args.case}\"")
    print(f"    이제 run_all_cases.py를 돌리면 이 케이스에 자동으로 반영됩니다.")


if __name__ == "__main__":
    main()