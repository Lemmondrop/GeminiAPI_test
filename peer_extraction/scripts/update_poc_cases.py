"""
scripts/update_poc_cases.py
────────────────────────────────────────────────────────────
IR PDF에서 core.ir_parser.IrExtractor로 business_text(사업설명, Stage3 LLM
피어매칭용 issuer 텍스트)를 추출해 poc_cases.json의 해당 케이스 항목에 반영한다.

update_ir_signals.py와 다른 점:
  - 대상 파일이 다르다: ir_signals.json(밸류에이션용) 아님, poc_cases.json
    (피어추출 Stage1~4용). 둘은 소비자가 다른 완전히 별개 파일.
  - poc_cases.json은 배열이고, name으로 찾는다(딕셔너리 키 아님).
  - business_text 외 필드(seed_code, track_type, high_sim_exempt, note)는
    KSIC 매핑·트랙 판단 등 사람의 도메인 판단이 들어간 필드라 절대 건드리지
    않는다 — business_text 하나만 갱신 대상.
  - 사람이 이미 정성껏 써둔 business_text를 실수로 날리지 않도록, 기본은
    미리보기(기존 vs 신규 비교)만 하고 --apply를 명시해야 실제로 저장한다.
  - --case가 poc_cases.json에 아직 없는 케이스면 에러로 멈춘다 — 새 항목
    생성은 seed_code/track_type 같은 판단이 필요해 이 스크립트 책임 밖.

사용법:
  # 1) 미리보기만 (기본, 아무것도 안 바뀜)
  python scripts\\update_poc_cases.py --case "주식회사 케이더블유씨 (KWC)" \\
      --pdf-path "C:\\...\\(주)케이더블유씨 IR 자료(20250905).pdf"

  # 2) 확인 후 실제 반영
  python scripts\\update_poc_cases.py --case "..." --pdf-path "..." --apply
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

POC_CASES_PATH = ROOT_DIR / "poc_cases.json"


def load_poc_cases() -> list:
    if not POC_CASES_PATH.exists():
        sys.exit(f"[ERROR] {POC_CASES_PATH} 없음.")
    data = json.loads(POC_CASES_PATH.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        sys.exit(f"[ERROR] poc_cases.json이 배열이 아닙니다 — 예상 구조와 다름, "
                 f"수동 확인 필요(실제 타입: {type(data).__name__}).")
    return data


def save_poc_cases(data: list):
    POC_CASES_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def find_case(cases: list, name: str):
    matches = [c for c in cases if c.get("name") == name]
    if not matches:
        available = [c.get("name") for c in cases]
        sys.exit(f"[ERROR] poc_cases.json에 \"{name}\" 없음.\n"
                 f"  이 스크립트는 business_text만 갱신하므로, 새 케이스는 "
                 f"seed_code/track_type 등을 먼저 사람이 채워 poc_cases.json에 "
                 f"등록해야 합니다.\n  등록된 케이스: {available}")
    if len(matches) > 1:
        sys.exit(f"[ERROR] \"{name}\"이 poc_cases.json에 {len(matches)}번 중복 등록되어 "
                 f"있습니다 — 어느 걸 갱신할지 모호해 중단합니다. 직접 정리 후 재시도하세요.")
    return matches[0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", required=True,
                    help="poc_cases.json의 name 필드와 정확히 일치해야 함")
    ap.add_argument("--pdf-path", default=None, help="IR PDF 원본 경로 (우선 사용)")
    ap.add_argument("--md-path", default=None,
                    help="OCR 추출된 .md 경로 (PDF 실패 시 폴백, 또는 PDF 없을 때 단독 사용)")
    ap.add_argument("--apply", action="store_true",
                    help="실제로 poc_cases.json에 반영. 없으면 미리보기만 하고 종료.")
    ap.add_argument("--force", action="store_true",
                    help="원문 근거 미확인 고유명사가 있어도 강제로 저장(--apply와 함께 사용). "
                         "IR 원문과 직접 대조 후에만 사용 권장.")
    ap.add_argument("--no-cache", action="store_true",
                    help="문서 해시 캐시를 무시하고 강제로 재호출")
    args = ap.parse_args()

    if not args.pdf_path and not args.md_path:
        sys.exit("[ERROR] --pdf-path 또는 --md-path 중 하나는 지정해야 합니다.")

    cases = load_poc_cases()
    target = find_case(cases, args.case)  # 없으면 여기서 종료됨

    extractor = IrExtractor()
    if args.no_cache:
        extractor.cache = {}

    source = args.pdf_path or args.md_path
    print(f"[1] IR 문서 읽는 중: {source}")
    extracted = extractor.extract_from_document(pdf_path=args.pdf_path, md_path=args.md_path)
    new_text = extracted.get("business_text")

    old_text = target.get("business_text")
    print(f"\n=== business_text 비교: {args.case} ===")
    print(f"  [기존]\n  {old_text or '(없음)'}")
    print(f"\n  [신규 추출]\n  {new_text or '(추출 안 됨 — 문서에서 사업설명을 못 찾음)'}")

    # 고유명사 자동 검증 결과 — LLM 재질의가 아니라 원문 텍스트와의
    # 결정론적 단어 대조(core.ir_parser.verify_entities). "지어내지 마라"는
    # 프롬프트 지시만으론 반복적으로 뚫린 게 실사례로 확인돼서 추가된 단계.
    #
    # 주의: pdfplumber는 텍스트 레이어만 읽는다. 디자인된 IR 슬라이드는 배경
    # 박스·아이콘 등과 함께 텍스트가 이미지로 렌더링되는 구간이 있어서, 실제로
    # 원문에 있는 내용도 pdfplumber가 못 읽어 "미검증"으로 잘못 뜰 수 있다
    # (2026-07 케이더블유씨 실사례로 확인 — "광동"/"Shamrock Milk"가 실제
    # 슬라이드엔 있는데 PDF 텍스트 레이어엔 없었음). --md-path를 같이 주면 OCR
    # 텍스트도 검증 코퍼스에 포함되어 이 문제가 줄어드니, 가능하면 항상 같이 주는
    # 걸 권장한다. 그래서 "🚫 못 찾음"은 "원문에 없다"가 아니라 "우리 도구로는
    # 확인 못 했다"는 뜻으로 표현한다 — 사람이 최종 판단해야 한다.
    if not args.md_path:
        print(f"\n  ℹ --md-path 없이 실행했습니다 — 디자인 슬라이드의 이미지 렌더링 "
              f"텍스트는 PDF 텍스트 레이어만으론 못 읽을 수 있어 미검증이 실제보다 "
              f"많이 뜰 수 있습니다. OCR .md 파일이 있으면 --md-path로 같이 주는 걸 권장합니다.")

    verification = extracted.get("_entity_verification")
    has_unverified = False
    if verification is None:
        print(f"\n  ⚠ 고유명사 자동 검증을 못 했습니다(텍스트 추출 실패 또는 "
              f"스캔본 등으로 추정) — 원문과 전체 대조가 필요합니다.")
    elif verification["unverified"]:
        has_unverified = True
        corpus_note = ("PDF+OCR 텍스트" if verification.get("corpus_included_ocr_md") else "PDF 텍스트만")
        print(f"\n  🚫 아래 항목은 저희 도구({corpus_note})로는 원문에서 근거를 못 찾았습니다"
              f" — '원문에 없다'가 아니라 '이 도구로 확인 못 했다'는 뜻입니다. 직접 대조하세요:")
        for e in verification["unverified"]:
            print(f"      ⚠ {e}")
        if verification["verified"]:
            print(f"  (나머지 {len(verification['verified'])}개는 근거 확인됨: "
                  + ", ".join(verification["verified"]) + ")")
    elif verification["verified"]:
        print(f"\n  ✅ 사용된 고유명사 {len(verification['verified'])}개 전부 근거 확인됨.")

    if new_text is None:
        print(f"\n[중단] 신규 추출값이 없어 반영하지 않습니다. 기존 값을 그대로 둡니다.")
        return

    if not args.apply:
        print(f"\n※ 미리보기 모드입니다 — 실제로 반영하려면 --apply를 붙여서 다시 실행하세요.")
        print(f"  (seed_code/track_type/high_sim_exempt/note 등 다른 필드는 이 스크립트가 "
              f"절대 안 건드립니다.)")
        return

    if has_unverified and not args.force:
        print(f"\n[중단] 원문 근거 미확인 고유명사가 있어 --apply만으로는 저장하지 않습니다.")
        print(f"  위 🚫 항목들을 IR 원문과 직접 대조한 뒤, 문제없으면 --force를 추가해 재실행하세요.")
        print(f"  (실제로 존재하지 않는 고객사명이 그대로 저장된 사례가 있어 추가된 안전장치입니다.)")
        return

    target["business_text"] = new_text
    save_poc_cases(cases)
    print(f"\n[2] poc_cases.json 반영 완료: \"{args.case}\"의 business_text만 갱신됨.")
    if has_unverified:
        print(f"  ⚠ --force로 미검증 항목을 포함한 채 저장했습니다 — 사람이 직접 확인했다는 전제입니다.")


if __name__ == "__main__":
    main()