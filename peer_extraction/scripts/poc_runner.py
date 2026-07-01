"""
poc_runner.py
────────────────────────────────────────────────────────────
여러 비상장 딜의 PoC(피어 추출)를 일관된 형식으로 실행·비교.

센스톤 단일 케이스 검증에서 했던 작업(run.py 실행 → 깔때기·최종
피어 확인)을, 다수 IR 케이스에 대해 반복 자동화해 Stage 1~4가
업종이 달라져도 일관되게 작동하는지(일반화) 확인한다.

사용 흐름:
  1) IR 자료를 Claude에게 제공 → 회사명/KSIC/사업설명 추출
  2) 추출된 정보를 poc_cases.json에 케이스로 추가 (최초 실행 시 예시 자동 생성)
  3) python poc_runner.py              → 전체 케이스 순차 실행
     python poc_runner.py --case 센스톤 → 특정 케이스만 실행
     python poc_runner.py --list        → 등록된 케이스 목록만 출력

내부적으로 run.py를 그대로 서브프로세스 호출 — 이미 검증된 CLI
경로를 재사용하므로 파이프라인 내부 구조 변경에 영향받지 않음.

출력:
  poc_results/{회사명}.json       케이스별 구조화 결과(깔때기·최종피어·rationale)
  poc_results/{회사명}_raw.txt    run.py 원본 stdout (감사·디버깅용)
  poc_results/_summary.json       전체 케이스 비교 요약
  콘솔에 비교 테이블 출력
────────────────────────────────────────────────────────────
"""
import sys, json, re, subprocess, argparse, time, os
from pathlib import Path

SCRIPT_DIR   = Path(__file__).parent          # peer_extraction/scripts/
ROOT_DIR     = SCRIPT_DIR.parent                # peer_extraction/  (run.py 실행 위치)
CASES_PATH   = ROOT_DIR / "poc_cases.json"
RESULTS_DIR  = ROOT_DIR / "poc_results"
RUN_PY       = ROOT_DIR / "run.py"

VALID_TRACKS = {"general", "tech_special", "konex_transfer"}

# ══════════════════════════════════════════════════════════
# 케이스 파일 — 없으면 검증된 센스톤 예시로 자동 생성
# ══════════════════════════════════════════════════════════
DEFAULT_CASES = [
    {
        "name": "센스톤",
        "seed_code": "58221",
        "business_text": (
            "OTAC(일회용 인증코드) 원천기술 기반 인증보안 소프트웨어 기업. "
            "단방향 다이내믹 인증 코드로 사용자·기기·명령을 식별·인증. "
            "주요 제품: 출입통제·로그인·가상카드번호·금융거래 인증(SWIDCH ACCESS), "
            "IoT 및 디지털키 인증(SWIDCH CONNECT), FIDO·mOTP·OTAC 통합 인증 SDK. "
            "OT/PLC 산업제어 인증보안 게이트웨이, 금융 인증, 디지털 ID 분야에 적용."
        ),
        "track_type": "general",
        "high_sim_exempt": True,
        "headline": "sim_weighted",
        "note": "1차 PoC 검증 완료 (인증보안 SW). 이미 캐시 적중이라 즉시 재현됨 — 회귀 테스트 용도로 유지."
    },
    {
        "name": "TODO_두번째_케이스",
        "seed_code": "00000",
        "business_text": "TODO: IR 자료에서 추출한 사업설명으로 교체",
        "track_type": "general",
        "high_sim_exempt": True,
        "headline": "sim_weighted",
        "note": "TODO: 업종/IR 출처 메모"
    },
]


def load_cases() -> list[dict]:
    if not CASES_PATH.exists():
        CASES_PATH.write_text(
            json.dumps(DEFAULT_CASES, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[안내] {CASES_PATH.name} 없어서 예시 2건으로 새로 생성했습니다.")
        print(f"       TODO 케이스를 실제 IR 정보로 교체한 뒤 다시 실행하세요.\n")
    raw = json.loads(CASES_PATH.read_text(encoding="utf-8"))
    return raw


def validate_case(case: dict) -> list[str]:
    """필수 필드·값 검증. 문제 목록(빈 리스트면 정상) 반환."""
    errs = []
    for field in ("name", "seed_code", "business_text"):
        if not str(case.get(field, "")).strip():
            errs.append(f"필수 필드 누락: {field}")
    if case.get("name", "").startswith("TODO"):
        errs.append("name이 TODO placeholder 입니다 — 실제 케이스로 교체 필요")
    if case.get("business_text", "").startswith("TODO"):
        errs.append("business_text가 TODO placeholder 입니다 — 실제 사업설명으로 교체 필요")
    track = case.get("track_type", "general")
    if track not in VALID_TRACKS:
        errs.append(f"track_type 값 오류: {track!r} (허용: {VALID_TRACKS})")
    return errs


# ══════════════════════════════════════════════════════════
# run.py 호출 (서브프로세스, 실시간 출력 스트리밍)
# ══════════════════════════════════════════════════════════
def build_args(case: dict) -> list[str]:
    args = [
        sys.executable, str(RUN_PY),
        "--name", case["name"],
        "--seed", case["seed_code"],
        "--text", case["business_text"],
        "--track", case.get("track_type", "general"),
    ]
    if case.get("high_sim_exempt"):
        args.append("--high-sim-exempt")
    if case.get("headline"):
        args += ["--headline", case["headline"]]
    # 자유 확장 — 케이스별 추가 CLI 플래그 (예: ["--top-k","40"])
    for extra in case.get("extra_args", []):
        args.append(str(extra))
    return args


def run_case(case: dict) -> tuple[str, float]:
    """run.py 실행. (전체 stdout, 소요초) 반환. 실시간으로 콘솔에도 그대로 출력."""
    args = build_args(case)
    print(f"\n{'─'*60}")
    print(f"▶ 실행: {case['name']}  (track={case.get('track_type','general')})")
    print(f"{'─'*60}")

    t0 = time.time()
    # ── Windows 콘솔(cp949) ↔ UTF-8 인코딩 불일치로 인한 글자 깨짐 방지 ──
    # subprocess.PIPE로 잡으면 콘솔이 아니므로 Python이 시스템 기본
    # 코드페이지(cp949)로 인코딩 — PYTHONIOENCODING을 강제해 자식
    # 프로세스의 stdout 인코딩을 UTF-8로 고정한다.
    child_env = dict(os.environ)
    child_env["PYTHONIOENCODING"] = "utf-8"

    proc = subprocess.Popen(
        args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace", bufsize=1,
        cwd=str(ROOT_DIR),   # run.py 요구사항: 반드시 peer_extraction 루트에서 실행 
        env=child_env,
    )

    lines: list[str] = []
    for raw_line in proc.stdout:
        # run.py의 진행률 표시(\r 캐리지리턴)는 그대로 콘솔에 흘려보내되,
        # 저장용 텍스트는 줄 단위로 정리.
        print(raw_line, end="")
        lines.append(raw_line)
    proc.wait()
    elapsed = time.time() - t0

    full_text = "".join(lines)
    if proc.returncode != 0:
        print(f"\n[경고] run.py 종료코드 {proc.returncode} (비정상 종료 가능)")
    return full_text, elapsed


# ══════════════════════════════════════════════════════════
# run.py stdout 파싱 — 깔때기·최종피어·대표멀티플 추출
# ══════════════════════════════════════════════════════════
FUNNEL_RE = re.compile(
    r"깔때기:\s*모집단\s*(\d+)\s*→\s*재무통과\s*(\d+)\s*→\s*사업유사\s*(\d+)\s*→\s*최종\s*(\d+)"
)
PEER_BLOCK_RE = re.compile(
    r"^\s*(\S.+?)\s+([0-9A-Za-z]{6})\s{2,}(높음|중간|낮음)\s+유사도\s+(\d+)\s+PER\s+([\d.]+)\s*$"
)
RATIONALE_RE = re.compile(r"^\s*└\s*(.+)$")
HEADLINE_RE = re.compile(
    r"대표\s*멀티플\(PER,\s*(\w+)\):\s*([\d.]+)\s*\[(.*?)\]"
)
NO_PEER_CANDIDATE_RE = re.compile(
    r"^\s*(\S.+?)\s{2,}(높음|중간|낮음)\s+유사도\s+(\d+)\s*$"
)


def parse_run_output(text: str) -> dict:
    result = {
        "funnel": None,            # {pool, stage2, stage3, final}
        "final_peers": [],         # [{name, tier, sim_score, per, rationale}]
        "headline": None,          # {method, value, detail}
        "no_peers": False,
        "stage3_candidates_only": [],  # 최종 0건일 때 Stage3 통과 후보 목록
    }

    m = FUNNEL_RE.search(text)
    if m:
        result["funnel"] = {
            "pool":    int(m.group(1)),
            "stage2":  int(m.group(2)),
            "stage3":  int(m.group(3)),
            "final":   int(m.group(4)),
        }

    if "최종 비교군 없음" in text:
        result["no_peers"] = True
        for line in text.splitlines():
            m2 = NO_PEER_CANDIDATE_RE.match(line)
            if m2:
                result["stage3_candidates_only"].append({
                    "name": m2.group(1).strip(),
                    "tier": m2.group(2),
                    "sim_score": int(m2.group(3)),
                })
    else:
        lines = text.splitlines()
        i = 0
        while i < len(lines):
            m2 = PEER_BLOCK_RE.match(lines[i])
            if m2:
                rationale = ""
                if i + 1 < len(lines):
                    m3 = RATIONALE_RE.match(lines[i + 1])
                    if m3:
                        rationale = m3.group(1).strip()
                result["final_peers"].append({
                    "name": m2.group(1).strip(),
                    "종목코드": m2.group(2),
                    "tier": m2.group(3),
                    "sim_score": int(m2.group(4)),
                    "per": float(m2.group(5)),
                    "rationale": rationale,
                })
            i += 1

    m4 = HEADLINE_RE.search(text)
    if m4:
        result["headline"] = {
            "method": m4.group(1),
            "value": float(m4.group(2)),
            "detail": m4.group(3).strip(),
        }

    return result


# ══════════════════════════════════════════════════════════
# 메인
# ══════════════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", type=str, default=None,
                    help="특정 케이스 이름만 실행 (미지정 시 전체)")
    ap.add_argument("--list", action="store_true",
                    help="등록된 케이스 목록만 출력하고 종료")
    args = ap.parse_args()

    if not RUN_PY.exists():
        sys.exit(f"[ERROR] run.py를 찾을 수 없습니다: {RUN_PY}")

    cases = load_cases()

    if args.list:
        print(f"등록된 케이스 {len(cases)}건:")
        for c in cases:
            flag = "  [TODO — 미완성]" if c["name"].startswith("TODO") else ""
            print(f"  - {c['name']}  (track={c.get('track_type','general')}){flag}")
        return

    if args.case:
        cases = [c for c in cases if c["name"] == args.case]
        if not cases:
            sys.exit(f"[ERROR] 케이스 '{args.case}'를 poc_cases.json에서 찾을 수 없습니다.")

    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    summary_rows = []

    for case in cases:
        errs = validate_case(case)
        if errs:
            print(f"\n[SKIP] {case.get('name','?')} — 케이스 검증 실패:")
            for e in errs:
                print(f"    · {e}")
            continue

        raw_text, elapsed = run_case(case)
        parsed = parse_run_output(raw_text)

        # ── 케이스별 결과 저장 ──────────────────────────────
        case_result = {
            "case": case,
            "elapsed_sec": round(elapsed, 1),
            **parsed,
        }
        out_json = RESULTS_DIR / f"{case['name']}.json"
        out_json.write_text(
            json.dumps(case_result, ensure_ascii=False, indent=2), encoding="utf-8")
        (RESULTS_DIR / f"{case['name']}_raw.txt").write_text(raw_text, encoding="utf-8")

        funnel = parsed["funnel"] or {}
        summary_rows.append({
            "name": case["name"],
            "track": case.get("track_type", "general"),
            "pool": funnel.get("pool"),
            "stage2": funnel.get("stage2"),
            "stage3": funnel.get("stage3"),
            "final": funnel.get("final"),
            "headline_per": parsed["headline"]["value"] if parsed["headline"] else None,
            "elapsed_sec": round(elapsed, 1),
        })

        print(f"\n✓ {case['name']} 완료 — 결과 저장: {out_json.name}  ({elapsed:.0f}초)")

    # ── 전체 비교 요약 ──────────────────────────────────────
    summary_path = RESULTS_DIR / "_summary.json"
    summary_path.write_text(
        json.dumps(summary_rows, ensure_ascii=False, indent=2), encoding="utf-8")

    if summary_rows:
        print(f"\n\n{'='*70}")
        print("케이스 비교 요약")
        print(f"{'='*70}")
        print(f"{'회사명':14s} {'트랙':14s} {'모집단':>6s} {'재무통과':>8s} "
              f"{'사업유사':>8s} {'최종':>6s} {'대표PER':>8s} {'소요(초)':>8s}")
        print("-"*70)
        for r in summary_rows:
            per_str = f"{r['headline_per']:.1f}" if r["headline_per"] is not None else "-"
            print(f"{r['name']:14s} {r['track']:14s} "
                  f"{str(r['pool']):>6s} {str(r['stage2']):>8s} "
                  f"{str(r['stage3']):>8s} {str(r['final']):>6s} "
                  f"{per_str:>8s} {r['elapsed_sec']:>8.0f}")
        print(f"\n  요약 저장: {summary_path}")
        print(f"  케이스별 상세: {RESULTS_DIR}\\{{회사명}}.json")

if __name__ == "__main__":
    main()