r"""
find_dead_code.py
────────────────────────────────────────────────────────────
지정한 entry point(들)에서 시작해, 실제로 참조되는 .py 파일을 추적해서
"살아있는 파일"과 "후보(고아일 수 있는) 파일"을 구분해 보여준다.

추측 없이 확인하는 게 목적이라, 아래 두 가지 관계를 모두 추적한다:
  1) 진짜 import (import X / from X import Y) — ast로 정확히 파싱.
  2) subprocess 실행 관계 (예: poc_runner.py가 run.py를 subprocess.Popen으로
     실행) — import가 아니라서 1번만으로는 못 잡는다. 소스 코드에서
     subprocess 관련 함수(Popen/run/call/check_call/check_output) 호출부
     "근처"에 있는 문자열 중 프로젝트 안에 실재하는 .py 파일명과 일치하는
     게 있으면 실행 의존성으로 간주한다(휴리스틱 — 안전한 쪽으로,
     즉 "혹시 몰라 살아있다고 표시"하는 쪽으로 치우친 판단을 한다).

주의:
  - "후보(고아)" 목록은 "삭제해도 됨"이 아니라 "확인이 필요함"이다.
    동적 import(importlib.import_module(변수)), 노트북에서만 쓰는 파일,
    문서화 목적으로 남겨둔 참고 스크립트 등은 이 도구가 놓칠 수 있다.
  - entry point는 여러 개를 줄 수 있고, 실제로 사람이 실행하는 진입점
    (run.py, scripts\\*.py 중 사용자가 직접 돌리는 것들)을 전부 넣어야
    정확도가 올라간다.

사용법:
  python find_dead_code.py --root "C:\...\peer_extraction" \\
      --entry run.py --entry scripts\\run_full_valuation.py \\
      --entry scripts\\run_all_cases.py --entry scripts\\poc_runner.py \\
      --entry scripts\\update_ir_signals.py

  --root을 여러 번 줘서 여러 프로젝트 폴더를 한 번에 스캔할 수도 있다:
  python find_dead_code.py --root "...\\peer_extraction" --root "...\\discount_rate_pipeline" \\
      --entry "...\\peer_extraction\\run.py" --entry "...\\discount_rate_pipeline\\main.py"

  "확인 필요" 목록 중 일부(예: 데이터 구축 스크립트 묶음)를 실제로 옮기기 전에,
  그 안에서 서로 어떤 core 모듈을 실제로 쓰는지 --cluster-entry로 검증할 수 있다.
  각 파일이 뭘 참조하는지, 그중 라이브 파이프라인과 겹치는 게 있는지(있으면
  이동 금지 신호)까지 같이 보여준다:
  python find_dead_code.py --root "...\\peer_extraction" \\
      --entry run.py --entry scripts\\run_full_valuation.py ... (기존 entry 그대로) \\
      --cluster-entry scripts\\build_cooccur_prior.py \\
      --cluster-entry scripts\\collect_financials.py \\
      --cluster-entry scripts\\dart_biz_batch.py
────────────────────────────────────────────────────────────
"""
from __future__ import annotations

import argparse
import ast
import re
from pathlib import Path
from typing import Optional


SUBPROCESS_CALL_RE = re.compile(
    r"subprocess\.(Popen|run|call|check_call|check_output)\s*\("
)
# subprocess 호출 근처(같은 함수 블록 전체를 다 파싱하긴 복잡하니, 실용적으로
# "파일 전체에서" .py로 끝나는 문자열 리터럴을 찾는다 — 오탐(과다 포함)은
# 안전한 방향이라 허용, 누락(과소 포함)이 위험한 방향이라 피한다.
PY_STRING_LITERAL_RE = re.compile(r"""["']([^"']*?\.py)["']""")


def find_all_py_files(roots: list[Path]) -> dict[str, Path]:
    """루트들 아래 모든 .py 파일. {파일명(확장자 제외): 경로} — 이름 충돌 시
    나중에 발견된 게 덮어쓰므로, 충돌 목록은 별도로 경고한다."""
    all_files: dict[str, Path] = {}
    duplicates: dict[str, list[Path]] = {}
    for root in roots:
        for p in root.rglob("*.py"):
            if "__pycache__" in p.parts or ".venv" in p.parts:
                continue
            stem = p.stem
            if stem in all_files and all_files[stem] != p:
                duplicates.setdefault(stem, [all_files[stem]]).append(p)
            all_files[stem] = p
    if duplicates:
        print("[경고] 같은 이름의 .py 파일이 여러 경로에 존재 (마지막 것만 추적 대상):")
        for stem, paths in duplicates.items():
            print(f"    {stem}.py:")
            for p in paths:
                print(f"      - {p}")
    return all_files


def parse_imports(path: Path) -> set[str]:
    """파일 하나에서 import된 모듈명(최상위 이름만, 예: 'core.peer_join' → 'peer_join'
    이 아니라 'core'와 'peer_join' 둘 다 후보로 반환 — 아래서 실제 파일과 매칭)."""
    try:
        tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
    except SyntaxError as e:
        print(f"[경고] 파싱 실패(문법 오류?) — 건너뜀: {path} ({e})")
        return set()

    names: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                # "import core.peer_join" → "core", "core.peer_join" 둘 다 후보
                parts = alias.name.split(".")
                names.add(parts[0])
                names.add(parts[-1])
        elif isinstance(node, ast.ImportFrom):
            if node.module:
                parts = node.module.split(".")
                names.add(parts[0])
                names.add(parts[-1])
    return names


def parse_subprocess_refs(path: Path, known_py_stems: set[str]) -> set[str]:
    """subprocess 호출이 파일 안에 있으면, 그 파일 전체에서 .py로 끝나는
    문자열 리터럴을 찾아 프로젝트 안에 실재하는 파일명과 매칭한다."""
    text = path.read_text(encoding="utf-8", errors="replace")
    if not SUBPROCESS_CALL_RE.search(text):
        return set()
    refs = set()
    for m in PY_STRING_LITERAL_RE.finditer(text):
        raw = m.group(1)
        stem = Path(raw.replace("\\", "/")).stem
        if stem in known_py_stems:
            refs.add(stem)
    return refs


def trace_reachable(entry_points: list[Path], all_files: dict[str, Path]) -> set[str]:
    known_stems = set(all_files.keys())
    reachable: set[str] = set()
    queue: list[str] = []

    for ep in entry_points:
        stem = ep.stem
        if stem not in all_files:
            print(f"[경고] entry point가 스캔된 파일 목록에 없음(경로 확인 필요): {ep}")
            continue
        reachable.add(stem)
        queue.append(stem)

    while queue:
        cur_stem = queue.pop()
        cur_path = all_files[cur_stem]

        imported = parse_imports(cur_path) & known_stems
        subprocess_refs = parse_subprocess_refs(cur_path, known_stems)

        for nxt in (imported | subprocess_refs):
            if nxt not in reachable:
                reachable.add(nxt)
                queue.append(nxt)

    return reachable


def trace_closure(start_stem: str, all_files: dict[str, Path]) -> set[str]:
    """단일 파일 하나에서 시작해 도달 가능한 로컬 의존성 전체(자기 자신 제외).
    클러스터 분석용 — "이 스크립트를 옮기면 뭘 같이 옮겨야 하는지" 확인."""
    known_stems = set(all_files.keys())
    visited = {start_stem}
    queue = [start_stem]
    while queue:
        cur = queue.pop()
        cur_path = all_files[cur]
        imported = parse_imports(cur_path) & known_stems
        subprocess_refs = parse_subprocess_refs(cur_path, known_stems)
        for nxt in (imported | subprocess_refs):
            if nxt not in visited:
                visited.add(nxt)
                queue.append(nxt)
    visited.discard(start_stem)
    return visited


def resolve_path(raw: str, roots: list[Path]) -> Path:
    p = Path(raw)
    if p.is_absolute():
        return p
    for r in roots:
        candidate = r / raw
        if candidate.exists():
            return candidate
    raise SystemExit(f"[ERROR] 경로를 찾을 수 없습니다: {raw}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", action="append", required=True,
                    help="스캔할 프로젝트 루트 (여러 번 지정 가능)")
    ap.add_argument("--entry", action="append", required=True,
                    help="라이브 파이프라인 진입점 (여러 번 지정 가능)")
    ap.add_argument("--cluster-entry", action="append", default=[],
                    help="개별 의존관계를 확인하고 싶은 파일들 (예: 데이터 구축 스크립트) "
                         "— 여러 번 지정 가능. 지정하면 각 파일이 실제로 뭘 참조하는지, "
                         "그중 라이브 파이프라인과 겹치는 게 있는지까지 보여준다.")
    args = ap.parse_args()

    roots = [Path(r) for r in args.root]
    for r in roots:
        if not r.exists():
            raise SystemExit(f"[ERROR] --root 경로가 없습니다: {r}")

    entry_points = [resolve_path(e, roots) for e in args.entry]

    print(f"스캔 루트: {[str(r) for r in roots]}")
    print(f"진입점: {[str(e) for e in entry_points]}\n")

    all_files = find_all_py_files(roots)
    print(f"발견된 .py 파일 총 {len(all_files)}개 (중복 이름 제외 유니크 기준)\n")

    reachable = trace_reachable(entry_points, all_files)
    unreachable = sorted(set(all_files.keys()) - reachable)

    print(f"{'='*70}")
    print(f"살아있는 파일 (entry point에서 도달 가능): {len(reachable)}개")
    print(f"{'='*70}")
    for stem in sorted(reachable):
        print(f"  ✅ {all_files[stem]}")

    print(f"\n{'='*70}")
    print(f"후보(고아 가능) — entry point에서 도달 안 됨: {len(unreachable)}개")
    print(f"{'='*70}")
    print("  ※ '삭제해도 됨'이 아니라 '확인이 필요함'입니다.")
    print("     동적 import, 노트북 전용, 수동 실행용 진단 스크립트는 여기 걸립니다.\n")
    for stem in unreachable:
        print(f"  ❓ {all_files[stem]}")

    print(f"\n요약: 전체 {len(all_files)}개 중 살아있음 {len(reachable)}개, "
          f"확인 필요 {len(unreachable)}개 "
          f"({len(unreachable)/len(all_files)*100:.0f}%)")

    if not args.cluster_entry:
        return

    cluster_entries = [resolve_path(e, roots) for e in args.cluster_entry]

    print(f"\n\n{'='*70}")
    print(f"클러스터 분석 — 개별 스크립트별 실제 의존관계 ({len(cluster_entries)}개 확인)")
    print(f"{'='*70}")
    print("  ⚠ = 라이브 파이프라인과 공유(옮기면 라이브도 영향받을 수 있음 — 이동 금지)")
    print("  →  = 이 스크립트(들) 전용 의존성(같이 옮겨야 함)\n")

    all_cluster_deps: set[str] = set()
    per_file_closure: dict[str, set[str]] = {}

    for ce in cluster_entries:
        stem = ce.stem
        if stem not in all_files:
            print(f"[경고] cluster entry를 스캔된 파일 목록에서 못 찾음: {ce}")
            continue
        closure = trace_closure(stem, all_files)
        per_file_closure[stem] = closure
        all_cluster_deps |= closure

        shared_with_live = closure & reachable
        unique_to_cluster = closure - reachable

        print(f"📄 {stem}.py  ({all_files[stem]})")
        if shared_with_live:
            for s in sorted(shared_with_live):
                print(f"    ⚠ {all_files[s]}")
        if unique_to_cluster:
            for s in sorted(unique_to_cluster):
                print(f"    →  {all_files[s]}")
        if not closure:
            print(f"    (로컬 의존성 없음 — 이 파일만 단독으로 옮겨도 안전)")
        print()

    combined_shared = all_cluster_deps & reachable
    combined_unique = all_cluster_deps - reachable

    print(f"{'='*70}")
    print("클러스터 전체 요약 (지정한 파일들을 다 같이 옮긴다면)")
    print(f"{'='*70}")
    if combined_shared:
        print(f"\n⚠ 라이브 파이프라인과 공유 — 이 파일들은 절대 옮기면 안 됨 "
              f"({len(combined_shared)}개):")
        for s in sorted(combined_shared):
            # 어떤 클러스터 파일이 이걸 참조하는지도 같이 보여준다
            users = [k for k, v in per_file_closure.items() if s in v]
            print(f"    ⚠ {all_files[s]}  (참조: {', '.join(users)})")
    if combined_unique:
        print(f"\n→ 클러스터 전용 의존성 — 함께 이동해야 함 ({len(combined_unique)}개):")
        for s in sorted(combined_unique):
            print(f"    → {all_files[s]}")
    if not combined_shared and not combined_unique:
        print("\n로컬 의존성이 전혀 없음 — 지정한 파일들 그대로 이동해도 안전.")


if __name__ == "__main__":
    main()