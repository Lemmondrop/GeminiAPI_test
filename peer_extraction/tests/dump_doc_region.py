"""신고서 본문의 특정 구간을 그대로 덤프 — 비교회사 PER 표 위치 확인용.

'PER 상대가치' 같은 앵커 주변(앞·뒤)을 넓게 출력해서 실제 표 구조를 눈으로 확인.
실행:
  python tests\\dump_doc_region.py --rcept 20220118000223 --anchor "PER" --before 4000 --after 500
  python tests\\dump_doc_region.py --rcept 20220118000223 --start 64000 --len 4000
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.dart_client import DartClient


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rcept", required=True)
    ap.add_argument("--anchor", default=None, help="이 문자열 첫 등장 기준으로 구간 출력")
    ap.add_argument("--occur", type=int, default=1, help="앵커의 몇 번째 등장(1-based)")
    ap.add_argument("--before", type=int, default=3000, help="앵커 앞 글자수")
    ap.add_argument("--after", type=int, default=500, help="앵커 뒤 글자수")
    ap.add_argument("--start", type=int, default=None, help="절대 시작 위치(앵커 대신)")
    ap.add_argument("--len", type=int, default=4000, help="start 사용 시 길이")
    a = ap.parse_args()

    text = DartClient().fetch_document(a.rcept)
    print(f"본문 {len(text):,}자\n" + "=" * 60)

    if a.start is not None:
        seg = text[a.start:a.start + a.len]
        print(f"[{a.start} ~ {a.start + a.len}]")
    elif a.anchor:
        pos, cnt = -1, 0
        start = 0
        while cnt < a.occur:
            pos = text.find(a.anchor, start)
            if pos < 0:
                break
            cnt += 1
            start = pos + 1
        if pos < 0:
            print(f"'{a.anchor}' {a.occur}번째 등장 없음")
            return
        lo, hi = max(0, pos - a.before), pos + a.after
        seg = text[lo:hi]
        print(f"앵커 '{a.anchor}' @{pos} (구간 {lo}~{hi})")
    else:
        seg = text[:a.len]
        print(f"[0 ~ {a.len}]")

    print("=" * 60)
    print(seg)


if __name__ == "__main__":
    main()
