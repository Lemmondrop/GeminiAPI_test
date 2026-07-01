"""과거 신고서 피어그룹 → 동시출현(co-occurrence) prior 행렬.

입력: 과거 IPO들의 (발행사, 선정된 피어들) 관계.
출력: {발행사_소분류: {피어_소분류: weight}}  — Stage1 Lane C가 키를 recall 신호로 사용.

설계 원칙
  - 피어는 모두 상장사이므로 회사명→KSIC는 상장 유니버스 CSV로 해소 가능.
  - 동시출현은 '소분류 단위 · IPO 단위 집계' — 한 IPO에서 같은 소분류 피어가 여럿이어도 1로 셈
    (한 건이 가중치를 독점하지 않도록). 따라서 count = 그 링크가 등장한 IPO 수.
  - min_count 임계 이상만 유지(단발 노이즈 제거) → 발행사별 행 정규화.
  - 발행사 자기 소분류는 prior에서 제외(Lane A 코드매칭이 이미 담당).

수동 하드코딩 prior({303:{242,281}})를 이 데이터 산출물로 대체하는 것이 목적.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

import pandas as pd

from config import CODE_COL, COOCCUR_PRIOR_JSON, CSV_ENCODING, LISTED_UNIVERSE_CSV, NAME_COL
from core.ksic import parse_ksic


# ----------------------------- 회사명 → 소분류 해소 -----------------------------
def _read_universe(path: Path = LISTED_UNIVERSE_CSV) -> pd.DataFrame:
    try:
        return pd.read_csv(path, encoding=CSV_ENCODING, dtype=str)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="utf-8-sig", dtype=str)


def name_to_sub_map(universe: pd.DataFrame | None = None) -> dict[str, str]:
    """{회사명(정규화): 소분류3자리}. 피어 회사명을 KSIC 소분류로 해소하는 사전."""
    df = universe if universe is not None else _read_universe()
    out: dict[str, str] = {}
    for _, r in df.iterrows():
        name = _norm_name(r.get(NAME_COL))
        code = r.get(CODE_COL)
        if name and code is not None and str(code).strip():
            out[name] = parse_ksic(str(code))[2]      # 소분류
    return out


def _norm_name(v) -> str:
    """회사명 정규화: 공백/대소문자/(주)·㈜ 제거 → 매칭 견고화."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    for token in ("주식회사", "(주)", "㈜", " "):
        s = s.replace(token, "")
    return s.lower()


_CODE_RE = re.compile(r"^[A-Za-z]\d{2,}$")            # KSIC 코드: 라틴문자1 + 숫자2+ (C00303)


def resolve_sub(token: str, name2sub: dict[str, str]) -> str | None:
    """피어 식별자 → 소분류. 코드(C00303/303)면 직접 파싱, 회사명이면 사전 조회.

    주의: 한글 첫 글자도 isalpha()=True이므로 코드 판별은 정규식(라틴+숫자)으로 한다
    (예: 'AP위성'·'쎄트렉아이'는 코드가 아니라 회사명 → 사전 조회).
    """
    s = str(token).strip()
    if not s:
        return None
    if _CODE_RE.match(s):                             # 'C00303' 형태 코드
        return parse_ksic(s)[2]
    if s.isdigit() and len(s) >= 3:                   # 순수 숫자코드 '303'/'30320'
        return s[:3]
    return name2sub.get(_norm_name(s))                # 회사명


# ----------------------------- prior 빌드 -----------------------------
def build_prior(
    ipo_peers: list[tuple[str, list[str]]],
    name2sub: dict[str, str],
    min_count: int = 2,
    normalize: bool = True,
) -> dict[str, dict[str, float]]:
    """과거 IPO별 (발행사, 피어목록) → {발행사_소분류: {피어_소분류: weight}}.

    ipo_peers: [(발행사_식별자, [피어_식별자...]), ...]  식별자 = 회사명 또는 코드.
    count = 해당 (발행사소분류→피어소분류) 링크가 등장한 IPO 수(소분류·IPO 단위).
    """
    counts: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    unresolved: set[str] = set()

    for issuer, peers in ipo_peers:
        isub = resolve_sub(issuer, name2sub)
        if isub is None:
            unresolved.add(str(issuer))
            continue
        peer_subs = set()
        for p in peers:
            psub = resolve_sub(p, name2sub)
            if psub is None:
                unresolved.add(str(p))
            elif psub != isub:                        # 자기 소분류 제외(Lane A 담당)
                peer_subs.add(psub)
        for psub in peer_subs:                        # IPO당 1회 카운트
            counts[isub][psub] += 1

    prior: dict[str, dict[str, float]] = {}
    for isub, row in counts.items():
        kept = {p: c for p, c in row.items() if c >= min_count}
        if not kept:
            continue
        if normalize:
            tot = sum(kept.values())
            prior[isub] = {p: round(c / tot, 4) for p, c in sorted(
                kept.items(), key=lambda kv: kv[1], reverse=True)}
        else:
            prior[isub] = {p: float(c) for p, c in sorted(
                kept.items(), key=lambda kv: kv[1], reverse=True)}

    prior["_meta"] = {                                # type: ignore[assignment]
        "n_ipos": len(ipo_peers),
        "n_issuer_subs": len([k for k in prior if k != "_meta"]),
        "min_count": min_count,
        "normalized": normalize,
        "n_unresolved": len(unresolved),
    }
    return prior


# ----------------------------- 입출력 -----------------------------
def save_prior(prior: dict, path: Path = COOCCUR_PRIOR_JSON) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(prior, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def load_prior(path: Path = COOCCUR_PRIOR_JSON) -> dict[str, dict[str, float]]:
    """저장된 prior 로드. Stage1이 쓰는 {소분류:{소분류:weight}} 형태(_meta 제외)."""
    if not Path(path).exists():
        return {}
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    return {k: v for k, v in raw.items() if k != "_meta"}


def edges_from_csv(path: str | Path,
                   issuer_col: str = "issuer",
                   peer_col: str = "peer") -> list[tuple[str, list[str]]]:
    """엣지리스트 CSV(issuer,peer 각 1행) → [(issuer, [peers...])]. 발행사별 그룹화."""
    df = pd.read_csv(path, dtype=str).fillna("")
    grouped: dict[str, list[str]] = defaultdict(list)
    for _, r in df.iterrows():
        iss, peer = str(r[issuer_col]).strip(), str(r[peer_col]).strip()
        if iss and peer:
            grouped[iss].append(peer)
    return list(grouped.items())
