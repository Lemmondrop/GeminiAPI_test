"""Stage 2 — 재무 유사성 필터.

stage1 모집단을 universe_financials와 조인해 TRACK_RULES[track_type]['financial'] 적용:
  - 결산월 == 12
  - 감사의견 startswith('적정')        (DART는 '적정의견'으로 반환)
  - 트랙별 순이익/영업이익 온기·LTM 조건 (general / tech_special / konex_transfer)

설계 원칙: 각 행에 통과여부(passed_stage2)와 drop_reason(감사용 provenance)을 태깅.
재무 결측사는 검증 불가이므로 보수적으로 탈락.
track_type 분기는 TRACK_RULES 한 곳에서만 — 트랙별 스크립트 복제 없음.
"""
from __future__ import annotations
import operator
import re

import pandas as pd

from config import (FISCAL_MONTH_COL, STOCK_COL, TRACK_RULES, UNIVERSE_FIN_CSV,
                    VALID_TRACKS)

_OPS = {">=": operator.ge, "<=": operator.le, "==": operator.eq,
        ">": operator.gt, "<": operator.lt}
_COND = re.compile(r"\s*(\w+)\s*(>=|<=|==|>|<)\s*(-?\d+\.?\d*)\s*")

FIN_COLS = [FISCAL_MONTH_COL, "op_income_annual", "op_income_ltm",
            "net_income_annual", "net_income_ltm", "revenue_annual", "revenue_ltm",
            "audit_opinion", "auditor", "market_cap", "has_financials"]


def _audit_ok(opinion) -> bool:
    return bool(opinion) and not pd.isna(opinion) and str(opinion).strip().startswith("적정")


def _has_fin(v) -> bool:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    return str(v).strip().lower() in ("true", "1", "1.0")


def _fiscal_month(v) -> int | None:
    """'12월' / '12' / 12 / 12.0 → 12. 숫자 없으면 None."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    digits = "".join(ch for ch in str(v) if ch.isdigit())
    return int(digits) if digits else None


def _cond_ok(row, cond: str) -> bool:
    """'net_income_ltm>0' 형태 조건 평가. 결측이면 False(보수적 탈락)."""
    m = _COND.fullmatch(cond)
    if not m:
        raise ValueError(f"조건 파싱 실패: {cond}")
    field, op, val = m.group(1), m.group(2), float(m.group(3))
    v = row.get(field)
    if v is None or pd.isna(v):
        return False
    try:
        return _OPS[op](float(v), val)
    except (TypeError, ValueError):
        return False


def load_financials(path=UNIVERSE_FIN_CSV) -> pd.DataFrame:
    df = pd.read_csv(path, dtype={STOCK_COL: str})
    df[STOCK_COL] = df[STOCK_COL].astype(str).str.split(".").str[0].str.zfill(6)
    if "has_financials" in df.columns:
        df["has_financials"] = df["has_financials"].map(_has_fin)
        # 종목코드 중복 시 재무 보유 행 우선 유지 (merge fan-out 방지)
        df = df.sort_values("has_financials", ascending=False)
    df = df.drop_duplicates(subset=[STOCK_COL], keep="first").reset_index(drop=True)
    return df


def apply_stage2(pool: pd.DataFrame, financials: pd.DataFrame,
                 track_type: str) -> pd.DataFrame:
    """모집단 ⨝ 재무 → 트랙별 필터. 통과여부·사유·재무컬럼이 붙은 DataFrame 반환."""
    if track_type not in VALID_TRACKS:
        raise ValueError(f"track_type must be one of {VALID_TRACKS}")
    rules = TRACK_RULES[track_type]["financial"]

    p = pool.copy()
    p[STOCK_COL] = p[STOCK_COL].astype(str).str.split(".").str[0].str.zfill(6)
    p = p.drop_duplicates(subset=[STOCK_COL], keep="first")   # 모집단 종목코드 유일화
    fin_cols = [c for c in FIN_COLS if c in financials.columns]
    fin_sub = financials[[STOCK_COL, *fin_cols]].drop_duplicates(subset=[STOCK_COL], keep="first")
    merged = p.merge(fin_sub, on=STOCK_COL, how="left")

    passed, reasons = [], []
    for _, row in merged.iterrows():
        why: list[str] = []
        if not _has_fin(row.get("has_financials")):
            why.append("재무없음")
        else:
            fm = _fiscal_month(row.get(FISCAL_MONTH_COL))
            if fm is None or fm != rules["fiscal_month"]:
                why.append(f"결산월≠{rules['fiscal_month']}")
            if not _audit_ok(row.get("audit_opinion")):
                why.append("감사의견부적정")
            for c in rules["require"]:
                if not _cond_ok(row, c):
                    why.append(f"미충족:{c}")
        passed.append(len(why) == 0)
        reasons.append(";".join(why))

    merged["passed_stage2"] = passed
    merged["drop_reason"] = reasons
    return merged


def survivors(stage2_df: pd.DataFrame) -> pd.DataFrame:
    """통과한 피어만."""
    return stage2_df[stage2_df["passed_stage2"]].reset_index(drop=True)