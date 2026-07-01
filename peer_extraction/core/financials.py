"""재무 수집: DART 주요계정 → 온기/LTM 영업이익·순이익·매출, 감사의견, pykrx 시총.

LTM(최근 12개월) = FY(사업보고서 당기) − YTD(전기 동기 누적) + YTD(당기 누적)
  - 분기보고서 응답은 thstrm_amount(당기 누적)·frmtrm_amount(전기 동기 누적)를 함께 제공.
  - 분기 데이터가 없으면 LTM = 온기(FY)로 대체.

재무제표 기준: FS_DIV_PREFERENCE(연결 CFS 우선, 없으면 별도 OFS).
※ 추출 로직은 네트워크 없이 단위검증 가능(test_financials). 라이브 검증은 collect_financials --smoke.
"""
from __future__ import annotations
import json
import time
from typing import Iterable

from config import (ANNUAL_FALLBACK_YEAR, ANNUAL_YEAR, FISCAL_MONTH_COL,
                    FS_DIV_PREFERENCE, LTM_QUARTER_CODE, LTM_QUARTER_YEAR,
                    NAME_COL, REPRT_ANNUAL, STOCK_COL, UNIVERSE_FIN_CSV,
                    FIN_CKPT_JSON)

# account_nm 매칭 후보 (공백 제거 후 부분일치)
ACCOUNTS = {
    "revenue": ("매출액", "수익(매출액)", "영업수익", "매출"),
    "op_income": ("영업이익", "영업이익(손실)", "영업손익"),
    "net_income": ("당기순이익", "당기순이익(손실)", "당기순손익", "분기순이익", "반기순이익"),
}


def _to_num(s) -> float | None:
    """'1,234' / '(1,234)' / '△1,234' / '' → float|None."""
    if s is None:
        return None
    t = str(s).strip().replace(",", "")
    if t in ("", "-"):
        return None
    neg = (t.startswith("(") and t.endswith(")")) or t.startswith("△") or t.startswith("▲") or t.startswith("-")
    t = t.strip("()").lstrip("△▲-")
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


def _pick(items: list[dict], keys: tuple[str, ...],
          fs_pref: tuple[str, ...] = FS_DIV_PREFERENCE):
    """fs 우선순위로 손익계산서(IS/CIS)에서 계정 매칭 → (당기, 전기동기, fs_div)."""
    for fs in fs_pref:
        for it in items:
            if it.get("fs_div") != fs:
                continue
            if it.get("sj_div") not in ("IS", "CIS"):
                continue
            nm = (it.get("account_nm") or "").replace(" ", "")
            if any(k in nm for k in keys):
                return _to_num(it.get("thstrm_amount")), _to_num(it.get("frmtrm_amount")), fs
    return None, None, None


def compute_metrics(annual_items: list[dict], q_items: list[dict],
                    fs_pref: tuple[str, ...] = FS_DIV_PREFERENCE) -> dict:
    """온기(FY)·LTM 메트릭 산출. 각 지표에 _annual, _ltm, _fs 키 생성."""
    out: dict = {}
    for metric, keys in ACCOUNTS.items():
        fy, _, fs_a = _pick(annual_items, keys, fs_pref)          # 온기 = 사업보고서 당기
        ytd_cur, ytd_prev, fs_q = _pick(q_items, keys, fs_pref) if q_items else (None, None, None)
        out[f"{metric}_annual"] = fy
        if fy is not None and ytd_cur is not None and ytd_prev is not None:
            out[f"{metric}_ltm"] = fy - ytd_prev + ytd_cur
        else:
            out[f"{metric}_ltm"] = fy                              # 분기 없으면 온기로 대체
        out[f"{metric}_fs"] = fs_a or fs_q
    return out


def extract_audit(items: list[dict]) -> tuple[str | None, str | None]:
    """(감사의견, 감사인). 응답 다수 행이면 첫 행(최근 사업연도) 사용."""
    if not items:
        return None, None
    it = items[0]
    opinion = it.get("adt_opinion") or it.get("audit_opinion")
    auditor = it.get("adtor") or it.get("auditor")
    return opinion, auditor


# ----------------------------- 단건 수집 -----------------------------
def collect_one(client, corp_code: str) -> dict:
    """한 회사의 재무(온기/LTM) + 감사의견. corp_code 없으면 빈 dict."""
    annual = client.fnltt_singl_acnt(corp_code, ANNUAL_YEAR, REPRT_ANNUAL)
    if not annual and ANNUAL_FALLBACK_YEAR:
        annual = client.fnltt_singl_acnt(corp_code, ANNUAL_FALLBACK_YEAR, REPRT_ANNUAL)
    quarter = client.fnltt_singl_acnt(corp_code, LTM_QUARTER_YEAR, LTM_QUARTER_CODE)
    metrics = compute_metrics(annual, quarter)

    audit_year = ANNUAL_YEAR if annual else ANNUAL_FALLBACK_YEAR
    opinion, auditor = extract_audit(client.audit_opinion(corp_code, audit_year, REPRT_ANNUAL))
    metrics["audit_opinion"] = opinion
    metrics["auditor"] = auditor
    metrics["has_financials"] = any(
        metrics.get(f"{m}_annual") is not None for m in ACCOUNTS
    )
    return metrics


# ----------------------------- 시가총액 (다중 소스 폴백) -----------------------------
def market_caps(date_str: str, verbose: bool = True) -> dict[str, float]:
    """{6자리 종목코드: 시가총액}. 소스 순서대로 시도, 처음 성공분 반환.

    1) KRX 공식 OpenAPI (KRX_API_BUSINESS_KEY, 인가 필요)
    2) FinanceDataReader  — fdr.StockListing('KRX')의 Marcap (무인증, 현재 스냅샷)
    3) pykrx              — get_market_cap_by_ticker (일부 버전은 KRX 로그인 필요)
    """
    # 1) KRX 공식 — 양 시장(KOSPI+KOSDAQ) 다 인가됐을 때만 채택
    try:
        from core.krx_client import KrxClient
        from config import KRX_MIN_COVERAGE
        caps = KrxClient().market_caps(date_str)
        if len(caps) >= KRX_MIN_COVERAGE:
            if verbose:
                print(f"  시총: KRX 공식 OpenAPI {len(caps)}종목 ({date_str})")
            return caps
        if caps and verbose:
            print(f"  [KRX 일부 시장만 인가: {len(caps)}종목 — 코스닥(ksq_bydd_trd) "
                  f"미승인 추정 → fdr 폴백]")
    except Exception:
        pass

    # 2) FinanceDataReader (무인증, 현재 스냅샷 — date_str 무시)
    try:
        import FinanceDataReader as fdr
        df = fdr.StockListing("KRX")
        code_col = next((c for c in ("Code", "Symbol", "종목코드") if c in df.columns), None)
        cap_col = next((c for c in ("Marcap", "시가총액", "MarketCap") if c in df.columns), None)
        if code_col and cap_col:
            caps = {str(c).zfill(6): float(v)
                    for c, v in zip(df[code_col], df[cap_col]) if v == v and v}
            if caps:
                if verbose:
                    print(f"  시총: FinanceDataReader {len(caps)}종목 [현재 스냅샷]")
                return caps
    except Exception as e:
        if verbose:
            print(f"  [fdr 폴백 실패] {e}")

    # 3) pykrx
    try:
        from pykrx import stock
        df = stock.get_market_cap_by_ticker(date_str)
        if df is not None and "시가총액" in df.columns:
            caps = {str(t).zfill(6): float(v) for t, v in df["시가총액"].items() if v == v and v}
            if caps:
                if verbose:
                    print(f"  시총: pykrx {len(caps)}종목 ({date_str})")
                return caps
    except Exception as e:
        if verbose:
            print(f"  [경고] 시총 수집 실패(KRX·fdr·pykrx 모두): {e}")
    return {}


# ----------------------------- 유니버스 수집 (체크포인트) -----------------------------
def collect_universe(df, client, mcap_date: str | None = None,
                     limit: int | None = None, verbose: bool = True):
    """df(유니버스) 각 행의 재무+감사+시총을 수집해 행 dict 리스트 반환. 중단 시 재개.

    df: 최소 STOCK_COL, NAME_COL, FISCAL_MONTH_COL 보유.
    mcap_date: 'YYYYMMDD'. 지정 시 pykrx로 시총 일괄 조인.
    """
    import pandas as pd

    code_map = client.corp_code_map()
    done: dict[str, dict] = {}
    if FIN_CKPT_JSON.exists():
        done = json.loads(FIN_CKPT_JSON.read_text(encoding="utf-8"))
        if verbose:
            print(f"체크포인트 발견 — {len(done)}건 재개")

    caps = {}
    if mcap_date:
        try:
            caps = market_caps(mcap_date)
        except Exception as e:                          # pykrx 실패해도 재무 수집은 진행
            if verbose:
                print(f"[경고] pykrx 시총 수집 실패: {e}")

    rows = df if limit is None else df.head(limit)
    results = []
    for i, (_, row) in enumerate(rows.iterrows(), 1):
        stock_code = str(row[STOCK_COL]).zfill(6)
        base = {
            STOCK_COL: stock_code,
            NAME_COL: row.get(NAME_COL),
            FISCAL_MONTH_COL: row.get(FISCAL_MONTH_COL),
            "corp_code": code_map.get(stock_code),
            "market_cap": caps.get(stock_code),
        }
        if stock_code in done:                          # 이미 수집됨 → 재개
            rec = done[stock_code]
            if caps and rec.get("market_cap") is None:  # 시총만 나중에 백필 가능
                rec["market_cap"] = caps.get(stock_code)
            results.append(rec); continue
        if not base["corp_code"]:
            base["has_financials"] = False
            results.append(base); done[stock_code] = base
        else:
            try:
                base.update(collect_one(client, base["corp_code"]))
            except Exception as e:
                base["error"] = str(e)[:200]
                base["has_financials"] = False
            results.append(base); done[stock_code] = base

        if verbose and i % 50 == 0:
            print(f"  ...수집 {i}/{len(rows)}", flush=True)
        if i % 100 == 0:                                # 주기적 체크포인트 저장
            FIN_CKPT_JSON.write_text(json.dumps(done, ensure_ascii=False), encoding="utf-8")

    FIN_CKPT_JSON.write_text(json.dumps(done, ensure_ascii=False), encoding="utf-8")
    out = pd.DataFrame(results)
    out.to_csv(UNIVERSE_FIN_CSV, index=False, encoding="utf-8-sig")
    if verbose:
        print(f"\n저장: {UNIVERSE_FIN_CSV}  ({len(out)}행)")
    return out