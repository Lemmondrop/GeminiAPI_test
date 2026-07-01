"""OpenDART 저수준 클라이언트.

- corp_code_map(): corpCode.xml(zip) 1회 다운로드 → {종목코드: corp_code} 캐시.
- fnltt_singl_acnt(): 단일회사 주요계정 (영업이익/당기순이익/매출 등).
- audit_opinion(): 회계감사인의 명칭 및 감사의견.
- fetch_document(): 공시 원본문서(증권신고서 등) 본문 텍스트 (document.xml zip → 태그제거 → 캐시).

상태코드 처리: '000' 정상 / '013' 데이터없음(빈 리스트) / '020' 요청한도초과(대기 후 재시도).
인증키 env 변수명은 config.DART_API_KEY_ENV(기본 'DART_API_BUSINESS_KEY')에서 자동 로드.
"""
from __future__ import annotations
import io
import os
import time
import zipfile
import xml.etree.ElementTree as ET
from collections import deque

import html
import re

from config import (DART_API_KEY_ENV, DART_BASE, DART_CORPCODE_CACHE,
                    DART_DOC_CACHE_DIR, DART_MAX_PER_MIN, DART_MAX_RETRIES)


class DartError(Exception):
    pass


def strip_markup(raw: str) -> str:
    """공시 원본 XML/HTML → 평문. 모든 태그를 공백으로(표 셀 병합 방지) + 엔티티 복원."""
    s = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", raw)
    s = re.sub(r"(?s)<[^>]+>", " ", s)               # 태그 → 공백
    s = html.unescape(s)
    s = re.sub(r"[ \t\u00a0]+", " ", s)
    s = re.sub(r"\n\s*\n+", "\n", s)
    return s.strip()


class _RateLimiter:
    """슬라이딩 윈도우 분당 호출 제한."""

    def __init__(self, max_per_min: int):
        self.max = max_per_min
        self.win: deque[float] = deque()

    def acquire(self, n: int = 1) -> None:
        now = time.monotonic()
        while self.win and now - self.win[0] > 60:
            self.win.popleft()
        while self.win and len(self.win) + n > self.max:
            time.sleep(60 - (now - self.win[0]) + 0.2)
            now = time.monotonic()
            while self.win and now - self.win[0] > 60:
                self.win.popleft()
        self.win.extend([now] * n)


class DartClient:
    def __init__(self, api_key: str | None = None,
                 max_per_min: int = DART_MAX_PER_MIN,
                 max_retries: int = DART_MAX_RETRIES):
        self.key = api_key or os.environ.get(DART_API_KEY_ENV)
        if not self.key:
            raise RuntimeError(f"{DART_API_KEY_ENV} 없음 — 상위 .env에 {DART_API_KEY_ENV}=... 설정 필요")
        self.max_retries = max_retries
        self._limiter = _RateLimiter(max_per_min)

    # ----------------------------- 공통 요청 -----------------------------
    def _request_json(self, path: str, as_list: bool = True, **params) -> list[dict] | dict:
        import requests
        params["crtfc_key"] = self.key
        url = f"{DART_BASE}/{path}"
        last: Exception | None = None
        for attempt in range(self.max_retries):
            self._limiter.acquire(1)
            try:
                r = requests.get(url, params=params, timeout=30)
                data = r.json()
            except Exception as e:                       # 네트워크/파싱 오류 → 백오프 재시도
                last = e
                time.sleep(2 * (attempt + 1))
                continue
            status = data.get("status")
            if status == "000":
                return data if not as_list else (data.get("list", []) or [])
            if status == "013":                          # 조회된 데이터 없음 (정상 케이스)
                return {} if not as_list else []
            if status == "020":                          # 요청 한도 초과 → 대기 후 재시도
                last = DartError("020 요청 제한 초과")
                time.sleep(60)
                continue
            raise DartError(f"{path}: status={status}, msg={data.get('message')}")
        raise last if last else DartError(f"{path}: 재시도 한도 초과")

    def company_profile(self, corp_code: str) -> dict:
        """기업개황(company.json). induty_code(표준산업분류코드), corp_name 등 flat dict."""
        return self._request_json("company.json", as_list=False, corp_code=corp_code) or {}

    # ----------------------------- corp_code 매핑 -----------------------------
    @staticmethod
    def _norm_name(v) -> str:
        s = str(v or "").strip()
        for t in ("주식회사", "(주)", "㈜", " "):
            s = s.replace(t, "")
        return s.lower()

    def corp_code_by_name(self, refresh: bool = False) -> dict[str, str]:
        """{정규화 회사명: 8자리 corp_code}. 상장 전 IPO사도 포함(corpCode.xml 전체)."""
        import pandas as pd
        cache = DART_CORPCODE_CACHE.with_name("dart_corpname_map.csv")
        if not refresh and cache.exists():
            df = pd.read_csv(cache, dtype=str)
            return dict(zip(df["norm_name"], df["corp_code"]))

        rows = self._corpcode_rows()                 # (corp_name, corp_code, stock_code)
        out, seen = [], set()
        for name, cc, _sc in rows:
            n = self._norm_name(name)
            if n and cc and n not in seen:           # 동명이인은 첫 건만(상장사 우선 정렬 후)
                seen.add(n)
                out.append((n, cc))
        df = pd.DataFrame(out, columns=["norm_name", "corp_code"])
        df.to_csv(cache, index=False)
        return dict(zip(df["norm_name"], df["corp_code"]))

    def _corpcode_rows(self) -> list[tuple[str, str, str]]:
        """corpCode.xml → [(corp_name, corp_code, stock_code)]. 상장사(종목코드 보유) 우선 정렬."""
        import requests
        self._limiter.acquire(1)
        r = requests.get(f"{DART_BASE}/corpCode.xml",
                         params={"crtfc_key": self.key}, timeout=60)
        zf = zipfile.ZipFile(io.BytesIO(r.content))
        root = ET.fromstring(zf.read(zf.namelist()[0]).decode("utf-8"))
        rows = []
        for el in root.iter("list"):
            cc = (el.findtext("corp_code") or "").strip()
            nm = (el.findtext("corp_name") or "").strip()
            sc = (el.findtext("stock_code") or "").strip()
            if cc and nm:
                rows.append((nm, cc, sc))
        rows.sort(key=lambda x: x[2] == "")          # 종목코드 있는(상장) 회사 먼저
        return rows

    def corp_code_map(self, refresh: bool = False) -> dict[str, str]:
        """{6자리 종목코드: 8자리 corp_code}. corpCode.xml 1회 다운로드 후 CSV 캐시."""
        import pandas as pd
        import requests
        if not refresh and DART_CORPCODE_CACHE.exists():
            df = pd.read_csv(DART_CORPCODE_CACHE, dtype=str)
            return dict(zip(df["stock_code"], df["corp_code"]))

        self._limiter.acquire(1)
        r = requests.get(f"{DART_BASE}/corpCode.xml",
                         params={"crtfc_key": self.key}, timeout=60)
        zf = zipfile.ZipFile(io.BytesIO(r.content))
        xml = zf.read(zf.namelist()[0]).decode("utf-8")
        root = ET.fromstring(xml)

        rows = []
        for el in root.iter("list"):
            sc = (el.findtext("stock_code") or "").strip()
            cc = (el.findtext("corp_code") or "").strip()
            if sc and cc:                                # 종목코드 있는 상장사만
                rows.append((sc.zfill(6), cc))
        df = pd.DataFrame(rows, columns=["stock_code", "corp_code"]).drop_duplicates("stock_code")
        df.to_csv(DART_CORPCODE_CACHE, index=False)
        return dict(zip(df["stock_code"], df["corp_code"]))

    # ----------------------------- 엔드포인트 -----------------------------
    def fnltt_singl_acnt(self, corp_code: str, bsns_year: str, reprt_code: str) -> list[dict]:
        """단일회사 주요계정 (재무상태표+손익계산서 주요계정)."""
        return self._request_json(
            "fnlttSinglAcnt.json",
            corp_code=corp_code, bsns_year=bsns_year, reprt_code=reprt_code,
        )

    def audit_opinion(self, corp_code: str, bsns_year: str, reprt_code: str) -> list[dict]:
        """회계감사인의 명칭 및 감사의견."""
        return self._request_json(
            "accnutAdtorNmNdAdtOpinion.json",
            corp_code=corp_code, bsns_year=bsns_year, reprt_code=reprt_code,
        )

    # ----------------------------- 공시 원본문서 -----------------------------
    def list_filings(self, corp_code: str, bgn_de: str = "20100101",
                     end_de: str | None = None, pblntf_ty: str = "C") -> list[dict]:
        """공시목록(list.json). pblntf_ty='C'=발행공시(증권신고서 등). 접수일 오름차순.

        반환 행: report_nm(보고서명), rcept_no(접수번호), rcept_dt(접수일) 등.
        """
        import datetime
        end_de = end_de or datetime.date.today().strftime("%Y%m%d")
        out: list[dict] = []
        page = 1
        while True:
            rows = self._request_json(
                "list.json", corp_code=corp_code, bgn_de=bgn_de, end_de=end_de,
                pblntf_ty=pblntf_ty, page_no=page, page_count=100,
            )
            if not rows:
                break
            out.extend(rows)
            if len(rows) < 100:
                break
            page += 1
            if page > 20:                            # 안전 상한
                break
        out.sort(key=lambda r: r.get("rcept_dt", ""))
        return out

    def find_initial_registration(self, corp_code: str) -> dict | None:
        """비교회사 표가 담긴 '최초 증권신고서(지분증권)' 1건 탐색.

        발행조건확정/정정 본이 아니라 최초 제출본을 우선. 못 찾으면 가장 이른 증권신고서.
        """
        filings = self.list_filings(corp_code, pblntf_ty="C")
        regs = [f for f in filings if "증권신고서" in (f.get("report_nm") or "")]
        if not regs:
            return None
        # 최초 제출본 우선: 정정/발행조건확정 표기가 없는 가장 이른 건
        plain = [f for f in regs
                 if not any(t in (f.get("report_nm") or "")
                            for t in ("정정", "발행조건확정", "철회"))]
        return (plain or regs)[0]                     # rcept_dt 오름차순이라 [0]=최초

    def fetch_document(self, rcept_no: str, refresh: bool = False) -> str:
        """공시 원본문서(증권신고서 등) 본문 평문. document.xml(zip) → 태그제거 → 캐시.

        DART document API는 PDF가 아니라 원본 XML/HTML(ZIP)을 반환 → 태그 제거 후 텍스트화.
        rcept_no별로 캐시(.txt)하여 재실행 시 재다운로드 안 함.
        """
        import requests
        cache = DART_DOC_CACHE_DIR / f"{rcept_no}.txt"
        if not refresh and cache.exists():
            return cache.read_text(encoding="utf-8")

        self._limiter.acquire(1)
        r = requests.get(f"{DART_BASE}/document.xml",
                         params={"crtfc_key": self.key, "rcept_no": rcept_no}, timeout=60)
        content = r.content
        if content[:2] != b"PK":                     # ZIP 아님 → 에러 응답(XML/JSON)
            txt = content.decode("utf-8", "replace")
            m = (re.search(r"<status>(\d+)</status>", txt)
                 or re.search(r'"status"\s*:\s*"(\d+)"', txt))
            raise DartError(f"document {rcept_no}: status={m.group(1) if m else '?'} "
                            f"(키 미인가/접수번호 오류/데이터없음 등)")

        zf = zipfile.ZipFile(io.BytesIO(content))
        parts = []
        for name in zf.namelist():                   # 보통 1개(메인 문서), 첨부 포함 가능
            b = zf.read(name)
            for enc in ("utf-8", "euc-kr", "cp949"):  # 공시원본은 euc-kr이 흔함
                try:
                    parts.append(b.decode(enc))
                    break
                except UnicodeDecodeError:
                    continue
            else:
                parts.append(b.decode("utf-8", "replace"))
        text = strip_markup("\n".join(parts))

        DART_DOC_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache.write_text(text, encoding="utf-8")
        return text