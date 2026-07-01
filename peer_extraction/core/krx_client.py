"""KRX Data Marketplace OpenAPI 클라이언트.

일별매매정보(stk_bydd_trd/ksq_bydd_trd/knx_bydd_trd)의 MKTCAP(시가총액)을 종목코드별로 수집.
인증: 헤더 AUTH_KEY = .env의 KRX_API_BUSINESS_KEY.  파라미터: basDd(YYYYMMDD).

시장당 1콜이면 해당 시장 전 종목 시총이 한 번에 온다(유니버스 3종목씩이 아니라 전체).
"""
from __future__ import annotations
import os

from config import KRX_API_KEY_ENV, KRX_BASE, KRX_DAILY_TRADE


class KrxClient:
    def __init__(self, api_key: str | None = None):
        self.key = (api_key or os.environ.get(KRX_API_KEY_ENV) or "").strip()
        if not self.key:
            raise RuntimeError(f"{KRX_API_KEY_ENV} 없음 — 상위 .env에 {KRX_API_KEY_ENV}=... 설정 필요")

    def market_caps(self, date_str: str) -> dict[str, float]:
        """{6자리 종목코드(ISU_SRT_CD): 시가총액(MKTCAP, 원)}. 기준일 basDd=YYYYMMDD."""
        import requests
        caps: dict[str, float] = {}
        for market, ep in KRX_DAILY_TRADE.items():
            try:
                r = requests.get(
                    KRX_BASE + ep,
                    params={"basDd": date_str},
                    headers={"AUTH_KEY": self.key},
                    timeout=30,
                )
                r.raise_for_status()
                rows = r.json().get("OutBlock_1", []) or []
            except Exception:
                continue  # 해당 시장 엔드포인트 없거나 실패 시 건너뜀 (예: KONEX)
            for row in rows:
                raw = (row.get("ISU_SRT_CD") or row.get("ISU_CD") or "").strip()
                cap = row.get("MKTCAP")
                if not raw or cap in (None, ""):
                    continue
                try:
                    caps[raw.zfill(6)] = float(str(cap).replace(",", ""))
                except ValueError:
                    pass
        return caps