"""KRX 시총 진단 — 왜 0종목인지 원시 응답으로 원인 파악.

실행: peer_extraction 루트에서  python tests/diag_krx.py [YYYYMMDD]
  (날짜 생략 시 20260616. 휴장일이면 직전 영업일로 바꿔서 재실행)

확인 포인트:
  - HTTP status (200=정상 / 4xx=인증·권한 / 그 외)
  - 응답 본문 (인증 실패·미구독 시 에러 메시지가 여기 보임)
  - OutBlock_1 행수 + 실제 필드명 (ISU_SRT_CD / MKTCAP 가 맞는지 확인)
"""
from __future__ import annotations
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config


def main():
    import requests
    key = (os.environ.get(config.KRX_API_KEY_ENV) or "").strip()
    print(f"{config.KRX_API_KEY_ENV} 로드: {'OK (' + str(len(key)) + '자)' if key else '없음 ❌'}")
    if not key:
        print("→ .env에 KRX_API_BUSINESS_KEY 없음. config._load_env / 키 이름 확인.")
        return

    date = sys.argv[1] if len(sys.argv) > 1 else "20260616"
    ep = config.KRX_BASE + config.KRX_DAILY_TRADE["KOSPI"]
    print(f"\nGET {ep}\n  params basDd={date}  header AUTH_KEY=***{key[-4:]}")

    try:
        r = requests.get(ep, params={"basDd": date},
                         headers={"AUTH_KEY": key}, timeout=30)
    except Exception as e:
        print("요청 자체 실패:", e); return

    print("HTTP status:", r.status_code)
    print("Content-Type:", r.headers.get("Content-Type"))
    print("응답 본문(앞 1000자):")
    print(r.text[:1000])

    try:
        d = r.json()
    except Exception as e:
        print("\nJSON 파싱 실패:", e); return

    ob = d.get("OutBlock_1")
    if ob:
        print(f"\n✅ OutBlock_1 {len(ob)}행")
        print("첫 행 전체 키:", list(ob[0].keys()))
        sample = {k: ob[0][k] for k in list(ob[0])[:10]}
        print("첫 행 샘플:", sample)
        print("\n→ 시총 필드가 'MKTCAP'이 맞는지, 종목코드가 'ISU_SRT_CD'인지 위 키에서 확인.")
    else:
        print("\n❌ OutBlock_1 없음. 응답 최상위 키:", list(d.keys()))
        print("→ 본문에 에러 메시지가 있으면 미구독/인증 문제. KRX 마이페이지에서")
        print("   '유가증권 일별매매정보(stk_bydd_trd)' API 사용 신청이 됐는지 확인.")


if __name__ == "__main__":
    main()
