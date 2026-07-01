"""
test_business_text.py
────────────────────────────────────────────────────────────
stage3.py 패치(business_text 함수)가 DART 보강 캐시를
정상적으로 읽어오는지 확인하는 최소 스모크 테스트.

전체 파이프라인(Gemini 호출 등) 없이 함수 단위로 빠르게 검증.

실행 위치: peer_extraction/  (stages/, stage3_business_text_final.json 와 같은 레벨)
실행:
  python test_business_text.py
────────────────────────────────────────────────────────────
"""
import sys
import pandas as pd
from pathlib import Path

ROOT_DIR = Path(__file__).parent.parent                   # peer_extraction/
sys.path.insert(0, str(ROOT_DIR))                          # config.py 등 루트 모듈

from stages.stage3_business import business_text, _BIZ_CACHE_PATH, _load_biz_cache
from config import STOCK_COL

print("="*60)
print("business_text() 패치 검증")
print("="*60)

# ── 캐시 파일 경로/존재 확인 ────────────────────────────────
print(f"\n[캐시 경로] {_BIZ_CACHE_PATH}")
print(f"[존재 여부] {_BIZ_CACHE_PATH.exists()}")

cache = _load_biz_cache()
print(f"[로드된 엔트리 수] {len(cache)}개")

if not cache:
    print("\n⚠ 캐시가 비어있습니다. stage3_business_text_final.json 경로를 확인하세요.")
    sys.exit(1)

# ── 검증 대상 종목 (이전 검증에서 사용한 샘플) ──────────────
TEST_STOCKS = {
    "053350": "이니텍 (DART 보강 성공 케이스)",
    "041460": "한국전자인증 (DART 보강 성공 케이스)",
    "999999": "존재하지 않는 종목 (fallback 테스트)",
}

print(f"\n{'─'*55}")
for stock_code, desc in TEST_STOCKS.items():
    row = pd.Series({STOCK_COL: stock_code, "주요제품": "(테스트용 원본 텍스트)", "업종": "테스트업종"})
    text = business_text(row)

    in_cache = stock_code.zfill(6) in cache
    source = "DART 보강" if (in_cache and len(text) >= 10 and text != "(테스트용 원본 텍스트) | 테스트업종") else "fallback(원본)"

    print(f"\n▶ {desc}")
    print(f"  종목코드     : {stock_code}")
    print(f"  캐시 존재    : {in_cache}")
    print(f"  추정 소스    : {source}")
    print(f"  텍스트 길이  : {len(text)}자")
    print(f"  앞 100자     : {text[:100]}")

print(f"\n{'='*60}")
print("판정 기준")
print("  이니텍/한국전자인증 → 텍스트 길이 200자 이상이면 DART 보강 정상 적용 ✅")
print("  존재하지 않는 종목   → '(테스트용 원본 텍스트) | 테스트업종' 그대로 나오면 fallback 정상 ✅")