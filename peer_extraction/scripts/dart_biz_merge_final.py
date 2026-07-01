"""
dart_biz_merge_final.py
────────────────────────────────────────────────────────────
DART 보강 캐시(dart_biz_enrichment_cache.json)와
원본 CSV(company_cord_prototype.csv)를 병합하여
Stage 3 LLM 입력용 최종 캐시 파일을 생성합니다.

우선순위:
  1. biz_summary (DART 보강, 감사보고서 오염 아닌 경우)
  2. 주요제품 (원본 CSV 컬럼, fallback)

실행:
  python dart_biz_merge_final.py

출력:
  stage3_business_text_final.json
  { "종목코드": {"corp_name":, "business_text":, "source":, "char_len":} }
────────────────────────────────────────────────────────────
"""
import json, csv, re, sys
from pathlib import Path
from collections import Counter

SCRIPT_DIR = Path(__file__).parent          # peer_extraction/scripts/
ROOT_DIR   = SCRIPT_DIR.parent                # peer_extraction/
CACHE_PATH = ROOT_DIR / "data" / "cache" / "dart_biz_enrichment_cache.json"
CSV_PATH   = ROOT_DIR.parent / "data" / "metadata" / "company_cord_prototype.csv"
OUT_PATH   = ROOT_DIR / "data" / "cache" / "stage3_business_text_final.json"

# ── 감사보고서 오염 판정 ────────────────────────────────────
JUNK_MARKERS = ["감사보고서 6.0", "연결감사보고서 6.0",
                "100000000000", "사업보고서제출기한연장신고서"]

def is_junk(summary: str) -> bool:
    if not summary or len(summary.strip()) < 10:
        return True
    return any(m in summary[:60] for m in JUNK_MARKERS)

# ── 로드 ────────────────────────────────────────────────────
def load_cache() -> dict:
    if not CACHE_PATH.exists():
        sys.exit(f"[ERROR] 캐시 없음: {CACHE_PATH}")
    return json.loads(CACHE_PATH.read_text(encoding="utf-8"))

def load_csv_baseline() -> dict:
    """종목코드 → 주요제품(원본) 매핑."""
    for enc in ("cp949", "utf-8", "utf-8-sig"):
        try:
            with open(CSV_PATH, encoding=enc, newline="") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
            break
        except (UnicodeDecodeError, FileNotFoundError):
            continue
    else:
        sys.exit(f"[ERROR] CSV 읽기 실패: {CSV_PATH}")

    cols = list(rows[0].keys())
    stock_col = next((c for c in ["종목코드","stock_code"] if c in cols), cols[0])
    name_col  = next((c for c in ["회사명","corp_name"] if c in cols), cols[1])
    prod_col  = next((c for c in ["주요제품","주요 제품"] if c in cols), "")

    baseline = {}
    for r in rows:
        sc = str(r[stock_col]).zfill(6)
        if sc not in baseline:   # 중복 행은 첫 번째만
            baseline[sc] = {
                "corp_name": r.get(name_col, sc),
                "주요제품":  r.get(prod_col, "") if prod_col else "",
            }
    return baseline

# ── 병합 ────────────────────────────────────────────────────
def main():
    print("="*60)
    print("Stage 3 최종 입력 텍스트 병합")
    print("="*60)

    cache    = load_cache()
    baseline = load_csv_baseline()
    print(f"[로드] DART 캐시: {len(cache)}개  /  CSV 베이스라인: {len(baseline)}개\n")

    final = {}
    source_count = Counter()
    len_before, len_after = [], []

    all_codes = set(cache.keys()) | set(baseline.keys())

    for sc in all_codes:
        c = cache.get(sc, {})
        b = baseline.get(sc, {})

        corp_name = c.get("corp_name") or b.get("corp_name") or sc
        original_text = (b.get("주요제품") or "").strip()

        dart_summary = c.get("biz_summary", "").strip()
        dart_status  = c.get("parse_status", "")

        # 우선순위 판정
        if dart_summary and dart_status in ("section_found", "overview_fallback") \
           and not is_junk(dart_summary):
            business_text = dart_summary
            source = "dart_enriched"
        elif dart_summary and not is_junk(dart_summary) and len(dart_summary) > len(original_text):
            business_text = dart_summary
            source = "dart_enriched_fallback"
        else:
            business_text = original_text
            source = "csv_original"

        final[sc] = {
            "corp_name":     corp_name,
            "business_text": business_text,
            "source":        source,
            "char_len":      len(business_text),
        }
        source_count[source] += 1
        len_before.append(len(original_text))
        len_after.append(len(business_text))

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(
        json.dumps(final, ensure_ascii=False, indent=2), encoding="utf-8")

    # ── 통계 리포트 ───────────────────────────────────────
    n = len(final)
    avg_before = sum(len_before) / n if n else 0
    avg_after  = sum(len_after)  / n if n else 0
    improvement_pct = (avg_after / avg_before - 1) * 100 if avg_before > 0 else 0

    enriched_n = source_count["dart_enriched"] + source_count["dart_enriched_fallback"]
    enriched_pct = enriched_n / n * 100 if n else 0

    print(f"=== 병합 결과 ===")
    print(f"  최종 기업 수        : {n:,}개")
    print(f"  DART 보강 적용      : {enriched_n:,}개  ({enriched_pct:.1f}%)")
    print(f"  CSV 원본 유지       : {source_count['csv_original']:,}개  "
          f"({source_count['csv_original']/n*100:.1f}%)")
    print()
    print(f"=== 텍스트 길이 비교 (전체 평균) ===")
    print(f"  병합 전 (CSV 주요제품) : {avg_before:.1f}자")
    print(f"  병합 후 (최종)         : {avg_after:.1f}자")
    print(f"  평균 길이 개선율       : +{improvement_pct:.0f}%")
    print()

    # 보강 성공 기업만 별도 비교 (실질 개선 체감)
    enriched_lens = [v["char_len"] for v in final.values()
                     if v["source"].startswith("dart_enriched")]
    if enriched_lens:
        avg_enriched = sum(enriched_lens) / len(enriched_lens)
        print(f"=== DART 보강 성공 기업만 (n={len(enriched_lens):,}) ===")
        print(f"  평균 텍스트 길이 : {avg_enriched:.1f}자")
        print(f"  (원본 평균 19자 대비 약 {avg_enriched/19:.1f}배)")

    print(f"\n  저장: {OUT_PATH}")

if __name__ == "__main__":
    main()