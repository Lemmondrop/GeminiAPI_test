"""
dart_biz_retry.py
────────────────────────────────────────────────────────────
dart_biz_enrichment_cache.json에서 parse_status="no_file"인
항목만 재시도. 전체 재실행 없이 실패분만 개선된 로직으로 보강.

개선 사항:
  - 공백 제거(compact) 키워드 매칭 추가
    (DART 일부 문서는 "주 요  제 품"처럼 글자 사이 공백 삽입됨)
  - 최종 fallback: 키워드 매칭 실패 시
    감사보고서가 아닌 가장 긴 텍스트 파일 선택 (내용 추출 자체는 보장)

실행:
  python dart_biz_retry.py
────────────────────────────────────────────────────────────
"""
import os, sys, json, re, time, zipfile, io
import requests
from pathlib import Path

# ── .env 로드 ─────────────────────────────────────────────
def _parse_val(raw: str) -> str:
    v = raw.strip()
    if v.startswith('"'):
        e = v.find('"', 1); return v[1:e].strip() if e != -1 else v[1:].strip()
    if v.startswith("'"):
        e = v.find("'", 1); return v[1:e].strip() if e != -1 else v[1:].strip()
    for sep in (' #', '\t#'):
        if sep in v: v = v[:v.index(sep)]
    return v.strip().strip('"').strip("'").strip()

def load_env() -> str:
    targets = {"DART_API_BUSINESS_KEY", "DART_API_KEY", "DART_KEY"}
    found = {}
    sd = Path(__file__).parent.resolve()
    for p in [sd / ".." / ".env", sd / ".env"]:
        p = p.resolve()
        if not p.exists(): continue
        with open(p, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line: continue
                k, _, v = line.partition("=")
                k = k.strip()
                if k in targets and k not in found:
                    val = _parse_val(v)
                    if val: found[k] = val
        if found: break
    for t in ("DART_API_BUSINESS_KEY", "DART_API_KEY", "DART_KEY"):
        if found.get(t):
            print(f"[.env] {t} 로드 완료 (길이={len(found[t])})")
            return found[t]
    return ""

KEY = load_env()
if not KEY or len(KEY) != 40:
    sys.exit(f"[ERROR] 키 오류 (len={len(KEY)})")

BASE = "https://opendart.fss.or.kr/api"
DELAY = 0.6
SCRIPT_DIR = Path(__file__).parent          # peer_extraction/scripts/
ROOT_DIR   = SCRIPT_DIR.parent                # peer_extraction/
CACHE_PATH = ROOT_DIR / "data" / "cache" / "dart_biz_enrichment_cache.json"

# ══════════════════════════════════════════════════════════
def _get(url, params, timeout=90, retries=2):
    for attempt in range(retries + 1):
        try:
            return requests.get(url, params=params, timeout=timeout)
        except requests.exceptions.RequestException:
            if attempt < retries: time.sleep(2 ** attempt)
            else: raise

def _quality(text: str) -> float:
    if not text: return 0.0
    kor = len(re.findall(r'[가-힣]', text))
    return (kor / max(len(text), 1)) * min(len(text) / 500, 1.0)

def _despace(s: str) -> str:
    """글자 사이 공백 제거 (DART 일부 문서의 character-spaced 렌더링 대응)."""
    return re.sub(r"\s+", "", s)

def _best_section(clean: str) -> tuple:
    patterns = [
        r"주요\s*제품\s*및\s*서비스(.*?)(?=\n\s*\d+\.\s*[가-힣]|원재료|매출\s*현황|생산\s*설비|\Z)",
        r"주요\s*제품(.*?)(?=\n\s*\d+\.\s*[가-힣]|원재료|매출\s*현황|\Z)",
        r"제품\s*및\s*서비스(.*?)(?=원재료|매출|생산|\Z)",
        r"영업\s*개황(.*?)(?=\d+\.\s*[가-힣]|원재료|\Z)",
    ]
    candidates = []
    for pat in patterns:
        for m in re.finditer(pat, clean, re.DOTALL):
            text = m.group(1)[:3000].strip()
            q = _quality(text)
            candidates.append((text, q))
    if candidates:
        best = max(candidates, key=lambda x: x[1])
        if best[1] > 0.05: return best[0], "section_found"
    for kw in ["사업의 개요", "사업 개요", "회사의 개요"]:
        idx = clean.find(kw)
        if idx != -1:
            snippet = clean[idx:idx+1500].strip()
            if _quality(snippet) > 0.05: return snippet, "overview_fallback"
    return clean[:2000], "head_fallback"

def fetch_biz_summary_v2(rcept_no: str) -> dict:
    """개선판: 공백 제거 매칭 + 최종 fallback(최장 텍스트 파일)."""
    result = {"parse_status": "pending", "biz_summary": "", "biz_detail": ""}
    try:
        r = _get(f"{BASE}/document.xml", {"crtfc_key": KEY, "rcept_no": rcept_no})
        if r.content[:2] != b'PK':
            result["parse_status"] = "not_zip"; return result
        zf = zipfile.ZipFile(io.BytesIO(r.content))
    except Exception as e:
        result["parse_status"] = f"error:{type(e).__name__}"; return result

    names = zf.namelist()
    AUDIT = ["감사보고서","감사의견","audit","Audit"]
    BIZ   = ["주요 제품","사업의 내용","주요제품","사업내용",
             "제품 및 서비스","영업 개황","매출구성"]
    BIZ_COMPACT = [_despace(k) for k in BIZ]   # 공백 제거 버전

    decoded = {}   # 파일명 → (raw_clean_text)
    for n in names:
        if not n.lower().endswith((".htm",".html",".xml")):
            continue
        try:
            raw = zf.read(n).decode("utf-8", errors="ignore")
            txt = re.sub(r"<[^>]+>", " ", raw)
            txt = re.sub(r"\s+", " ", txt)
            decoded[n] = txt
        except: pass

    biz_file = None
    match_mode = ""

    # 1차: 일반 키워드 매칭 (감사보고서 제외)
    for n, txt in decoded.items():
        if any(k in txt[:800] for k in AUDIT): continue
        if any(k in txt for k in BIZ):
            biz_file = n; match_mode = "normal"; break

    # 2차: 공백 제거(compact) 키워드 매칭
    if not biz_file:
        for n, txt in decoded.items():
            if any(k in txt[:800] for k in AUDIT): continue
            txt_compact = _despace(txt)
            if any(k in txt_compact for k in BIZ_COMPACT):
                biz_file = n; match_mode = "compact"; break

    # 3차: 최종 fallback — 감사보고서 아닌 파일 중 가장 긴(품질 높은) 파일
    if not biz_file and decoded:
        non_audit = [(n, txt) for n, txt in decoded.items()
                     if not any(k in txt[:800] for k in AUDIT)]
        pool = non_audit if non_audit else list(decoded.items())
        if pool:
            scored = [(n, _quality(txt[:5000]), len(txt)) for n, txt in pool]
            scored.sort(key=lambda x: (x[1], x[2]), reverse=True)
            biz_file = scored[0][0]
            match_mode = "fallback_longest"

    if not biz_file:
        result["parse_status"] = "no_file"; return result

    clean = decoded[biz_file]
    clean = re.sub(r"\n{3,}", "\n\n", clean).strip()

    section, status = _best_section(clean)
    result["parse_status"]  = status
    result["match_mode"]    = match_mode
    result["target_file"]   = biz_file

    ov = re.search(r"사업(?:의)?\s*개요(.*?)(?=주요\s*제품|원재료|\Z)", clean, re.DOTALL)
    summary = ov.group(1)[:500].strip() if ov else section[:500]
    if _quality(summary) < 0.05: summary = section[:500]
    result["biz_summary"] = summary
    result["biz_detail"]  = section[:1000]
    return result

# ══════════════════════════════════════════════════════════
def main():
    print("="*60)
    print("DART no_file 재시도 — 개선 로직 (공백제거 매칭 + fallback)")
    print("="*60)

    if not CACHE_PATH.exists():
        sys.exit(f"[ERROR] 캐시 파일 없음: {CACHE_PATH}")

    cache = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    targets = {sc: v for sc, v in cache.items()
               if v.get("parse_status") == "no_file" and v.get("rcept_no")}

    print(f"[대상] no_file 재시도: {len(targets)}개\n")

    ok = still_fail = 0
    for idx, (sc, v) in enumerate(targets.items(), 1):
        corp_name = v.get("corp_name", sc)
        rcept_no  = v.get("rcept_no")
        print(f"[{idx:3d}/{len(targets)}]  {corp_name} ({sc})", end="  ")

        try:
            doc = fetch_biz_summary_v2(rcept_no)
            time.sleep(DELAY)
        except Exception as e:
            print(f"오류:{e}")
            still_fail += 1
            continue

        status = doc["parse_status"]
        if status != "no_file":
            cache[sc].update({
                "parse_status": status,
                "biz_summary":  doc["biz_summary"],
                "biz_detail":   doc["biz_detail"],
                "match_mode":   doc.get("match_mode",""),
                "target_file":  doc.get("target_file",""),
            })
            q = _quality(doc["biz_summary"])
            print(f"{status}  mode={doc.get('match_mode','')}  품질={q:.3f}  "
                  f"biz={doc['biz_summary'][:50].replace(chr(10),' ')}...")
            ok += 1
        else:
            print("여전히 no_file")
            still_fail += 1

        # 체크포인트
        if idx % 50 == 0:
            CACHE_PATH.write_text(
                json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n  ── 중간 저장 ({idx}/{len(targets)}) ──\n")

    CACHE_PATH.write_text(
        json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\n{'='*60}")
    print("재시도 완료")
    print(f"  개선됨    : {ok}개")
    print(f"  여전히실패: {still_fail}개")
    print(f"  캐시 갱신: {CACHE_PATH}")

if __name__ == "__main__":
    main()