"""
dart_biz_batch.py
────────────────────────────────────────────────────────────
DART 사업보고서 배치 수집 스크립트
- company_cord_prototype.csv의 전체 상장사(~3,281개)에 대해
  사업 개요(biz_summary)를 수집하여 캐시 파일로 저장
- DART_API_BUSINESS_KEY 단일 키 사용
- 체크포인트 100개마다 저장 (중단/재개 지원)

실행:
  python dart_biz_batch.py                  # 전체 실행
  python dart_biz_batch.py --limit 50       # 50개만 테스트
  python dart_biz_batch.py --resume         # 이전 중단 지점부터 재개

출력:
  dart_biz_enrichment_cache.json
────────────────────────────────────────────────────────────
"""

import os, sys, json, re, time, zipfile, io, datetime, argparse, csv
import requests
from pathlib import Path

# ── 인자 파싱 ────────────────────────────────────────────
parser = argparse.ArgumentParser()
parser.add_argument("--limit",  type=int, default=0,
                    help="처리할 최대 기업 수 (0=전체)")
parser.add_argument("--resume", action="store_true",
                    help="기존 캐시를 유지하고 미처리 기업만 수집")
parser.add_argument("--delay",  type=float, default=0.6,
                    help="API 호출 간 딜레이 (초, 기본 0.6)")
args = parser.parse_args()

DELAY     = args.delay
CKPT_EVERY = 100   # N개마다 체크포인트 저장

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
    # 우선순위: BUSINESS_KEY > API_KEY > KEY
    for t in ("DART_API_BUSINESS_KEY", "DART_API_KEY", "DART_KEY"):
        if found.get(t):
            print(f"[.env] {t} 로드 완료 (길이={len(found[t])})")
            return found[t]
        v = os.environ.get(t, "")
        if v:
            print(f"[env] {t} 로드 완료 (환경변수)")
            return v
    return ""

KEY = load_env()
if not KEY or len(KEY) != 40:
    sys.exit(f"[ERROR] DART_API_BUSINESS_KEY 없음 또는 길이 오류 (len={len(KEY)})")

BASE = "https://opendart.fss.or.kr/api"

# ── 경로 설정 ─────────────────────────────────────────────
# config.py가 정의한 경로를 그대로 재사용 (CSV·corp_code 캐시 중복 방지).
SCRIPT_DIR = Path(__file__).parent          # peer_extraction/scripts/
ROOT_DIR   = SCRIPT_DIR.parent                # peer_extraction/
sys.path.insert(0, str(ROOT_DIR))
import config  # noqa: E402  (peer_extraction 루트의 config.py)

CSV_PATH           = config.LISTED_UNIVERSE_CSV       # 상장사+KSIC 원본 CSV (공유)
CACHE_DIR           = config.CACHE_DIR                 # peer_extraction/data/cache
CACHE_PATH          = CACHE_DIR / "dart_biz_enrichment_cache.json"
DART_CORPCODE_CACHE = config.DART_CORPCODE_CACHE       # stock_code → corp_code 맵 (공유 캐시, CSV)

# ══════════════════════════════════════════════════════════
# Step 0. stock_code → corp_code 맵 구축 (1회)
# ══════════════════════════════════════════════════════════
def build_corpcode_map(stock_codes: set) -> dict:
    """corpCode.xml 다운로드 → {stock_code: corp_code} 딕셔너리.

    config.DART_CORPCODE_CACHE(CSV)를 공유 캐시로 사용 — 다른 모듈
    (재무 수집 등)과 동일 파일을 공유해 corpCode.xml 중복 다운로드 방지.
    CSV 스키마: stock_code, corp_code, corp_name
    """
    # ── 로컬 캐시 있으면 재사용 ─────────────────────────────
    if DART_CORPCODE_CACHE.exists():
        cached = {}
        with open(DART_CORPCODE_CACHE, encoding="utf-8", newline="") as f:
            for row in csv.DictReader(f):
                cached[row["stock_code"]] = row["corp_code"]
        missing = stock_codes - set(cached.keys())
        if not missing:
            print(f"[corpCode] 공유 캐시 사용 ({len(cached)}개) — {DART_CORPCODE_CACHE}")
            return cached
        print(f"[corpCode] 캐시에 없는 종목 {len(missing)}개 — 재다운로드")

    print("[corpCode] DART corpCode.xml 다운로드 중...")
    r = requests.get(f"{BASE}/corpCode.xml",
                     params={"crtfc_key": KEY}, timeout=60)
    with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
        xml = zf.read(zf.namelist()[0]).decode("utf-8", errors="ignore")

    rows = []
    mapping = {}
    for block in re.findall(r"<list>(.*?)</list>", xml, re.DOTALL):
        sc = re.search(r"<stock_code>(.*?)</stock_code>", block)
        cc = re.search(r"<corp_code>(.*?)</corp_code>",  block)
        nm = re.search(r"<corp_name>(.*?)</corp_name>",  block)
        sc_val = sc.group(1).strip() if sc else ""
        cc_val = cc.group(1).strip() if cc else ""
        if sc_val and cc_val:                 # 상장사(종목코드 있는 행)만 저장
            mapping[sc_val] = cc_val
            rows.append({
                "stock_code": sc_val,
                "corp_code":  cc_val,
                "corp_name":  nm.group(1).strip() if nm else "",
            })

    # ── 공유 캐시(CSV)로 저장 — 다른 모듈도 동일 파일 재사용 ──
    DART_CORPCODE_CACHE.parent.mkdir(parents=True, exist_ok=True)
    with open(DART_CORPCODE_CACHE, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["stock_code", "corp_code", "corp_name"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"[corpCode] {len(mapping)}개 매핑 저장 → {DART_CORPCODE_CACHE}")
    return mapping

# ══════════════════════════════════════════════════════════
# Step 1. CSV 로드
# ══════════════════════════════════════════════════════════
def load_companies() -> list[dict]:
    """company_cord_prototype.csv → [{stock_code, corp_name, ksic_code, ...}]"""
    rows = []
    # cp949 인코딩 (기존 확인된 인코딩)
    for enc in ("cp949", "utf-8", "utf-8-sig"):
        try:
            with open(CSV_PATH, encoding=enc, newline="") as f:
                reader = csv.DictReader(f)
                rows = [dict(r) for r in reader]
            print(f"[CSV] {len(rows)}개 로드 (encoding={enc})")
            return rows
        except (UnicodeDecodeError, FileNotFoundError):
            continue
    sys.exit(f"[ERROR] {CSV_PATH} 읽기 실패")

# ══════════════════════════════════════════════════════════
# DART API 호출 함수들 (재시도 포함)
# ══════════════════════════════════════════════════════════
def _get(url, params, timeout=30, retries=2) -> dict | bytes:
    for attempt in range(retries + 1):
        try:
            r = requests.get(url, params=params, timeout=timeout)
            return r
        except requests.exceptions.RequestException as e:
            if attempt < retries:
                time.sleep(2 ** attempt)
            else:
                raise
    raise RuntimeError("요청 실패")

def fetch_rcept_no(corp_code: str) -> tuple:
    today  = datetime.date.today().strftime("%Y%m%d")
    bgn_de = (datetime.date.today()-datetime.timedelta(days=365*4)).strftime("%Y%m%d")
    r = _get(f"{BASE}/list.json",
             {"crtfc_key": KEY, "corp_code": corp_code,
              "pblntf_ty": "A", "bgn_de": bgn_de,
              "end_de": today, "page_count": 10})
    d = r.json()
    if d.get("status") != "000": return None, None
    items = d.get("list", [])
    # 원본 사업보고서 우선
    for it in items:
        nm, rno = it.get("report_nm",""), it.get("rcept_no","")
        if "사업보고서" in nm and rno and not nm.startswith("["):
            return rno, it.get("rcept_dt","")
    # 정정본 허용
    for it in items:
        nm, rno = it.get("report_nm",""), it.get("rcept_no","")
        if "사업보고서" in nm and rno:
            return rno, it.get("rcept_dt","")
    return None, None

def fetch_revenue(corp_code: str) -> float:
    for yr in ("2024","2023","2022"):
        for fs in ("CFS","OFS"):
            try:
                r = _get(f"{BASE}/fnlttSinglAcntAll.json",
                         {"crtfc_key": KEY, "corp_code": corp_code,
                          "bsns_year": yr, "reprt_code": "11011", "fs_div": fs})
                d = r.json(); time.sleep(DELAY*0.5)
                if d.get("status") == "000":
                    for it in d.get("list", []):
                        if it.get("account_nm","") in ("매출액","수익(매출액)"):
                            v = int(it.get("thstrm_amount","0").replace(",",""))
                            return round(v / 1e8, 1)
            except: pass
    return 0.0

def _quality(text: str) -> float:
    if not text: return 0.0
    kor = len(re.findall(r'[가-힣]', text))
    return (kor / max(len(text), 1)) * min(len(text) / 500, 1.0)

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
    # 사업 개요 대체
    for kw in ["사업의 개요", "사업 개요", "회사의 개요"]:
        idx = clean.find(kw)
        if idx != -1:
            snippet = clean[idx:idx+1500].strip()
            if _quality(snippet) > 0.05: return snippet, "overview_fallback"
    return clean[:2000], "head_fallback"

def fetch_biz_summary(rcept_no: str) -> dict:
    """document.xml 다운로드 → biz_summary 추출."""
    result = {"parse_status": "pending", "biz_summary": "", "biz_detail": ""}
    try:
        r = _get(f"{BASE}/document.xml",
                 {"crtfc_key": KEY, "rcept_no": rcept_no}, timeout=90)
        if r.content[:2] != b'PK':
            result["parse_status"] = f"not_zip"
            return result
        zf = zipfile.ZipFile(io.BytesIO(r.content))
    except Exception as e:
        result["parse_status"] = f"error:{type(e).__name__}"
        return result

    names = zf.namelist()
    AUDIT = ["감사보고서","감사의견","audit","Audit"]
    BIZ   = ["주요 제품","사업의 내용","주요제품","사업내용",
             "제품 및 서비스","영업 개황","매출구성"]

    biz_file = None

    # htm 우선
    for n in names:
        if not n.lower().endswith((".htm",".html")): continue
        try:
            raw = zf.read(n).decode("utf-8",errors="ignore")
            txt = re.sub(r"<[^>]+>"," ",raw)
            if not any(k in txt[:600] for k in AUDIT) and any(k in txt for k in BIZ):
                biz_file = n; break
        except: pass

    # xml fallback
    if not biz_file:
        xmls = [n for n in names if n.lower().endswith(".xml")]
        scored = []
        for n in xmls:
            try:
                raw = zf.read(n).decode("utf-8",errors="ignore")
                txt = re.sub(r"<[^>]+>"," ",raw)
                txt = re.sub(r"\s+"," ",txt)
                if any(k in txt[:800] for k in AUDIT): continue
                biz_hit = sum(1 for k in BIZ if k in txt)
                q = _quality(txt[:5000])
                scored.append((n, biz_hit, q, txt))
            except: pass
        if scored:
            scored.sort(key=lambda x:(x[1],x[2]),reverse=True)
            biz_file = scored[0][0]

    if not biz_file:
        result["parse_status"] = "no_file"; return result

    raw = zf.read(biz_file).decode("utf-8",errors="ignore")
    clean = re.sub(r"<[^>]+>"," ",raw)
    clean = re.sub(r"[ \t]+"," ",clean)
    clean = re.sub(r"\n{3,}","\n\n",clean).strip()

    section, status = _best_section(clean)
    result["parse_status"] = status

    # biz_summary: 개요 앞 500자
    ov = re.search(r"사업(?:의)?\s*개요(.*?)(?=주요\s*제품|원재료|\Z)",
                   clean, re.DOTALL)
    summary = ov.group(1)[:500].strip() if ov else section[:500]
    if _quality(summary) < 0.05: summary = section[:500]
    result["biz_summary"] = summary

    # biz_detail: 섹션 전체 앞 1000자 (Stage 3 LLM에 더 풍부한 컨텍스트)
    result["biz_detail"] = section[:1000]

    return result

# ══════════════════════════════════════════════════════════
# 메인 배치
# ══════════════════════════════════════════════════════════
def main():
    print("="*60)
    print("DART 사업보고서 배치 수집")
    print("="*60)

    # 기존 캐시 로드
    cache: dict = {}
    if CACHE_PATH.exists() and args.resume:
        cache = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
        print(f"[캐시] 기존 {len(cache)}개 로드 (--resume)")

    # CSV 로드
    companies = load_companies()
    if args.limit > 0:
        companies = companies[:args.limit]
        print(f"[제한] {args.limit}개만 처리 (--limit)")

    # 종목 코드 컬럼명 탐지 (CSV 구조에 따라 다를 수 있음)
    if companies:
        cols = list(companies[0].keys())
        print(f"[CSV] 컬럼: {cols}")
    stock_col = next((c for c in ["종목코드","stock_code","Stock Code","종목"] if c in cols), cols[0])
    name_col  = next((c for c in ["회사명","corp_name","Company","기업명"] if c in cols), cols[1])
    ksic_col  = next((c for c in ["업종코드","ksic_code","KSIC","업종"] if c in cols), "")
    print(f"[CSV] 종목코드={stock_col}, 회사명={name_col}, 업종={ksic_col or 'N/A'}")

    # stock_code → corp_code 맵
    stock_codes = {str(c[stock_col]).zfill(6) for c in companies}
    corpcode_map = build_corpcode_map(stock_codes)
    time.sleep(1)

    # 처리 대상 필터 — 중복 제거 + resume 지원
    seen_sc = set()
    todo = []
    dup_count = 0
    for co in companies:
        sc = str(co[stock_col]).zfill(6)
        if sc in seen_sc:
            dup_count += 1
            continue           # CSV 중복 행 제거
        seen_sc.add(sc)
        if args.resume and sc in cache:
            continue           # 이미 수집된 기업 건너뜀
        todo.append(co)
    if dup_count:
        print(f"[중복] CSV 중복 종목코드 {dup_count}개 제거")
    print(f"\n[배치] 처리 대상: {len(todo)}개 / 전체 {len(companies)}개")
    print(f"[배치] 예상 시간: {len(todo)*4*DELAY/60:.0f}분\n")

    # ── 배치 루프 ──────────────────────────────────────────
    ok = skip = fail = 0
    start_time = time.time()

    for idx, co in enumerate(todo, 1):
        sc        = str(co[stock_col]).zfill(6)
        corp_name = co.get(name_col, sc)
        ksic      = co.get(ksic_col, "") if ksic_col else ""
        corp_code = corpcode_map.get(sc, "")

        print(f"[{idx:4d}/{len(todo)}]  {corp_name} ({sc})", end="  ")

        if not corp_code:
            print("corp_code 없음 → SKIP")
            cache[sc] = {"corp_name":corp_name,"skip_reason":"no_corp_code"}
            skip += 1
            continue

        entry = {
            "corp_name":   corp_name,
            "corp_code":   corp_code,
            "stock_code":  sc,
            "ksic_code":   ksic,
            "revenue_억원": 0.0,
            "biz_summary": "",
            "biz_detail":  "",
            "parse_status":"",
        }

        # list.json → rcept_no
        try:
            rcept_no, rcept_dt = fetch_rcept_no(corp_code)
            time.sleep(DELAY)
        except Exception as e:
            print(f"list 오류:{e} → SKIP")
            entry["parse_status"] = f"list_error:{type(e).__name__}"
            cache[sc] = entry; fail += 1; continue

        if not rcept_no:
            print("사업보고서 없음 → SKIP")
            entry["parse_status"] = "no_rcept_no"
            cache[sc] = entry; skip += 1; continue

        # document.xml → biz_summary
        try:
            doc = fetch_biz_summary(rcept_no)
            time.sleep(DELAY)
        except Exception as e:
            print(f"doc 오류:{e} → SKIP")
            entry["parse_status"] = f"doc_error:{type(e).__name__}"
            cache[sc] = entry; fail += 1; continue

        entry.update({
            "rcept_no":    rcept_no,
            "rcept_dt":    rcept_dt or "",
            "parse_status": doc["parse_status"],
            "biz_summary": doc["biz_summary"],
            "biz_detail":  doc["biz_detail"],
        })

        status = doc["parse_status"]
        q      = _quality(doc["biz_summary"])
        print(f"{status}  품질={q:.3f}  biz={doc['biz_summary'][:60].replace(chr(10),' ')}...")
        if "found" in status or "fallback" in status:
            ok += 1
        else:
            fail += 1

        cache[sc] = entry

        # 체크포인트 저장
        if idx % CKPT_EVERY == 0:
            CACHE_DIR.mkdir(parents=True, exist_ok=True)
            CACHE_PATH.write_text(
                json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
            elapsed = (time.time()-start_time)/60
            remain  = elapsed/idx*(len(todo)-idx) if idx>0 else 0
            print(f"\n  ── 체크포인트 저장 ({idx}/{len(todo)}) "
                  f"경과={elapsed:.0f}분 / 예상남은={remain:.0f}분 ──\n")

    # ── 최종 저장 ──────────────────────────────────────────
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    CACHE_PATH.write_text(
        json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")

    elapsed = (time.time()-start_time)/60
    print(f"\n{'='*60}")
    print("배치 완료 리포트")
    print(f"{'='*60}")
    print(f"  총 처리    : {len(todo)}개  ({elapsed:.0f}분)")
    print(f"  성공(biz)  : {ok}개  ({ok/max(len(todo),1)*100:.1f}%)")
    print(f"  SKIP       : {skip}개  (보고서 없음 / corp_code 없음)")
    print(f"  실패       : {fail}개  (API 오류)")
    print(f"  캐시 저장  : {CACHE_PATH}")

    # 품질 분포
    print(f"\n  [biz_summary 품질 분포]")
    qs = [_quality(v.get("biz_summary","")) for v in cache.values()
          if v.get("biz_summary")]
    if qs:
        import statistics
        print(f"    평균={statistics.mean(qs):.3f}  "
              f"중앙값={statistics.median(qs):.3f}  "
              f"min={min(qs):.3f}  max={max(qs):.3f}")
        good = sum(1 for q in qs if q > 0.1)
        print(f"    품질>0.1(양호): {good}개 ({good/len(qs)*100:.1f}%)")

if __name__ == "__main__":
    main()