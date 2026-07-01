"""
dart_biz_sampler.py  v5
────────────────────────────────────────────────────────────
개선 사항:
  - DART_API_BUSINESS_KEY 단일 키로 통일
  - 이니텍형 XBRL 구조 개선:
      · 목차(TOC) 매칭 → 본문 최고품질 섹션 선택
      · 동일 키워드 복수 occurrence 중 한글 비율 최대 섹션 선택
      · XBRL % 없는 소수점 비율 패턴 추가 (52.3 형태)
      · _00761.xml 도 후보에 포함
────────────────────────────────────────────────────────────
"""
import os, sys, json, re, time, zipfile, io, datetime
import requests
from pathlib import Path

# ── .env 로드 ─────────────────────────────────────────────
def _parse_val(raw: str) -> str:
    v = raw.strip()
    if v.startswith('"'):
        e = v.find('"', 1); return v[1:e].strip() if e!=-1 else v[1:].strip()
    if v.startswith("'"):
        e = v.find("'", 1); return v[1:e].strip() if e!=-1 else v[1:].strip()
    for sep in (' #', '\t#'):
        if sep in v: v = v[:v.index(sep)]
    return v.strip().strip('"').strip("'").strip()

def load_env() -> dict:
    targets = {"DART_API_KEY","DART_API_BUSINESS_KEY","DART_KEY","GEMINI_API_KEY"}
    found = {}
    sd = Path(__file__).parent.resolve()
    for p in [sd/".."/ ".env", sd/".env"]:
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
        if found:
            print(f"[.env] {sorted(found.keys())} → {p}")
            break
    return found

ENV  = load_env()
# DART_API_BUSINESS_KEY 단일 키 사용
KEY  = (ENV.get("DART_API_BUSINESS_KEY")
        or ENV.get("DART_API_KEY")
        or ENV.get("DART_KEY", ""))

print(f"[키] DART_API_BUSINESS_KEY  길이={len(KEY)}  {'✅' if len(KEY)==40 else '❌'}")
if not KEY: sys.exit("[ERROR] DART API 키를 찾을 수 없습니다.")

BASE  = "https://opendart.fss.or.kr/api"
DELAY = 0.5

SAMPLES: dict = {
    "053350": ("이니텍",       "00353610"),
    "041460": ("한국전자인증",  "00361169"),
}

# ══ A. company.json ════════════════════════════════════════
def fetch_company(corp_code: str) -> dict:
    r = requests.get(f"{BASE}/company.json",
                     params={"crtfc_key": KEY, "corp_code": corp_code}, timeout=15)
    d = r.json()
    print(f"  [A] company  status={d.get('status')}  "
          f"induty={d.get('induty_code','')}  acc_mt={d.get('acc_mt','')}")
    return d

# ══ B. 매출액 (fnlttSinglAcntAll) ══════════════════════════
def fetch_revenue(corp_code: str) -> float:
    for yr in ("2024", "2023"):
        for fs in ("CFS", "OFS"):
            r = requests.get(f"{BASE}/fnlttSinglAcntAll.json",
                             params={"crtfc_key": KEY, "corp_code": corp_code,
                                     "bsns_year": yr, "reprt_code": "11011",
                                     "fs_div": fs}, timeout=15)
            d = r.json(); time.sleep(DELAY)
            if d.get("status") == "000":
                for it in d.get("list", []):
                    if it.get("account_nm", "") in ("매출액", "수익(매출액)"):
                        try:
                            v = int(it.get("thstrm_amount","0").replace(",",""))
                            print(f"  [B] 재무  {yr}/{fs}  매출액={v/1e8:.1f}억원")
                            return round(v / 1e8, 1)
                        except: pass
    print("  [B] 재무  매출액 미추출"); return 0.0

# ══ C. list.json → rcept_no ════════════════════════════════
def fetch_rcept_no(corp_code: str) -> tuple:
    today  = datetime.date.today().strftime("%Y%m%d")
    bgn_de = (datetime.date.today()-datetime.timedelta(days=365*3)).strftime("%Y%m%d")
    r = requests.get(f"{BASE}/list.json",
                     params={"crtfc_key": KEY, "corp_code": corp_code,
                             "pblntf_ty": "A", "bgn_de": bgn_de,
                             "end_de": today, "page_count": 10}, timeout=15)
    d = r.json()
    print(f"  [C] list  status={d.get('status')}  총건수={d.get('total_count','N/A')}")
    if d.get("status") != "000": return None, None, None

    items = d.get("list", [])
    for it in items:
        print(f"      {it.get('rcept_dt','')}  "
              f"{it.get('report_nm','')[:45]}  rcept_no={it.get('rcept_no','')}")

    # 1순위: 정정 없는 원본 사업보고서
    for it in items:
        nm, rno = it.get("report_nm",""), it.get("rcept_no","")
        if "사업보고서" in nm and rno and not nm.startswith("["):
            return rno, it.get("rcept_dt",""), nm
    # 2순위: 정정본
    for it in items:
        nm, rno = it.get("report_nm",""), it.get("rcept_no","")
        if "사업보고서" in nm and rno:
            print(f"      ⚠ 정정본 사용: {nm[:40]}")
            return rno, it.get("rcept_dt",""), nm
    return None, None, None

# ══ 텍스트 품질 스코어 ══════════════════════════════════════
def _quality(text: str) -> float:
    """한글 비율 × min(len/500, 1.0) — 짧거나 한글 없으면 낮은 점수."""
    if not text: return 0.0
    kor = len(re.findall(r'[가-힣]', text))
    return (kor / max(len(text), 1)) * min(len(text) / 500, 1.0)

# ══ 사업 섹션 최고품질 occurrence 선택 ═════════════════════
def _best_section(clean: str) -> tuple[str, str]:
    """(section_text, status) 반환. TOC 매칭 건너뛰고 본문 선택."""
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
            candidates.append((text, q, pat))

    if candidates:
        best = max(candidates, key=lambda x: x[1])
        if best[1] > 0.05:   # 한글 5% 이상이어야 유효
            return best[0], "section_found"

    # 사업 개요로 대체
    for kw in ["사업의 개요", "사업 개요", "회사의 개요"]:
        idx = clean.find(kw)
        if idx != -1:
            snippet = clean[idx:idx+1500].strip()
            if _quality(snippet) > 0.05:
                return snippet, "overview_fallback"

    return clean[:2000], "head_fallback"

# ══ 매출 비중 추출 (% 기호 있는 경우 + XBRL 소수점 패턴) ══
def _extract_products(section: str) -> list:
    products = []; seen = set()
    # 패턴 1: "제품명 ... 52.3%" (일반 HTML)
    for m in re.finditer(
            r"([가-힣A-Za-z0-9\s/·]{2,25}?)\s+(\d{1,3}(?:\.\d+)?)\s*%", section):
        nm, sh = m.group(1).strip(), float(m.group(2))
        if nm not in seen and 0 < sh <= 100:
            products.append({"name": nm, "revenue_share_pct": sh, "src": "pct"})
            seen.add(nm)
    # 패턴 2: XBRL — "제품명  0.523" 소수점 비율 형태
    if not products:
        for m in re.finditer(
                r"([가-힣]{2,15}(?:솔루션|서비스|사업|제품|SW|시스템|플랫폼)?)"
                r"\s+(0\.\d{2,4}|\d{1,2}\.\d{1,4})\s", section):
            nm, val = m.group(1).strip(), float(m.group(2))
            sh = val * 100 if val < 1 else val   # 0.523 → 52.3%, 52.3 → 52.3%
            if nm not in seen and 0 < sh <= 100:
                products.append({"name": nm, "revenue_share_pct": round(sh,1), "src": "xbrl_ratio"})
                seen.add(nm)
            if len(products) >= 8: break
    return products[:8]

# ══ D. document.xml → ZIP → 파싱 ══════════════════════════
def fetch_and_parse(rcept_no: str, corp_name: str) -> dict:
    out = {"parse_status":"pending","biz_summary":"","products":[],"raw_section":""}

    r = requests.get(f"{BASE}/document.xml",
                     params={"crtfc_key": KEY, "rcept_no": rcept_no}, timeout=90)
    print(f"  [D] doc  HTTP={r.status_code}  "
          f"크기={len(r.content):,}bytes  ZIP={r.content[:2]==b'PK'}")
    if r.content[:2] != b'PK':
        out["parse_status"] = f"not_zip:{r.content[:150]}"; return out

    try:
        zf = zipfile.ZipFile(io.BytesIO(r.content))
    except Exception as e:
        out["parse_status"] = f"zip_error:{e}"; return out

    names = zf.namelist()
    out["zip_files"] = names
    print(f"      ZIP({len(names)}개): {names}")

    # ── 파일 후보 선택 ─────────────────────────────────────
    # htm 우선, 없으면 xml — 감사보고서는 스캔해서 제외
    AUDIT_KEYS = ["감사보고서", "감사의견", "audit", "Audit"]
    BIZ_KEYS   = ["주요 제품", "사업의 내용", "주요제품", "사업내용",
                  "제품 및 서비스", "영업 개황", "매출구성"]

    biz_file = None

    # 1차: htm/html 탐색
    htms = [n for n in names if n.lower().endswith((".htm",".html"))]
    for n in htms:
        try:
            raw = zf.read(n).decode("utf-8", errors="ignore")
            txt = re.sub(r"<[^>]+>", " ", raw)
            if not any(k in txt[:600] for k in AUDIT_KEYS):
                if any(k in txt for k in BIZ_KEYS):
                    biz_file = n; print(f"      ✅ htm 사업키워드: {n}"); break
        except: pass

    # 2차: xml 탐색 (XBRL, 2024년 이후)
    if not biz_file:
        xmls = [n for n in names if n.lower().endswith(".xml")]
        scored = []
        for n in xmls:
            try:
                raw = zf.read(n).decode("utf-8", errors="ignore")
                txt = re.sub(r"<[^>]+>", " ", raw)
                txt = re.sub(r"\s+", " ", txt)
                if any(k in txt[:800] for k in AUDIT_KEYS):
                    print(f"      건너뜀(감사): {n}"); continue
                biz_hit = sum(1 for k in BIZ_KEYS if k in txt)
                q = _quality(txt[:5000])
                scored.append((n, biz_hit, q, txt))
                print(f"      후보 {n}: biz_hit={biz_hit}  quality={q:.3f}")
            except Exception as e:
                print(f"      읽기실패 {n}: {e}")

        if scored:
            # 사업 키워드 히트 수 → 품질 순으로 정렬
            scored.sort(key=lambda x: (x[1], x[2]), reverse=True)
            biz_file = scored[0][0]
            print(f"      → 선택: {biz_file}  "
                  f"(biz_hit={scored[0][1]}, quality={scored[0][2]:.3f})")

    if not biz_file:
        out["parse_status"] = "no_file"; return out

    out["target_file"] = biz_file

    # ── 텍스트 추출 ────────────────────────────────────────
    raw = zf.read(biz_file).decode("utf-8", errors="ignore")
    clean = re.sub(r"<[^>]+>", " ", raw)
    clean = re.sub(r"[ \t]+", " ", clean)          # 가로 공백만 정리
    clean = re.sub(r"\n{3,}", "\n\n", clean).strip()

    # ── 사업 섹션 최고품질 선택 ────────────────────────────
    section, status = _best_section(clean)
    out["parse_status"] = status
    out["raw_section"]  = section

    # ── 매출 비중 추출 ─────────────────────────────────────
    out["products"] = _extract_products(section)

    # ── 사업 개요 (앞 600자) ──────────────────────────────
    ov = re.search(r"사업(?:의)?\s*개요(.*?)(?=주요\s*제품|원재료|\Z)",
                   clean, re.DOTALL)
    summary = ov.group(1)[:600].strip() if ov else section[:600]
    # 의미없는 짧은 결과 체크
    if _quality(summary) < 0.05:
        # 사업개요 못 찾으면 섹션 앞 600자로 대체
        summary = section[:600]
    out["biz_summary"] = summary

    return out

# ══ 드림시큐리티 corp_code 검색 ════════════════════════════
def find_listed_corp(keyword: str) -> dict | None:
    r = requests.get(f"{BASE}/corpCode.xml",
                     params={"crtfc_key": KEY}, timeout=30)
    with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
        xml = zf.read(zf.namelist()[0]).decode("utf-8", errors="ignore")
    hits = []
    for block in re.findall(r"<list>(.*?)</list>", xml, re.DOTALL):
        nm = re.search(r"<corp_name>(.*?)</corp_name>", block)
        cc = re.search(r"<corp_code>(.*?)</corp_code>", block)
        sc = re.search(r"<stock_code>(.*?)</stock_code>", block)
        if nm and keyword in nm.group(1):
            hits.append({"corp_name": nm.group(1).strip(),
                         "corp_code": cc.group(1).strip() if cc else "",
                         "stock_code": sc.group(1).strip() if sc else ""})
    listed = [h for h in hits if h.get("stock_code","").strip()]
    return listed[0] if listed else (hits[0] if hits else None)

# ══ 메인 ════════════════════════════════════════════════════
def main():
    print("\n" + "="*60)
    print("DART 사업보고서 검증  v5  (Business 키 단일 사용)")
    print("="*60)

    print("\n[사전] 드림시큐리티 corp_code 검색...")
    h = find_listed_corp("드림시큐리티")
    if h:
        print(f"  → {h}")
        if h["stock_code"]:
            SAMPLES[h["stock_code"]] = ("드림시큐리티", h["corp_code"])

    all_results = {}

    for stock, (corp_name, corp_code) in SAMPLES.items():
        print(f"\n{'─'*55}")
        print(f"▶  {corp_name} ({stock})  corp_code={corp_code}")
        res = {"corp_name": corp_name, "corp_code": corp_code}

        co = fetch_company(corp_code); time.sleep(DELAY)
        res["induty_code"] = co.get("induty_code","")
        res["acc_mt"]      = co.get("acc_mt","")

        res["revenue_억원"] = fetch_revenue(corp_code); time.sleep(DELAY)

        rcept_no, rcept_dt, report_nm = fetch_rcept_no(corp_code); time.sleep(DELAY)
        res["rcept_no"] = rcept_no
        if not rcept_no:
            print("  → rcept_no 없음 SKIP")
            res["doc"] = {"parse_status":"no_rcept_no"}
            all_results[stock] = res; continue

        print(f"  → rcept_no={rcept_no}  ({report_nm})")

        doc = fetch_and_parse(rcept_no, corp_name); time.sleep(DELAY)
        res["doc"] = doc

        print(f"\n  ── 파싱 결과 ──")
        print(f"  status={doc['parse_status']}  파일={doc.get('target_file','N/A')}")
        print(f"  [매출구성] {len(doc['products'])}건:")
        for p in doc["products"]:
            print(f"    {p['name']:22s}  {p['revenue_share_pct']:5.1f}%  ({p.get('src','')})")
        if not doc["products"]: print("    (미검출)")
        print(f"  [사업개요 앞200자]\n    {doc['biz_summary'][:200]}")
        print(f"  [raw_section 앞300자]\n    {doc['raw_section'][:300]}")

        all_results[stock] = res

    # ── 저장 ───────────────────────────────────────────────
    out = Path(__file__).parent / "dart_biz_cache_v5.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)

    print(f"\n{'='*60}\n종합 평가")
    for sc, r in all_results.items():
        doc = r.get("doc", {}); st = doc.get("parse_status","?")
        ok  = "✅" if "found" in st else ("⚠️" if "fallback" in st else "❌")
        print(f"  {r['corp_name']:12s}  매출={r.get('revenue_억원',0):>8.1f}억  "
              f"파싱={ok}({st})  매출구성={len(doc.get('products',[]))}개")
    print(f"\n  저장: {out}")

if __name__ == "__main__":
    main()