"""
Phase 3 v2: 증권신고서 XML 다운로드 및 피처 추출

DART document.xml API → ZIP → XML 파싱 (태그 제거 후 정규식)

추출 피처:
  [공모가 구조]
  - price_band_low        : 희망공모가 하단 (원)
  - price_band_high       : 희망공모가 상단 (원)
  - final_offering_price  : 확정 공모가 (원)
  - band_midpoint         : (low+high)/2
  - price_vs_mid          : final/mid - 1  (밴드 내 위치, 양수=상단초과)

  [밸류에이션]
  - net_income_bn         : 당기순이익 (백만원 → 억원)
  - peer_per              : 유사기업 PER (배)
  - per_stock_value       : 주당 평가가액 (원)
  - total_shares          : 평가 기준 주식수 (주)

  [공모 구조]
  - offering_shares       : 공모 주식수 (주)
  - offering_amount       : 공모총액 (원)
  - institutional_ratio   : 기관투자자 배정비율 (%)
  - retail_ratio          : 일반투자자 배정비율 (%)

  [상장 정보]
  - listing_market        : 코스닥 / 코스피
  - is_tech_special       : 기술특례 여부 (0/1)
  - is_growth_special     : 성장성특례 여부 (0/1)

입력:  ipo_with_rcpno.csv
출력:  ipo_features_raw.csv
       ipo_features_raw_log.csv

사용법:
  python phase3_extract_features.py --input ipo_with_rcpno.csv --output ipo_features_raw.csv
  python phase3_extract_features.py --input ipo_with_rcpno.csv --output ipo_features_raw.csv --max-rows 5
  python phase3_extract_features.py --input ipo_with_rcpno.csv --output ipo_features_raw.csv --skip-download --xml-dir xml
"""

import os, io, re, sys, time, zipfile, argparse, logging
import requests
import pandas as pd
from pathlib import Path
from typing import Optional, Dict, Any
from dotenv import load_dotenv

load_dotenv()
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

DART_DOC_URL = "https://opendart.fss.or.kr/api/document.xml"
RETRY        = 3
RETRY_WAIT   = 2.0


def normalize_rcept_no(value) -> str:
    """DART 접수번호를 14자리 문자열로 정규화.

    pandas가 빈 값이 섞인 rcept_no 컬럼을 float로 읽으면
    20220111000212 -> 20220111000212.0 형태가 되어 document API가 실패한다.
    """
    if value is None:
        return ""
    text = str(value).strip()
    if text == "" or text.lower() == "nan":
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    text = re.sub(r"\D", "", text)
    return text if len(text) == 14 else ""


# ══════════════════════════════════════════════════════════
# 1. 다운로드
# ══════════════════════════════════════════════════════════

def download_xml(api_key: str, rcept_no: str, save_dir: Path,
                 rate_limit: float) -> Optional[Path]:
    """ZIP 다운로드 → XML 저장. 이미 있으면 재사용."""
    save_dir.mkdir(parents=True, exist_ok=True)
    dest = save_dir / f"{rcept_no}.xml"
    if dest.exists() and dest.stat().st_size > 500:
        return dest

    url = f"{DART_DOC_URL}?crtfc_key={api_key}&rcept_no={rcept_no}"
    for attempt in range(RETRY):
        try:
            r = requests.get(url, timeout=60)
            time.sleep(rate_limit)
            if r.status_code != 200:
                log.warning(f"{rcept_no}: HTTP {r.status_code}")
                time.sleep(RETRY_WAIT); continue
            if r.content[:4] != b'PK\x03\x04':
                log.warning(f"{rcept_no}: ZIP 아님")
                return None
            with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                xml_names = [n for n in z.namelist() if n.endswith('.xml')]
                if not xml_names:
                    log.warning(f"{rcept_no}: XML 없음 → {z.namelist()}")
                    return None
                dest.write_bytes(z.read(xml_names[0]))
            log.info(f"{rcept_no}: 저장 {dest.stat().st_size:,}B")
            return dest
        except requests.RequestException as e:
            log.warning(f"{rcept_no}: 요청실패({attempt+1}) {e}")
            time.sleep(RETRY_WAIT)
    return None


# ══════════════════════════════════════════════════════════
# 2. XML → plain text
# ══════════════════════════════════════════════════════════

def xml_to_plain(path: Path) -> str:
    """XML 태그 제거 후 공백 정규화."""
    for enc in ('utf-8', 'euc-kr', 'cp949'):
        try:
            raw = path.read_text(encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        return ""
    plain = re.sub(r'<[^>]+>', ' ', raw)
    plain = re.sub(r'\s+', ' ', plain)
    return plain


# ══════════════════════════════════════════════════════════
# 3. 숫자 파싱 유틸
# ══════════════════════════════════════════════════════════

def _num(s: str) -> Optional[float]:
    """'1,234.5' → 1234.5 / None on failure."""
    if not s: return None
    s = re.sub(r'[,\s원주배%]', '', str(s).strip())
    s = re.sub(r'\(?\s*주\s*\d+\s*\)?', '', s)
    try:    return float(s)
    except: return None


def _fmt_num(v: Optional[float]) -> str:
    """천단위 콤마 포맷. None이면 'N/A' (status=='ok'여도 일부 필드는 None일 수 있음)."""
    return f"{v:,}" if v is not None else "N/A"


# ══════════════════════════════════════════════════════════
# 4. 피처 추출
# ══════════════════════════════════════════════════════════

def extract(plain: str) -> Dict[str, Any]:
    f: Dict[str, Any] = {}

    # ── 4-1. 희망공모가 밴드 ─────────────────────────────
    # 패턴: "모집(매출)가액(예정): 8,000원 ~ 9,100원"
    #   또는 "공모희망가액인 8,000원 ~ 9,100원"
    m = re.search(
        r'(?:모집\(매출\)가액\(예정\)|공모희망가액)[^\d]*([\d,]+)원\s*~\s*([\d,]+)원',
        plain
    )
    if m:
        low, high = _num(m.group(1)), _num(m.group(2))
        if low and high and low >= 100 and high >= 100:
            f['price_band_low']  = low
            f['price_band_high'] = high
            f['band_midpoint']   = round((low + high) / 2, 2)

    # ── 4-2. 확정 공모가 ─────────────────────────────────
    # 패턴A: "확정 공모가액인 9,100원"
    # 패턴B: "확정공모가액을 9,100원으로 최종 결정"
    # 패턴C: 정정후 표에서 "모집(매출)가액(예정): 9,100원" (단일값, 범위 아님)
    final = None
    for pat in [
        r'확정\s*공모가액인\s*([\d,]+)원',
        r'확정공모가액을\s*([\d,]+)원으로',
        r'확정\s*공모가액\s*:\s*([\d,]+)원',
        r'1주당\s*확정공모가액을\s*([\d,]+)원',
    ]:
        m = re.search(pat, plain)
        if m:
            v = _num(m.group(1))
            if v and v >= 100:
                final = v; break

    # 패턴D: 정정후 칸 "모집(매출)가액(예정): 9,100원" (단일 숫자, ~없음)
    if final is None:
        # "정정 후" 이후 텍스트에서 단일 공모가 탐색
        after_idx = plain.find('정정 후')
        if after_idx == -1: after_idx = plain.find('정정후')
        search_area = plain[after_idx:after_idx+2000] if after_idx >= 0 else plain[:3000]
        m = re.search(
            r'모집\(매출\)가액\(예정\)\s*:\s*([\d,]+)원(?!\s*~)',
            search_area
        )
        if m:
            v = _num(m.group(1))
            if v and v >= 100: final = v

    f['final_offering_price'] = final
    if final and f.get('band_midpoint'):
        f['price_vs_mid'] = round(final / f['band_midpoint'] - 1, 4)

    # ── 4-3. 밸류에이션 산정표 ───────────────────────────
    # 당기순이익 — 단위: 백만원 또는 천원
    # 패턴A: "당기순이익 8,120 백만원"
    # 패턴B: "지배주주 당기순이익 10,658,266천원"
    # 패턴C: "연결당기순이익 11,760,417천원"
    m = re.search(
        r'(?:지배기업\s*소유주\s*지분\s*연결)?(?:지배주주\s*)?당기순이익[^\d]*([\d,]+)\s*(백만원|천원)',
        plain
    )
    if m:
        v, unit = _num(m.group(1)), m.group(2)
        if v:
            if unit == '백만원':
                f['net_income_mn'] = v
                f['net_income_bn'] = round(v / 100, 4)
            else:  # 천원
                f['net_income_mn'] = round(v / 1000, 4)
                f['net_income_bn'] = round(v / 100_000, 4)

    # PER — "유사기업 PER" 또는 "비교대상회사 PER"
    m = re.search(r'(?:유사기업|비교대상회사)\s*PER\s*([\d,.]+)배', plain)
    if m:
        v = _num(m.group(1))
        if v: f['peer_per'] = v

    # EV/EBITDA 모형 fallback (한국피아이엠 등)
    if f.get('peer_per') is None:
        if 'EV/EBITDA' in plain:
            f['valuation_model'] = 'EV/EBITDA'
            # 배수: "비교기업 평 균 EV/EBITDA 배 11.81"  (공백 포함, 단위 분리 형식)
            m2 = re.search(
                r'(?:유사기업|비교기업|비교대상회사)[^\d\n]*EV/EBITDA\s*배?\s*([\d,.]+)',
                plain
            )
            if m2:
                v = _num(m2.group(1))
                if v: f['peer_per'] = v

            # EBITDA 금액: "EBITDA 천원 9,062,292"
            m3 = re.search(r'EBITDA\s*(천원|백만원)\s*([\d,]+)', plain)
            if m3:
                v, unit = _num(m3.group(2)), m3.group(1)
                if v:
                    if unit == '천원':
                        f['net_income_mn'] = round(v / 1000, 4)
                        f['net_income_bn'] = round(v / 100_000, 4)
                    else:
                        f['net_income_mn'] = v
                        f['net_income_bn'] = round(v / 100, 4)
                f['net_income_type'] = 'EBITDA'

            # 주식수: "⑤ 주식수 주 6,080,457"
            if f.get('total_shares') is None:
                m4 = re.search(r'[④⑤]\s*(?:공모유입자금[^\n]*\n[^\n]*)?주식수\s*주\s*([\d,]+)', plain)
                if not m4:
                    m4 = re.search(r'주식수\s*주\s*([\d,]+)', plain)
                if m4:
                    v = _num(m4.group(1))
                    if v: f['total_shares'] = v
        else:
            f['valuation_model'] = 'PER'
    else:
        f['valuation_model'] = 'PER'

    # 주당 평가가액: "주당 평가 가액 13,342원"
    m = re.search(r'주당\s*평가\s*가\s*액\s*([\d,]+)원', plain)
    if m:
        v = _num(m.group(1))
        if v and v >= 100: f['per_stock_value'] = v

    # 주식수(평가기준): "③ 주식수 15,524,430 주"
    m = re.search(r'③\s*주식수\s*([\d,]+)\s*주', plain)
    if not m:
        m = re.search(r'주식수\s*([\d,]+)\s*주\s*①', plain)
    if m:
        v = _num(m.group(1))
        if v: f['total_shares'] = v

    # ── 4-4. 공모 구조 ────────────────────────────────────
    # 공모주식수: "기명식 보통주 3,000,000주"
    m = re.search(r'기명식\s*보통주\s*([\d,]+)주', plain)
    if m:
        v = _num(m.group(1))
        if v: f['offering_shares'] = v

    # 공모총액: "모집 또는 매출금액 : 27,300,000,000원"
    m = re.search(r'모집\s*또는\s*매출금액\s*:\s*([\d,]+)원', plain)
    if m:
        v = _num(m.group(1))
        if v: f['offering_amount'] = v

    # 기관투자자 배정비율 (확정 후 단일값)
    # 패턴: "기관투자자 배정비율: 72.0%"  (정정후 = 단일값)
    # 정정후 섹션에서만 탐색
    after_idx = plain.find('정정 후')
    if after_idx == -1: after_idx = plain.find('정정후')
    search_area = plain[after_idx:after_idx+2000] if after_idx >= 0 else ''

    m = re.search(r'기관투자자\s*배정비율\s*:\s*([\d.]+)%(?!\s*~)', search_area)
    if m:
        v = _num(m.group(1))
        if v: f['institutional_ratio'] = v

    m = re.search(r'일반투자자\s*배정비율\s*:\s*([\d.]+)%(?!\s*~)', search_area)
    if m:
        v = _num(m.group(1))
        if v: f['retail_ratio'] = v

    # ── 4-5. 할인율 추출 (교차검증용) ───────────────────────
    # 신버전: "주당 평가가액에 대한 할인율 39.67% ~ 31.42%"  (high→low)
    # 구버전: "평가액 대비 할인율(주1) 35.52% ~ 21.35%"
    # 불일치 방지: 정정후 섹션 우선, fallback으로 전체 탐색
    # 정정후 섹션 추출 (2000자 → 5000자로 확대)
    after_idx2 = plain.find('정정 후')
    if after_idx2 == -1: after_idx2 = plain.find('정정후')
    # 가장 마지막 "정정 후" 위치 사용 (확정본이 마지막에 위치)
    last_after = plain.rfind('정정 후')
    if last_after == -1: last_after = plain.rfind('정정후')
    search_zones = []
    if last_after >= 0:
        search_zones.append(plain[last_after:last_after+5000])   # 정정후 우선
    search_zones.append(plain)                                    # 전체 fallback

    discount_patterns = [
        r'주당\s*평가\s*가액에\s*대한\s*할인율\s*([\d.]+)%\s*~\s*([\d.]+)%',
        r'평가액\s*대비\s*할인율(?:\(주\d+\))?\s*([\d.]+)%\s*~\s*([\d.]+)%',
        r'할인율\s*\([④⑤]\)\s*([\d.]+)%\s*~\s*([\d.]+)%',
        r'할인율\s*[：:]\s*([\d.]+)%\s*~\s*([\d.]+)%',
    ]
    for zone in search_zones:
        matched = False
        for pat in discount_patterns:
            m = re.search(pat, zone)
            if m:
                v1, v2 = float(m.group(1)), float(m.group(2))
                f['discount_rate_xml_low']  = round(min(v1, v2), 4)
                f['discount_rate_xml_high'] = round(max(v1, v2), 4)
                f['discount_rate_xml_mid']  = round((v1 + v2) / 2, 4)
                matched = True
                break
        if matched:
            break

    # ── 4-6. 상장 정보 ────────────────────────────────────
    f['listing_market']    = '코스닥' if '코스닥' in plain[:8000] else \
                             ('코스피' if '코스피' in plain[:8000] else None)
    f['is_tech_special']   = int(bool(re.search(r'기술특례|기술성장기업|기술평가', plain[:8000])))
    f['is_growth_special'] = int(bool(re.search(r'성장성특례|성장성\s*추천', plain[:8000])))

    # ── 4-7. 추출 품질 평가 ──────────────────────────────
    # Fallbacks for older/newly added IPO filings.  Many filings express the
    # valuation table as "적용 순이익", "적용 PER", "주당 평가가액", while the
    # first-pass patterns above only catch a narrower set of phrases.
    def _plain_num(value: str) -> Optional[float]:
        value = re.sub(r'[^\d.\-]', '', str(value or ''))
        try:
            return float(value) if value else None
        except ValueError:
            return None

    if f.get('peer_per') is None:
        for pat in [
            r'적용\s*PER\s*([\d,.]+)\s*배',
            r'적용PER\s*([\d,.]+)\s*배',
            r'평균\s*PER\s*([\d,.]+)\s*배',
        ]:
            m = re.search(pat, plain)
            if m:
                v = _plain_num(m.group(1))
                if v and 0 < v < 200:
                    f['peer_per'] = v
                    f['valuation_model'] = 'PER'
                    break

    if f.get('net_income_bn') is None:
        for pat in [
            r'적용\s*(?:당기)?순이익[^\d]{0,30}([\d,.]+)\s*(원|천원|백만원)',
            r'(?:추정\s*)?당기순이익[^\d]{0,30}([\d,.]+)\s*(원|천원|백만원)',
            r'순이익[^\d]{0,30}([\d,.]+)\s*(원|천원|백만원)',
        ]:
            m = re.search(pat, plain)
            if not m:
                continue
            v = _plain_num(m.group(1))
            unit = m.group(2)
            if not v:
                continue
            if unit == '원':
                f['net_income_mn'] = round(v / 1_000_000, 4)
                f['net_income_bn'] = round(v / 100_000_000, 4)
            elif unit == '천원':
                f['net_income_mn'] = round(v / 1000, 4)
                f['net_income_bn'] = round(v / 100_000, 4)
            else:  # 백만원
                f['net_income_mn'] = v
                f['net_income_bn'] = round(v / 100, 4)
            break

    if f.get('per_stock_value') is None:
        for pat in [
            r'주당\s*평가\s*가액\s*([\d,.]+)\s*원',
            r'주당평가가액\s*([\d,.]+)\s*원',
        ]:
            m = re.search(pat, plain)
            if m:
                v = _plain_num(m.group(1))
                if v and v >= 100:
                    f['per_stock_value'] = v
                    break

    if f.get('total_shares') is None:
        m = re.search(r'적용\s*주식수\s*([\d,]+)\s*주', plain)
        if m:
            v = _plain_num(m.group(1))
            if v:
                f['total_shares'] = v

    key_fields = ['price_band_low', 'price_band_high',
                  'final_offering_price', 'net_income_bn', 'peer_per']
    filled = sum(1 for k in key_fields if f.get(k) is not None)
    f['filled_key_fields'] = filled
    f['extraction_status'] = 'ok' if filled >= 4 else \
                             ('partial' if filled >= 2 else 'fail')
    return f


# ══════════════════════════════════════════════════════════
# 5. 메인
# ══════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input',          required=True)
    parser.add_argument('--output',         required=True)
    parser.add_argument('--xml-dir',        default='xml')
    parser.add_argument('--skip-download',  action='store_true')
    parser.add_argument('--rate-limit-sec', type=float, default=0.5)
    parser.add_argument('--max-rows',       type=int,   default=None)
    args = parser.parse_args()

    api_key = os.environ.get('DART_API_BUSINESS_KEY') or os.environ.get('OPENDART_API_KEY')
    if not api_key:
        print("❌ DART_API_BUSINESS_KEY 없음"); return

    xml_dir = Path(args.xml_dir)
    df = pd.read_csv(
        args.input,
        encoding='utf-8-sig',
        dtype={'rcept_no': str, 'corp_code': str, 'stock_code': str},
    )
    if args.max_rows:
        df = df.head(args.max_rows)

    print(f"\n{'='*60}")
    print(f"Phase 3 v2: XML 다운로드 및 피처 추출")
    print(f"{'='*60}")
    print(f"대상: {len(df)}개 | XML 저장: {xml_dir}/")

    records = []
    stats   = {'dl_ok': 0, 'dl_fail': 0,
                'ext_ok': 0, 'ext_partial': 0, 'ext_fail': 0}

    for i, (_, row) in enumerate(df.iterrows()):
        company  = row['company_name']
        rcept_no = normalize_rcept_no(row.get('rcept_no'))
        print(f"\n[{i+1:3d}/{len(df)}] {company} ({rcept_no})")

        record = {
            'company_name':           company,
            'listing_date':           row['listing_date'],
            'corp_code':              row['corp_code'],
            'rcept_no':               rcept_no,
            'report_nm':              row.get('report_nm', ''),
            'track_type':             row.get('track_type', ''),
            'discount_rate_low_pct':  row.get('discount_rate_low_pct'),
            'discount_rate_high_pct': row.get('discount_rate_high_pct'),
            'discount_rate_mid_pct':  row.get('discount_rate_mid_pct'),
        }

        if not rcept_no:
            stats['dl_fail'] += 1
            record['extraction_status'] = 'no_rcept_no'
            records.append(record)
            print(f"  ⚠️ rcept_no 없음 - 다운로드 스킵")
            continue

        # 다운로드
        if not args.skip_download:
            xml_path = download_xml(api_key, rcept_no, xml_dir, args.rate_limit_sec)
        else:
            xml_path = xml_dir / f"{rcept_no}.xml"
            xml_path = xml_path if xml_path.exists() else None

        if xml_path is None:
            stats['dl_fail'] += 1
            print(f"  ❌ 다운로드 실패")
            record['extraction_status'] = 'download_failed'
            records.append(record); continue

        stats['dl_ok'] += 1

        # 추출
        plain = xml_to_plain(xml_path)
        feats = extract(plain)
        record.update(feats)
        record['xml_size'] = xml_path.stat().st_size

        status = feats.get('extraction_status', 'fail')
        if status == 'ok':
            stats['ext_ok'] += 1
            print(f"  ✅ 밴드: {_fmt_num(feats.get('price_band_low'))}~{_fmt_num(feats.get('price_band_high'))}"
                  f" | 확정: {_fmt_num(feats.get('final_offering_price'))}"
                  f" | PER: {feats.get('peer_per')}"
                  f" | 순이익: {feats.get('net_income_bn')}억")
        elif status == 'partial':
            stats['ext_partial'] += 1
            filled = feats.get('filled_key_fields', 0)
            print(f"  ⚠️  부분추출 ({filled}/5) — "
                  f"밴드={feats.get('price_band_low')}/{feats.get('price_band_high')} "
                  f"확정={feats.get('final_offering_price')} "
                  f"PER={feats.get('peer_per')} 순이익={feats.get('net_income_bn')}")
        else:
            stats['ext_fail'] += 1
            print(f"  ❌ 추출실패")

        # ── 역산 할인율 보완 ────────────────────────────────
        # XML 할인율이 없거나 불일치 시: 1 - final/per_stock_value 로 역산
        fp  = feats.get('final_offering_price')
        psv = feats.get('per_stock_value')
        if fp and psv and psv > 0 and fp >= 100:
            implied = round((1 - fp / psv) * 100, 4)
            record['discount_rate_implied'] = implied  # 역산 할인율 (참고용)

        # ── 교차검증: XML 할인율 vs 한라캐스트 할인율 ────────
        xml_mid   = feats.get('discount_rate_xml_mid')
        halla_mid = record.get('discount_rate_mid_pct')

        # XML 할인율이 없으면 역산값으로 대체 시도
        if xml_mid is None and record.get('discount_rate_implied') is not None:
            implied = record['discount_rate_implied']
            # 역산값이 합리적 범위(5~70%)이면 사용
            if 5 <= implied <= 70:
                record['discount_rate_xml_low']  = None
                record['discount_rate_xml_high'] = None
                record['discount_rate_xml_mid']  = implied
                record['discount_rate_xml_source'] = 'implied'
                xml_mid = implied

        if xml_mid is not None and pd.notna(halla_mid):
            diff = round(abs(xml_mid - float(halla_mid)), 4)
            record['discount_rate_diff'] = diff
            src = record.get('discount_rate_xml_source', 'xml')
            if diff > 3.0:   # 3%p 초과: 심각한 불일치 → 한라 데이터 오기재 의심
                stats['xval_mismatch'] = stats.get('xval_mismatch', 0) + 1
                record['data_quality_flag'] = 'MISMATCH_CHECK'
                print(f"  🚨 할인율 심각 불일치({src}): XML={xml_mid}% vs 한라={halla_mid}% (차이={diff}%p) → 수동확인 필요")
            elif diff > 1.0:   # 1~3%p: 경미한 불일치 (정정전 값 캡처 등)
                stats['xval_mismatch'] = stats.get('xval_mismatch', 0) + 1
                record['data_quality_flag'] = 'MINOR_DIFF'
                print(f"  ⚠️  할인율 불일치({src}): XML={xml_mid}% vs 한라={halla_mid}% (차이={diff}%p)")
            else:
                stats['xval_match'] = stats.get('xval_match', 0) + 1
                record['data_quality_flag'] = 'OK'
                print(f"  🔍 할인율 검증({src}): XML={xml_mid}% ≈ 한라={halla_mid}% ✓")
        elif xml_mid is None:
            record['discount_rate_diff'] = None
            record['data_quality_flag'] = 'NO_XML_DR'
            stats['xval_no_xml'] = stats.get('xval_no_xml', 0) + 1
            print(f"  ℹ️  XML 할인율 미추출")

        records.append(record)

        # 10건마다 중간저장
        if (i + 1) % 10 == 0:
            pd.DataFrame(records).to_csv(
                args.output.replace('.csv', '_partial.csv'),
                index=False, encoding='utf-8-sig'
            )
            print(f"  💾 중간저장 ({i+1}건)")

    # 최종 저장
    out_df = pd.DataFrame(records)
    out_df.to_csv(args.output, index=False, encoding='utf-8-sig')

    # 로그 저장
    log_cols = ['company_name', 'rcept_no', 'report_nm', 'track_type',
                'extraction_status', 'filled_key_fields', 'xml_size',
                'price_band_low', 'price_band_high', 'final_offering_price',
                'net_income_bn', 'peer_per', 'per_stock_value',
                'offering_shares', 'listing_market']
    log_df = out_df[[c for c in log_cols if c in out_df.columns]]
    log_df.to_csv(args.output.replace('.csv', '_log.csv'), index=False, encoding='utf-8-sig')

    # 결과 요약
    print(f"\n{'='*60}")
    print(f"Phase 3 완료")
    print(f"{'='*60}")
    print(f"  다운로드 성공:    {stats['dl_ok']:3d}개")
    print(f"  다운로드 실패:    {stats['dl_fail']:3d}개")
    print(f"  추출 완료 (ok):   {stats['ext_ok']:3d}개")
    print(f"  추출 부분:        {stats['ext_partial']:3d}개")
    print(f"  추출 실패:        {stats['ext_fail']:3d}개")
    print(f"\n저장: {args.output}")

    # 교차검증 결과 요약
    print(f"\n[할인율 교차검증]")
    print(f"  일치 (diff ≤ 1%p): {stats.get('xval_match', 0):3d}개")
    print(f"  불일치 (diff > 1%p): {stats.get('xval_mismatch', 0):3d}개")
    print(f"  XML 미추출:          {stats.get('xval_no_xml', 0):3d}개")

    # 불일치 상세
    if 'discount_rate_diff' in out_df.columns:
        mismatch_df = out_df[out_df['discount_rate_diff'].notna() &
                             (out_df['discount_rate_diff'] > 1.0)]
        if len(mismatch_df) > 0:
            print(f"\n  [불일치 상세]")
            for _, r in mismatch_df.iterrows():
                print(f"  ⚠️  {r['company_name']:15s} | "
                      f"XML={r.get('discount_rate_xml_mid')}% | "
                      f"한라={r.get('discount_rate_mid_pct')}% | "
                      f"차이={r['discount_rate_diff']}%p")

    feat_cols = ['price_band_low', 'price_band_high', 'final_offering_price',
                 'band_midpoint', 'price_vs_mid', 'net_income_bn', 'peer_per',
                 'per_stock_value', 'total_shares', 'offering_shares',
                 'offering_amount', 'institutional_ratio', 'listing_market',
                 'discount_rate_xml_low', 'discount_rate_xml_high',
                 'discount_rate_xml_mid', 'discount_rate_diff']
    print(f"\n[피처 커버리지]")
    for col in feat_cols:
        if col in out_df.columns:
            n   = out_df[col].notna().sum()
            pct = n / len(out_df) * 100
            bar = '█' * int(pct / 5) + '░' * (20 - int(pct / 5))
            print(f"  {col:25s}: {bar} {n:3d}/{len(out_df)} ({pct:.0f}%)")


if __name__ == '__main__':
    main()
