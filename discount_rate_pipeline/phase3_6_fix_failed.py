"""
Phase 3.6: 실패 8건 올바른 rcpNo로 재추출

원인: 같은 날 2개의 발행조건확정 신고서 중 빈 문서(경량본)를 가져온 것
해결: 실제 내용이 있는 다른 rcpNo로 교체 후 재다운로드 및 피처 추출

입력:  ipo_features_raw.csv
출력:  ipo_features_raw.csv (인플레이스 수정)

사용법:
  python phase3_6_fix_failed.py --input ipo_features_raw.csv --xml-dir xml
"""
import os, io, re, time, zipfile, argparse, logging
import requests
import pandas as pd
import numpy as np
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

DART_DOC_URL = "https://opendart.fss.or.kr/api/document.xml"

# ── 올바른 rcpNo 매핑 ─────────────────────────────────────
CORRECT_RCPNO = {
    '듀켐바이오':    '20241209000393',
    '탑런토탈솔루션': '20241022000128',
    'HB인베스트먼트': '20240115000350',
    '스톰테크':      '20231108000232',
    '뷰티스킨':      '20230616000415',  # [첨부추가]증권신고서
    'LB인베스트먼트': '20230316000293',
    '필에너지':      '20230704000142',
    '빅텐츠':        None,  # 별도 처리 필요
}


def download_xml(api_key, rcept_no, save_dir, rate_limit=0.5):
    save_dir.mkdir(parents=True, exist_ok=True)
    dest = save_dir / f"{rcept_no}.xml"
    if dest.exists() and dest.stat().st_size > 1000:
        log.info(f"{rcept_no}: 캐시 사용 ({dest.stat().st_size:,}B)")
        return dest
    url = f"{DART_DOC_URL}?crtfc_key={api_key}&rcept_no={rcept_no}"
    try:
        r = requests.get(url, timeout=60)
        time.sleep(rate_limit)
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
    except Exception as e:
        log.error(f"{rcept_no}: {e}")
        return None


def xml_to_plain(path):
    for enc in ('utf-8', 'euc-kr', 'cp949'):
        try:
            raw = path.read_text(encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        return ""
    plain = re.sub(r'<[^>]+>', ' ', raw)
    return re.sub(r'\s+', ' ', plain)


def _num(s):
    if not s: return None
    s = re.sub(r'[,\s원주배%]', '', str(s).strip())
    s = re.sub(r'\(?\s*주\s*\d+\s*\)?', '', s)
    try: return float(s)
    except: return None


def extract(plain):
    """phase3_extract_features.py의 extract() 함수와 동일한 로직."""
    f = {}

    # 공모가 밴드
    m = re.search(
        r'(?:모집\(매출\)가액\(예정\)|공모희망가액)[^\d]*([\d,]+)원\s*~\s*([\d,]+)원',
        plain
    )
    if not m:
        m = re.search(
            r'(?:모집\(매출\)가액\(예정\))\s*:\s*([\d,.]+)\s*원\s*~\s*([\d,.]+)\s*원',
            plain
        )
    if m:
        low, high = _num(m.group(1)), _num(m.group(2))
        if low and high and low >= 100 and high >= 100:
            f['price_band_low']  = low
            f['price_band_high'] = high
            f['band_midpoint']   = round((low + high) / 2, 2)

    # 확정공모가
    final = None
    for pat in [
        r'확정\s*공모가액인\s*([\d,]+)원',
        r'확정공모가액을\s*([\d,]+)원으로',
        r'1주당\s*확정공모가액을\s*([\d,]+)원',
        r'확정\s*공모가액\s*[：:]\s*([\d,]+)원',
    ]:
        m = re.search(pat, plain)
        if m:
            v = _num(m.group(1))
            if v and v >= 100:
                final = v; break

    if final is None:
        after = plain.find('정정 후')
        if after == -1: after = plain.find('정정후')
        area = plain[after:after+2000] if after >= 0 else plain[:3000]
        m = re.search(r'모집\(매출\)가액\(예정\)\s*:\s*([\d,]+)원(?!\s*~)', area)
        if m:
            v = _num(m.group(1))
            if v and v >= 100: final = v

    f['final_offering_price'] = final
    if final and f.get('band_midpoint'):
        f['price_vs_mid'] = round(final / f['band_midpoint'] - 1, 4)

    # 순이익
    m = re.search(
        r'(?:지배기업\s*소유주\s*지분\s*연결)?(?:지배주주\s*)?당기순이익[^\d]*([\d,]+)\s*(백만원|천원)',
        plain
    )
    if m:
        v, unit = _num(m.group(1)), m.group(2)
        if v:
            f['net_income_mn'] = round(v / (1 if unit=='백만원' else 1000), 4)
            f['net_income_bn'] = round(v / (100 if unit=='백만원' else 100_000), 4)

    # PER
    m = re.search(r'(?:유사기업|비교대상회사)\s*PER\s*([\d,.]+)배', plain)
    if m:
        v = _num(m.group(1))
        if v: f['peer_per'] = v

    # EV/EBITDA
    if f.get('peer_per') is None and 'EV/EBITDA' in plain:
        f['valuation_model'] = 'EV/EBITDA'
        m2 = re.search(r'(?:유사기업|비교기업|비교대상회사)[^\d\n]*EV/EBITDA\s*배?\s*([\d,.]+)', plain)
        if m2:
            v = _num(m2.group(1))
            if v: f['peer_per'] = v
        m3 = re.search(r'EBITDA\s*(천원|백만원)\s*([\d,]+)', plain)
        if m3:
            v, unit = _num(m3.group(2)), m3.group(1)
            if v:
                f['net_income_bn'] = round(v / (100 if unit=='백만원' else 100_000), 4)
                f['net_income_type'] = 'EBITDA'
    else:
        f['valuation_model'] = 'PER'

    # 주당 평가가액
    m = re.search(r'주당\s*평가\s*가\s*액\s*([\d,]+)원', plain)
    if m:
        v = _num(m.group(1))
        if v and v >= 100: f['per_stock_value'] = v

    # 주식수
    m = re.search(r'③\s*주식수\s*([\d,]+)\s*주', plain)
    if not m:
        m = re.search(r'⑥\s*주식수\s*주\s*([\d,]+)', plain)
    if not m:
        m = re.search(r'주식수\s*주\s*([\d,]+)', plain)
    if m:
        v = _num(m.group(1))
        if v: f['total_shares'] = v

    # 공모주식수
    m = re.search(r'기명식\s*보통주\s*([\d,]+)주', plain)
    if m:
        v = _num(m.group(1))
        if v: f['offering_shares'] = v

    # 공모총액
    m = re.search(r'모집\s*또는\s*매출금액\s*:\s*([\d,]+)원', plain)
    if m:
        v = _num(m.group(1))
        if v: f['offering_amount'] = v

    # 할인율
    for pat in [
        r'주당\s*평가\s*가액에\s*대한\s*할인율\s*([\d.]+)%\s*~\s*([\d.]+)%',
        r'평가액\s*대비\s*할인율(?:\(주\d+\))?\s*([\d.]+)%\s*~\s*([\d.]+)%',
        r'할인율\s*\([④⑤⑦]\)\s*([\d.]+)%\s*~\s*([\d.]+)%',
    ]:
        m = re.search(pat, plain)
        if m:
            v1, v2 = float(m.group(1)), float(m.group(2))
            f['discount_rate_xml_low']  = round(min(v1, v2), 4)
            f['discount_rate_xml_high'] = round(max(v1, v2), 4)
            f['discount_rate_xml_mid']  = round((v1 + v2) / 2, 4)
            break

    # 상장 정보
    f['listing_market']    = '코스닥' if '코스닥' in plain[:8000] else \
                             ('코스피' if '코스피' in plain[:8000] else None)
    f['is_tech_special']   = int(bool(re.search(r'기술특례|기술성장기업', plain[:8000])))
    f['is_growth_special'] = int(bool(re.search(r'성장성특례', plain[:8000])))

    key_fields = ['price_band_low','price_band_high','final_offering_price',
                  'net_income_bn','peer_per']
    filled = sum(1 for k in key_fields if f.get(k) is not None)
    f['filled_key_fields'] = filled
    f['extraction_status'] = 'ok' if filled >= 4 else \
                             ('partial' if filled >= 2 else 'fail')
    return f


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input',          required=True)
    parser.add_argument('--xml-dir',        default='xml')
    parser.add_argument('--rate-limit-sec', type=float, default=0.5)
    args = parser.parse_args()

    api_key = os.environ.get('DART_API_BUSINESS_KEY') or os.environ.get('OPENDART_API_KEY')
    if not api_key:
        print("❌ API 키 없음"); return

    xml_dir = Path(args.xml_dir)
    df = pd.read_csv(args.input, encoding='utf-8-sig', dtype={'rcept_no': str, 'corp_code': str})

    print(f"\n{'='*60}")
    print(f"Phase 3.6: 실패 8건 올바른 rcpNo 재추출")
    print(f"{'='*60}")

    stats = {'ok': 0, 'partial': 0, 'fail': 0}

    for company, new_rcpno in CORRECT_RCPNO.items():
        idx_list = df.index[df['company_name'] == company].tolist()
        if not idx_list:
            print(f"\n[{company}] 데이터에서 찾을 수 없음 — 스킵")
            continue

        idx = idx_list[0]
        old_rcpno = df.loc[idx, 'rcept_no']
        halla_mid = df.loc[idx, 'discount_rate_mid_pct']

        print(f"\n[{company}]")
        print(f"  기존 rcpNo: {old_rcpno} → 신규: {new_rcpno}")

        if new_rcpno is None:
            print(f"  ⚠️  올바른 rcpNo 미확정 — 스킵")
            continue

        # 다운로드
        xml_path = download_xml(api_key, new_rcpno, xml_dir, args.rate_limit_sec)
        if xml_path is None:
            print(f"  ❌ 다운로드 실패")
            stats['fail'] += 1
            continue

        print(f"  📄 {xml_path.name} ({xml_path.stat().st_size:,}B)")

        # 추출
        plain = xml_to_plain(xml_path)
        feats = extract(plain)

        # DataFrame 업데이트
        df.loc[idx, 'rcept_no']          = new_rcpno
        df.loc[idx, 'xml_size']          = xml_path.stat().st_size
        df.loc[idx, 'extraction_status'] = feats.get('extraction_status', 'fail')
        df.loc[idx, 'filled_key_fields'] = feats.get('filled_key_fields', 0)

        for col, val in feats.items():
            if col not in ('extraction_status', 'filled_key_fields') and val is not None:
                df.loc[idx, col] = val

        status = feats.get('extraction_status')
        if status == 'ok':
            stats['ok'] += 1
            print(f"  ✅ 밴드: {feats.get('price_band_low')}~{feats.get('price_band_high')}"
                  f" | 확정: {feats.get('final_offering_price')}"
                  f" | PER: {feats.get('peer_per')}"
                  f" | 순이익: {feats.get('net_income_bn')}억")
        elif status == 'partial':
            stats['partial'] += 1
            print(f"  ⚠️  부분추출 ({feats.get('filled_key_fields')}/5)"
                  f" 밴드={feats.get('price_band_low')}/{feats.get('price_band_high')}"
                  f" 확정={feats.get('final_offering_price')}")
        else:
            stats['fail'] += 1
            print(f"  ❌ 추출 실패")

        # 교차검증
        xml_mid  = feats.get('discount_rate_xml_mid')
        if xml_mid and halla_mid:
            diff = round(abs(xml_mid - float(halla_mid)), 4)
            df.loc[idx, 'discount_rate_diff'] = diff
            flag = '✓' if diff <= 1.0 else f'⚠️  차이={diff}%p'
            print(f"  🔍 할인율: XML={xml_mid}% vs 한라={halla_mid}% {flag}")

    # 저장 (인플레이스)
    df.to_csv(args.input, index=False, encoding='utf-8-sig')

    print(f"\n{'='*60}")
    print(f"Phase 3.6 완료")
    print(f"{'='*60}")
    print(f"  성공(ok):    {stats['ok']}건")
    print(f"  부분(partial): {stats['partial']}건")
    print(f"  실패:        {stats['fail']}건")
    print(f"\n저장: {args.input} (인플레이스 업데이트)")

    # 전체 커버리지 재확인
    print(f"\n[업데이트 후 전체 커버리지]")
    for col in ['price_band_low','final_offering_price','price_vs_mid',
                'net_income_bn','peer_per','discount_rate_xml_mid']:
        if col in df.columns:
            n = df[col].notna().sum()
            pct = n / len(df) * 100
            bar = '█' * int(pct/5) + '░' * (20-int(pct/5))
            print(f"  {col:25s}: {bar} {n:3d}/{len(df)} ({pct:.0f}%)")


if __name__ == '__main__':
    main()