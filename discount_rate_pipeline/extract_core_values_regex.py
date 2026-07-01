"""
증권신고서 PDF 본문 텍스트에서 핵심 7개 수치를 정규식으로 직접 추출

옵션 C 접근법:
- 표 구조 파싱 대신 본문 텍스트의 키워드+숫자 패턴 매칭
- 한라캐스트뿐만 아니라 표 형식이 약간 다른 다른 회사 증권신고서에도 적용 가능
- pdfplumber로 페이지별 텍스트만 뽑고, 정규식으로 7개 핵심 값 추출

추출 대상 (한라캐스트 ground truth):
  applied_net_income_million_won: 9910
  applied_shares_total: 36526017
  applied_per: 25.93
  intrinsic_value_per_share: 7033
  discount_rate_low_pct: 17.54
  discount_rate_high_pct: 27.49
  offering_band_low: 5100
  offering_band_high: 5800
"""
import re
import argparse
import json
from typing import Dict, Optional
from pathlib import Path


def clean_text(s: str) -> str:
    """공백·줄바꿈 정규화"""
    return re.sub(r'\s+', ' ', s).strip()


def extract_full_text(pdf_path: str) -> str:
    """PDF에서 전체 텍스트 추출"""
    import pdfplumber
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            pages_text.append(t)
    return "\n".join(pages_text)


def parse_number(s: str) -> Optional[float]:
    """'9,910' → 9910, '17.54' → 17.54"""
    if not s:
        return None
    s = s.replace(',', '').replace(' ', '').strip()
    try:
        return float(s)
    except ValueError:
        return None


# ============ 정규식 패턴 ============
# 각 패턴은 (필드명, 정규식, 후처리함수) 튜플
# 정규식에서 그룹 1이 추출할 숫자
PATTERNS = [
    # 적용 당기순이익: "9,910 백만원" 또는 "9,910백만원"
    (
        'applied_net_income_million_won',
        re.compile(r'적용\s*당기순이익\s*[:\s]*([0-9,]+)\s*백만원'),
        parse_number,
    ),
    # 적용 PER: "25.93 배" 또는 "25.93배"
    (
        'applied_per',
        re.compile(r'적용\s*PER\s*[:\s]*([0-9.]+)\s*배'),
        parse_number,
    ),
    # 적용 주식수: "36,526,017 주" 또는 "36,526,017주"
    (
        'applied_shares_total',
        re.compile(r'적용\s*주식수\s*[:\s]*([0-9,]+)\s*주'),
        parse_number,
    ),
    # 주당 평가가액: "7,033 원" (주의: "주당 평가가액"이 여러 번 등장하므로 컨텍스트 필요)
    (
        'intrinsic_value_per_share',
        re.compile(r'주당\s*평가가액\s*[:\s]*([0-9,]+)\s*원'),
        parse_number,
    ),
]


# 할인율 범위: "17.54% ~ 27.49%" 또는 "27.49% ~ 17.54%" (양방향)
DISCOUNT_PATTERN = re.compile(
    r'평가액\s*대비\s*할인율\s*[:\s]*([0-9.]+)\s*%\s*[~～\-]\s*([0-9.]+)\s*%'
)

# 희망공모가액 밴드: "5,100 원 ~ 5,800 원"
OFFERING_BAND_PATTERN = re.compile(
    r'희망공모가액\s*밴드\s*[:\s]*([0-9,]+)\s*원\s*[~～\-]\s*([0-9,]+)\s*원'
)


def extract_core_values(text: str) -> Dict:
    """본문 텍스트에서 핵심 7개 값 추출"""
    text_clean = clean_text(text)
    result = {}
    
    # 1. 단일 값 패턴들
    for field, pattern, postprocess in PATTERNS:
        match = pattern.search(text_clean)
        if match:
            value = postprocess(match.group(1))
            if value is not None:
                result[field] = value
    
    # 2. 할인율 범위 (low/high 자동 정규화: 작은 값=low, 큰 값=high)
    m = DISCOUNT_PATTERN.search(text_clean)
    if m:
        v1 = parse_number(m.group(1))
        v2 = parse_number(m.group(2))
        if v1 is not None and v2 is not None:
            result['discount_rate_low_pct'] = min(v1, v2)
            result['discount_rate_high_pct'] = max(v1, v2)
    
    # 3. 희망공모가액 밴드
    m = OFFERING_BAND_PATTERN.search(text_clean)
    if m:
        v1 = parse_number(m.group(1))
        v2 = parse_number(m.group(2))
        if v1 is not None and v2 is not None:
            result['offering_band_low'] = min(v1, v2)
            result['offering_band_high'] = max(v1, v2)
    
    return result


# ============ 검증 ============
HALLA_CAST_GT = {
    'applied_net_income_million_won': 9910,
    'applied_shares_total': 36526017,
    'applied_per': 25.93,
    'intrinsic_value_per_share': 7033,
    'discount_rate_low_pct': 17.54,
    'discount_rate_high_pct': 27.49,
    'offering_band_low': 5100,
    'offering_band_high': 5800,
}


def validate(extracted: Dict, gt: Dict, tolerance: float = 0.01) -> Dict:
    """ground truth와 비교"""
    result = {'match': [], 'mismatch': [], 'missing': []}
    for key, gt_val in gt.items():
        ex_val = extracted.get(key)
        if ex_val is None:
            result['missing'].append(key)
            continue
        try:
            if abs(float(ex_val) - float(gt_val)) / max(abs(gt_val), 1) <= tolerance:
                result['match'].append({'key': key, 'value': ex_val})
            else:
                result['mismatch'].append({
                    'key': key, 'extracted': ex_val, 'gt': gt_val
                })
        except (ValueError, TypeError):
            result['mismatch'].append({
                'key': key, 'extracted': ex_val, 'gt': gt_val, 'note': 'type_error'
            })
    accuracy = len(result['match']) / len(gt) * 100
    result['accuracy_pct'] = accuracy
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--pdf', required=True)
    parser.add_argument('--output', default=None)
    parser.add_argument('--validate', action='store_true')
    parser.add_argument('--debug', action='store_true', help='추출 텍스트 일부 출력')
    args = parser.parse_args()
    
    print(f"\n{'='*60}")
    print(f"PDF 핵심 수치 추출 (정규식 기반)")
    print(f"{'='*60}")
    print(f"입력: {args.pdf}")
    
    text = extract_full_text(args.pdf)
    print(f"전체 텍스트 길이: {len(text):,}자")
    
    if args.debug:
        # 핵심 키워드 주변 텍스트 출력
        for kw in ['적용 당기순이익', '적용 PER', '주당 평가가액', '할인율', '희망공모가액']:
            idx = text.find(kw)
            if idx >= 0:
                print(f"\n[{kw}] 발견 위치 {idx}")
                print(f"  context: ...{text[max(0,idx-30):idx+120]}...")
    
    result = extract_core_values(text)
    print(f"\n추출 결과 ({len(result)}/8 필드):")
    for k, v in result.items():
        print(f"  {k}: {v}")
    
    if args.validate:
        print(f"\n{'='*60}")
        print(f"한라캐스트 ground truth 검증")
        print(f"{'='*60}")
        v = validate(result, HALLA_CAST_GT)
        print(f"정확도: {v['accuracy_pct']:.0f}% ({len(v['match'])}/{len(HALLA_CAST_GT)})")
        
        if v['mismatch']:
            print(f"\n불일치:")
            for m in v['mismatch']:
                print(f"  {m['key']}: 추출 {m['extracted']} vs GT {m['gt']}")
        if v['missing']:
            print(f"\n누락: {v['missing']}")
    
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"\n저장: {args.output}")


if __name__ == '__main__':
    main()
