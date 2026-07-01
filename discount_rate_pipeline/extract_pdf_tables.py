"""
증권신고서 PDF에서 평가가액·할인율 표 추출

3단계 추출 전략 (난이도 순):
1. pdfplumber로 표 직접 추출 (가장 정확, 그러나 표 구조에 민감)
2. Camelot로 추출 (선/공백 인식 더 robust)
3. Gemini Vision으로 표 페이지 이미지 파싱 (마지막 수단, 가장 비싸지만 가장 robust)

추출 대상 4개 표:
  표A: PER에 의한 평가가치 (적용 당기순이익, 적용 PER, 주당평가가액)
  표B: 적용 당기순이익 산정 (일회성 손익 가감)
  표C: 적용 주식수 산정 (발행+공모+의무인수)
  표D: 희망공모가액 산출내역 (할인율 밴드 ← ⭐ 라벨)

사용 예:
    python extract_pdf_tables.py --rcpno 20250408000123 --output ipo_001.json
    또는 일괄:
    python extract_pdf_tables.py --batch ipo_metadata.csv --output extracted.csv
"""
import os
import re
import json
import argparse
from pathlib import Path
from typing import Dict, List, Optional

# ============ 추출 대상 표 시그니처 ============
TABLE_SIGNATURES = {
    'table_A_valuation': {
        'header_keywords': ['주당 평가가액 산출', 'PER에 의한', '평가가치'],
        'expected_rows': ['적용 당기순이익', '적용 주식수', '적용 주당 순이익', '적용 PER', '주당 평가가액'],
        'output_fields': ['applied_net_income', 'applied_shares_total', 'applied_eps', 'applied_per', 'intrinsic_value_per_share'],
    },
    'table_B_net_income': {
        'header_keywords': ['적용 당기순이익', '산정 근거'],
        'expected_rows': ['당기순이익', 'CB', 'RCPS', '일회성'],
        'output_fields': ['base_net_income', 'one_time_adjustments', 'applied_net_income'],
    },
    'table_C_shares': {
        'header_keywords': ['적용 주식수 산정', '발행주식'],
        'expected_rows': ['발행주식총수', '공모', '의무인수', '합계'],
        'output_fields': ['issued_shares_total', 'public_offering_shares', 'mandatory_subscription', 'applied_shares_total'],
    },
    'table_D_offering_band': {
        'header_keywords': ['희망공모가액', '산출내역', '할인율'],
        'expected_rows': ['주당 평가가액', '평가액 대비 할인율', '희망공모가액 밴드', '확정 주당 공모가액'],
        'output_fields': ['intrinsic_value_per_share', 'discount_rate_low', 'discount_rate_high', 'offering_band_low', 'offering_band_high', 'final_offering_price'],
    },
}

# ============ 한라캐스트 검증용 Ground Truth ============
HALLA_CAST_GT = {
    'table_A_valuation': {
        'applied_net_income': 9910_000_000,  # 9,910 백만원
        'applied_shares_total': 36_526_017,
        'applied_eps': 271,
        'applied_per': 25.93,
        'intrinsic_value_per_share': 7033,
    },
    'table_C_shares': {
        'issued_shares_total': 28_829_939,
        'public_offering_shares': 7_500_000,
        'mandatory_subscription': 196_078,
        'applied_shares_total': 36_526_017,
    },
    'table_D_offering_band': {
        'intrinsic_value_per_share': 7033,
        'discount_rate_low': 17.54,
        'discount_rate_high': 27.49,
        'offering_band_low': 5100,
        'offering_band_high': 5800,
        'final_offering_price': None,  # 미정
    },
}

# ============ Strategy 1: pdfplumber ============
def extract_with_pdfplumber(pdf_path: str) -> List[Dict]:
    """pdfplumber로 표 추출. 가장 정확하지만 표 구조가 명확해야 함."""
    try:
        import pdfplumber
    except ImportError:
        print("⚠️  pdfplumber 미설치: pip install pdfplumber")
        return []
    
    extracted = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            
            # 표 시그니처와 매칭되는 페이지만 처리
            for sig_name, sig in TABLE_SIGNATURES.items():
                if any(kw in text for kw in sig['header_keywords']):
                    tables = page.extract_tables()
                    for tbl in tables:
                        extracted.append({
                            'signature': sig_name,
                            'page': page_idx + 1,
                            'table': tbl,
                            'method': 'pdfplumber',
                        })
    return extracted

# ============ Strategy 2: Camelot ============
def extract_with_camelot(pdf_path: str) -> List[Dict]:
    """Camelot로 추출. pdfplumber 실패 시 fallback."""
    try:
        import camelot
    except ImportError:
        print("⚠️  Camelot 미설치: pip install camelot-py[cv]")
        return []
    
    extracted = []
    try:
        # lattice = 선이 있는 표, stream = 공백 기반 표
        for flavor in ['lattice', 'stream']:
            tables = camelot.read_pdf(pdf_path, pages='all', flavor=flavor)
            for i, tbl in enumerate(tables):
                df = tbl.df
                text = df.to_string()
                for sig_name, sig in TABLE_SIGNATURES.items():
                    if any(kw in text for kw in sig['header_keywords']):
                        extracted.append({
                            'signature': sig_name,
                            'page': tbl.page,
                            'table': df.values.tolist(),
                            'method': f'camelot_{flavor}',
                        })
    except Exception as e:
        print(f"  Camelot 오류: {e}")
    return extracted

# ============ Strategy 3: Gemini Vision (최후 수단) ============
def extract_with_gemini(pdf_path: str, page_numbers: List[int] = None) -> List[Dict]:
    """Gemini Vision으로 PDF 페이지 이미지를 직접 파싱.
    가장 비싸지만 가장 robust. 표 구조가 깨진 경우 사용.
    """
    try:
        import google.generativeai as genai
        from pdf2image import convert_from_path
    except ImportError:
        print("⚠️  필요 패키지: pip install google-generativeai pdf2image")
        return []
    
    api_key = os.environ.get('GEMINI_API_KEY')
    if not api_key:
        print("⚠️  GEMINI_API_KEY 환경변수 필요")
        return []
    
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.0-flash')
    
    # PDF → 이미지
    images = convert_from_path(pdf_path, dpi=200, 
                                first_page=min(page_numbers) if page_numbers else None,
                                last_page=max(page_numbers) if page_numbers else None)
    
    prompt = """
이 페이지에는 IPO 증권신고서의 평가가액 또는 희망공모가액 산출 표가 있을 수 있습니다.
다음 정보가 있으면 JSON으로 추출해주세요. 없으면 해당 키를 생략하세요.

{
  "applied_net_income_million_won": (적용 당기순이익, 백만원 단위 숫자),
  "applied_shares_total": (적용 주식수),
  "applied_per": (적용 PER, 소수점),
  "intrinsic_value_per_share": (주당 평가가액, 원 단위),
  "discount_rate_low_pct": (평가액 대비 할인율 하단, % 숫자만),
  "discount_rate_high_pct": (평가액 대비 할인율 상단, % 숫자만),
  "offering_band_low": (희망공모가액 밴드 하단, 원),
  "offering_band_high": (희망공모가액 밴드 상단, 원)
}

표가 없거나 위 정보가 없으면 빈 객체 {}를 반환하세요.
JSON만 출력, 다른 텍스트 없이.
"""
    
    extracted = []
    for i, img in enumerate(images):
        try:
            response = model.generate_content([prompt, img])
            text = response.text.strip()
            # JSON 파싱
            text = re.sub(r'^```json\s*', '', text)
            text = re.sub(r'\s*```$', '', text)
            data = json.loads(text)
            if data:
                extracted.append({
                    'page': i + 1,
                    'extracted': data,
                    'method': 'gemini_vision',
                })
        except Exception as e:
            print(f"  Gemini 페이지 {i+1} 오류: {e}")
    
    return extracted

# ============ 정규화 + 검증 ============
def parse_korean_number(s: str) -> Optional[float]:
    """'9,910 백만원' → 9910, '17.54%' → 17.54"""
    if not s:
        return None
    s = str(s).strip()
    s = s.replace(',', '').replace(' ', '')
    s = re.sub(r'(백만원|원|배|%|％|주)', '', s)
    s = s.replace('～', '~').replace('-', '~')  # 범위 구분자 통일
    
    try:
        return float(s)
    except ValueError:
        return None

def _extract_nums_from_cells(cells: List[str]) -> List[float]:
    """셀 목록에서 숫자 추출. 범위값(17.54%~27.49%) 분리 포함."""
    nums = []
    for cell in cells:
        s = str(cell or "").strip()
        if not s:
            continue
        # 추가 : 각주 표시 제거 (주1), (주2), 주1) 등
        s = re.sub(r'\(?\s*주\s*\d+\s*\)', '', s)

        unit_scale = 1_000_000 if '백만원' in s else 1
        cleaned = re.sub(r'[백만원주배%％,\s]', '', s).replace('～', '~')
        if '~' in cleaned:
            for part in cleaned.split('~'):
                try:
                    nums.append(float(part) * unit_scale)
                except ValueError:
                    pass
        else:
            for m in re.findall(r'-?\d+(?:\.\d+)?', cleaned):
                try:
                    nums.append(float(m) * unit_scale)
                except ValueError:
                    pass
    return nums


def _parse_table_by_signature(table: List[List], sig_name: str) -> Dict:
    """pdfplumber/camelot raw table(list of lists) → 표준 키 Dict"""
    result = {}
    for row in (table or []):
        if not row:
            continue
        cells = [str(c or "").strip() for c in row]
        row_text = " ".join(cells)
        # 첫 셀 이후를 값으로 시도, 없으면 전체 셀에서 추출
        nums = _extract_nums_from_cells(cells[1:]) or _extract_nums_from_cells(cells)

        if sig_name == 'table_A_valuation':
            if '적용 당기순이익' in row_text and nums:
                result['applied_net_income'] = nums[0]
            if '적용 주식수' in row_text and '주당' not in row_text and nums:
                result['applied_shares_total'] = nums[0]
            if '적용 주당 순이익' in row_text and nums:
                result['applied_eps'] = nums[0]
            if '적용 PER' in row_text and nums:
                result['applied_per'] = nums[0]
            if '주당 평가가액' in row_text and '적용' not in row_text and nums:
                result['intrinsic_value_per_share'] = nums[0]

        elif sig_name == 'table_B_net_income':
            if '당기순이익' in row_text and 'CB' not in row_text and nums:
                result.setdefault('base_net_income', nums[0])

        elif sig_name == 'table_C_shares':
            if '발행주식총수' in row_text and nums:
                result['issued_shares_total'] = nums[0]
            if '공모' in row_text and '희망' not in row_text and nums:
                result['public_offering_shares'] = nums[0]
            if '의무인수' in row_text and nums:
                result['mandatory_subscription'] = nums[0]
            if '합계' in row_text and nums:
                result['applied_shares_total'] = nums[0]

        elif sig_name == 'table_D_offering_band':
            if '주당 평가가액' in row_text and '적용' not in row_text and nums:
                result['intrinsic_value_per_share'] = nums[0]
            if '할인율' in row_text:
                # 행 번호·기타 잡값 제거: IPO 할인율은 현실적으로 5~100% 범위
                pct_nums = [n for n in nums if 5.0 <= n <= 100.0]
                if len(pct_nums) >= 2:
                    result['discount_rate_low'], result['discount_rate_high'] = pct_nums[0], pct_nums[1]
                elif len(pct_nums) == 1:
                    result['discount_rate_low'] = pct_nums[0]
            if '밴드' in row_text and len(nums) >= 2:
                result['offering_band_low'], result['offering_band_high'] = nums[0], nums[1]
            if '확정 주당' in row_text and nums:
                if not re.search(r'미정|TBD|수요예측|확정\s*예정|결정\s*예정', row_text):
                    # 값 범위 검증 : 100원 이상만 유효
                    valid = [n for n in nums if n >=100]
                    if valid :
                        result['final_offering_price'] = nums[0]

    return result


def normalize_extracted(extracted: List[Dict]) -> Dict:
    """다양한 추출 결과를 표준 키 구조로 정규화"""
    result = {}
    for item in extracted:
        if item['method'] == 'gemini_vision':
            result.update(item['extracted'])
        else:
            sig_name = item.get('signature')
            if sig_name:
                parsed = _parse_table_by_signature(item['table'], sig_name)
                result.update(parsed)
    return result

def validate_against_ground_truth(extracted: Dict, gt: Dict, tolerance: float = 0.05) -> Dict:
    """한라캐스트 ground truth와 비교하여 정확도 측정"""
    results = {'match': [], 'mismatch': [], 'missing': []}
    
    flat_gt = {}
    for sig, fields in gt.items():
        flat_gt.update(fields)
    
    for key, gt_val in flat_gt.items():
        if gt_val is None:
            continue
        extracted_val = extracted.get(key)
        if extracted_val is None:
            results['missing'].append(key)
        else:
            try:
                ex_v = float(extracted_val)
                gt_v = float(gt_val)
                if abs(ex_v - gt_v) / max(abs(gt_v), 1) <= tolerance:
                    results['match'].append({'key': key, 'extracted': ex_v, 'gt': gt_v})
                else:
                    results['mismatch'].append({'key': key, 'extracted': ex_v, 'gt': gt_v})
            except (ValueError, TypeError):
                results['mismatch'].append({'key': key, 'extracted': extracted_val, 'gt': gt_val, 'note': 'type_error'})
    
    return results

# ============ 메인 ============
def extract_pdf(pdf_path: str, fallback_to_gemini: bool = True) -> Dict:
    """3단계 추출 전략 적용"""
    print(f"\n추출 시작: {pdf_path}")
    
    # 1차: pdfplumber
    print("  1차 시도: pdfplumber")
    results = extract_with_pdfplumber(pdf_path)
    if results:
        print(f"    pdfplumber 결과: {len(results)}개 표 추출")
    
    # 2차: Camelot (1차 결과 부족 시)
    if len(results) < 4:
        print("  2차 시도: Camelot")
        camelot_results = extract_with_camelot(pdf_path)
        results.extend(camelot_results)
    
    # 3차: Gemini Vision (여전히 부족하면)
    if fallback_to_gemini and len(results) < 4:
        print("  3차 시도: Gemini Vision (비용 발생)")
        gemini_results = extract_with_gemini(pdf_path)
        results.extend(gemini_results)
    
    normalized = normalize_extracted(results)
    return {
        'pdf_path': pdf_path,
        'extracted_raw': results,
        'normalized': normalized,
    }

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--pdf', help='단일 PDF 파일 경로')
    parser.add_argument('--rcpno', help='DART rcpNo (PDF 자동 다운로드)')
    parser.add_argument('--validate', action='store_true', help='한라캐스트 ground truth로 검증')
    args = parser.parse_args()
    
    if args.pdf:
        result = extract_pdf(args.pdf)
        print(json.dumps(result['normalized'], ensure_ascii=False, indent=2))
        
        if args.validate:
            validation = validate_against_ground_truth(result['normalized'], HALLA_CAST_GT)
            print(f"\n검증 결과: 일치 {len(validation['match'])} / 불일치 {len(validation['mismatch'])} / 누락 {len(validation['missing'])}")
    else:
        print("사용법: python extract_pdf_tables.py --pdf <path>")
        print("       python extract_pdf_tables.py --rcpno <DART rcpNo>")
