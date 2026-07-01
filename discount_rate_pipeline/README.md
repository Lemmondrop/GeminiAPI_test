# 할인율 모델 데이터 수집 파이프라인

## 디렉토리 구성

- `할인율모델_데이터수집_명세서_v1.xlsx` (한 단계 위) — 수집 필드 명세 및 250건 진행 시트
- `collect_ipo_data.py` — Phase A+B 자동화 (IPO 목록 + 메타데이터)
- `extract_pdf_tables.py` — Phase C (PDF 표 추출, pdfplumber→Camelot→Gemini Vision)
- `halla_cast_ground_truth.json` — 추출 정확도 검증용 ground truth

## 시작하기

### 1단계: API 키 발급
1. https://opendart.fss.or.kr 회원가입
2. 인증키 신청 (무료, 즉시 발급)
3. 환경변수 설정: `export OPENDART_API_KEY=xxx`

### 2단계: 의존성 설치
```bash
pip install OpenDartReader pdfplumber camelot-py[cv] finance-datareader \
            google-generativeai pdf2image pandas openpyxl
```

### 3단계: Phase A+B 실행 (메타데이터)
```bash
python collect_ipo_data.py --start 2023-01-01 --end 2025-12-31
```

### 4단계: Phase C 검증 (한라캐스트 PDF로 추출 정확도 테스트)
```bash
# 한라캐스트 PDF가 있다면:
python extract_pdf_tables.py --pdf halla_cast_filing.pdf --validate
# → 'applied_net_income: 9910', 'discount_rate_low: 17.54' 등이 나오면 OK
```

### 5단계: 250건 일괄 추출 (검증 통과 후)
```bash
python extract_pdf_tables.py --batch ipo_metadata.csv --output extracted.csv
```

## 권장 작업 순서

1. ✅ 명세서 검토 (Excel 파일)
2. ✅ API 키 발급
3. ⏭️ 한라캐스트 PDF로 추출 정확도 1차 테스트 — **여기가 가장 중요**
4. ⏭️ 정확도 80% 이상 나오면 250건 일괄 시작
5. ⏭️ 정확도 부족하면 추출 로직 보강 (Gemini Vision 비중 ↑)

## 예상 일정

- API 발급 + 환경 설정: 0.5일
- 한라캐스트로 추출 검증: 1일
- 250건 메타데이터 수집: 2일
- 250건 PDF 추출 (자동): 2~3일
- 수동 검수: 2~3일
- **합계: 약 8~10일**

## 핵심 위험 요소

1. **PDF 표 형식 회사별 차이**: 일부 회사는 표가 깨져있거나 이미지로만 존재
   → Gemini Vision fallback으로 해결
2. **정정신고서 처리**: 최초 신고서 vs 정정신고서 가격 차이
   → 가장 최근 신고서 사용 (가격 확정본)
3. **결측치 비율**: 250건 중 100건이 추출 실패하면 모델 학습 부족
   → 사용자 검토 필요한 의문 케이스만 수동 보완

## 의문점이 있으면 알려주세요
