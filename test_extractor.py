import os
import pandas as pd
from table_extractor import TableExtractor

def setup_dummy_excel():
    """테스트를 위한 가상 기업의 재무제표 엑셀 파일을 자동 생성합니다."""
    data = {
        "항목 (단위: 백만원)": ["유동자산", "비유동자산", "자산총계", "유동부채", "부채총계", "자본총계", "매출액", "영업이익"],
        "2023년": [1200, 2500, 3700, 800, 1500, 2200, 4500, 450],
        "2024년": [1800, 2700, 4500, 1100, 1800, 2700, 6200, 800],
        "2025년(예상)": [2500, 3000, 5500, 1300, 2000, 3500, 8500, 1200]
    }
    df = pd.DataFrame(data)
    file_path = "dummy_financials.xlsx"
    df.to_excel(file_path, index=False)
    return file_path

def test_financial_extraction():
    print("=== [TEST 1] 재무제표 엑셀 -> Markdown 변환 및 프롬프트 생성 테스트 ===\n")
    excel_path = setup_dummy_excel()

    try:
        # 1. 엑셀을 마크다운으로 변환
        md_result = TableExtractor.extract_financial_data_to_md(excel_path)

        # 2. LLM에 던질 프롬프트 조립
        prompt = f"""다음은 분석 대상 기업의 재무제표 데이터입니다.

{md_result}

위 데이터를 바탕으로 다음 작업을 수행해 주세요:
1. 유동비율과 부채비율을 계산할 것.
2. 향후 3년간의 성장성과 재무 건전성을 평가할 것.
3. 결과를 JSON 형태로 규격화하여 출력할 것.
"""
        print("[생성된 프롬프트 미리보기]")
        print("-" * 60)
        print(prompt)
        print("-" * 60)
        print("✅ 엑셀 변환 및 프롬프트 조립 테스트 성공!\n")

    finally:
        # 테스트가 끝나면 생성했던 임시 엑셀 파일 삭제 (확인하고 싶다면 아래 두 줄을 주석 처리하세요)
        if os.path.exists(excel_path):
            os.remove(excel_path)

def test_pdf_extraction():
    print("=== [TEST 2] IR 자료 PDF -> Markdown 표 추출 테스트 ===\n")
    
    # 실제 가지고 계신 IR PDF 파일 이름으로 변경해 주세요.
    pdf_path = "C:/Users/Researcher/Desktop/Project V/OCR Sample/data/센스톤(SSenStone)/센스톤(SSenStone) IR Deck_202510_v5.09.pdf" 

    if not os.path.exists(pdf_path):
        print(f"⚠️ '{pdf_path}' 파일을 찾을 수 없어 PDF 테스트는 건너뜁니다.")
        print("   실제 IR 자료 PDF를 같은 폴더에 넣고 파일명을 맞춰주세요.\n")
        return

    # PDF에서 표 추출
    md_result = TableExtractor.extract_pdf_tables_to_md(pdf_path)

    print("[추출된 PDF Markdown 표]")
    print("-" * 60)
    if md_result.strip():
        print(md_result)
    else:
        print("추출된 표가 없습니다. (표출할 수 없는 이미지 형태의 표이거나, 텍스트만 있는 페이지일 수 있습니다.)")
    print("-" * 60)
    print("✅ PDF 표 추출 테스트 완료!\n")

if __name__ == "__main__":
    test_financial_extraction()
    test_pdf_extraction()