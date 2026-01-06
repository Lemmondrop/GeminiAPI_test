import fitz # PyMuPDF

def extract_text_from_pdf(file_path):
    doc = fitz.open(file_path)
    full_content = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text") # 레이아웃 순서대로 텍스트 추출

        # 페이지 구분을 주어 LLM이 문맥을 파악하기 쉽게 함
        page_content = f"\n--- Page {page_num + 1} ---\n{text}"
        full_content.append(page_content)
    
    doc.close()
    return "\n".join(full_content)

if __name__ == "__main__":
    pdf_path = "2025-12-15_보로노이_IR BOOK_f.pdf"

    try:
        raw_text = extract_text_from_pdf(pdf_path)
        # 텍스트가 너무 길면 앞부분만 확인
        print(raw_text[:2000])
    except Exception as e :
        print(f"오류 발생: {e}")