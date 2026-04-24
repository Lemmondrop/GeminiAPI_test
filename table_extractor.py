import os
import pandas as pd
import pdfplumber

class TableExtractor:
    """IR 문서 및 재무제표 파일에서 표 데이터를 추출하여 Markdown으로 변환하는 클래스"""

    @staticmethod
    def extract_pdf_tables_to_md(pdf_path: str) -> str:
        """
        1. IR 자료(PDF)에서 표를 추출하여 Markdown으로 변환합니다.
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {pdf_path}")

        md_tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                
                for table_idx, table in enumerate(tables):
                    if not table:
                        continue
                        
                    # pdfplumber는 빈 셀을 None으로 반환하므로 빈 문자열로 정제
                    # 줄바꿈(\n)이 포함된 텍스트는 띄어쓰기로 치환하여 Markdown 표가 깨지지 않게 방지
                    cleaned_table = [
                        [str(cell).replace('\n', ' ').strip() if cell is not None else "" for cell in row]
                        for row in table
                    ]
                    
                    # 헤더(첫 번째 행)와 데이터(나머지 행) 분리
                    if len(cleaned_table) > 1:
                        headers = cleaned_table[0]
                        # 헤더에 중복된 이름이나 빈 값이 있으면 pandas에서 에러가 날 수 있으므로 임의 처리
                        headers = [h if h else f"Column_{i}" for i, h in enumerate(headers)]
                        
                        try:
                            df = pd.DataFrame(cleaned_table[1:], columns=headers)
                            md_table = df.to_markdown(index=False)
                            md_tables.append(f"**[Page {page_num + 1} - Table {table_idx + 1}]**\n{md_table}")
                        except Exception as e:
                            print(f"Page {page_num + 1} 표 변환 오류: {e}")
                            
        return "\n\n".join(md_tables)

    @staticmethod
    def extract_financial_data_to_md(file_path: str) -> str:
        """
        2. 재무제표 원본(Excel, CSV)을 읽어 Markdown으로 변환합니다.
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

        ext = os.path.splitext(file_path)[1].lower()
        
        try:
            if ext == '.csv':
                df = pd.read_csv(file_path)
            elif ext in ['.xls', '.xlsx']:
                df = pd.read_excel(file_path)
            else:
                raise ValueError(f"지원하지 않는 파일 형식입니다: {ext}")
            
            # 결측치(NaN)를 빈 문자열로 처리
            df = df.fillna("")
            
            # 컬럼명 중 줄바꿈 제거 (Markdown 형식 유지)
            df.columns = [str(col).replace('\n', ' ') for col in df.columns]
            
            return df.to_markdown(index=False)
            
        except Exception as e:
            return f"파일 읽기/변환 중 오류 발생: {e}"