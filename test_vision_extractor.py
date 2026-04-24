import os
import time  # API 호출 간격을 위한 time 모듈 추가
import fitz  # PyMuPDF
import google.generativeai as genai
from PIL import Image
from dotenv import load_dotenv

# 환경변수(API KEY) 로드
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
if not API_KEY:
    raise ValueError("API_KEY가 설정되지 않았습니다. .env 파일을 확인해 주세요.")

genai.configure(api_key=API_KEY)

def pdf_page_to_image(doc, page_num: int, dpi: int = 200) -> Image.Image:
    """
    PDF의 특정 페이지를 고해상도 이미지(PIL Image)로 변환합니다.
    """
    page = doc.load_page(page_num)
    
    # 해상도 설정 (dpi 200~300 권장)
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img

def extract_markdown_from_image(image: Image.Image) -> str:
    """
    Gemini Vision API를 사용하여 이미지 내의 표와 텍스트를 Markdown으로 추출합니다.
    """
    # 텍스트와 레이아웃 이해도가 높은 gemini-2.0-flash 모델 적용
    model = genai.GenerativeModel('gemini-2.0-flash')
    
    prompt = """
    당신은 전문적인 데이터 추출 및 문서화 AI입니다.
    주어진 이미지는 스타트업의 IR(투자유치) 자료 슬라이드입니다. 이미지에 포함된 모든 텍스트, 표, 구조도를 완벽한 형태의 Markdown 문서로 변환해 주세요.

    [중요 지침]
    1. **표(Table)**: 데이터가 행과 열로 구성된 표 형태라면, 반드시 Markdown 표 형식(`| Column | Column |`)으로 깔끔하게 정리하십시오.
    2. **다이어그램/흐름도**: 시각적인 도형이나 화살표로 이어진 프로세스, 파이프라인 등은 계층적인 글머리기호(`-`, `*`, `1.`)를 사용하여 문맥이 매끄럽게 이어지도록 구조화하여 요약하십시오.
    3. **불필요한 파편화 방지**: 무의미한 띄어쓰기나 단어 쪼개짐을 수정하고, 문장과 단어를 자연스럽게 이어 붙이십시오.
    4. **디자인 요소 무시**: 의미 없는 배경, 단순 아이콘 디자인은 무시하고 '핵심 정보와 텍스트'의 논리적 배치에 집중하십시오.
    """
    
    try:
        response = model.generate_content([prompt, image])
        return response.text
    except Exception as e:
        return f"Error generation: {str(e)}"

if __name__ == "__main__":
    print("=== Vision LLM 기반 PDF 전체 페이지 -> Markdown 일괄 추출 ===")
    
    # data 폴더의 절대 경로
    data_dir = r"C:\Users\Researcher\Desktop\Project V\OCR Sample\data"
    
    # 🚨 탐색에서 제외할 폴더명 리스트 설정
    EXCLUDE_FOLDERS = ["metadata", "test", "밸류추정 인수인의 의견"]
    
    if not os.path.exists(data_dir):
        print(f"❌ 데이터 폴더를 찾을 수 없습니다: {data_dir}")
    else:
        # 1. data 폴더 내의 모든 기업 폴더들을 순회
        for company_folder in os.listdir(data_dir):
            company_path = os.path.join(data_dir, company_folder)
            
            # 폴더가 아니면 스킵
            if not os.path.isdir(company_path):
                continue
                
            # 🚨 제외 리스트에 포함된 폴더면 스킵
            if company_folder in EXCLUDE_FOLDERS:
                print(f"⏭️ [Skip] 제외 설정된 폴더입니다: {company_folder}")
                continue
                
            print(f"\n📂 [{company_folder}] 폴더 탐색 중...")
            
            # 2. 해당 기업 폴더 내의 PDF 파일 찾기
            for file_name in os.listdir(company_path):
                if file_name.lower().endswith(".pdf"):
                    target_pdf = os.path.join(company_path, file_name)
                    
                    # 3. 생성할 Markdown 파일 이름 및 경로 설정 (PDF파일명_전체추출결과.md)
                    md_filename = f"{os.path.splitext(file_name)[0]}_전체추출결과.md"
                    output_md = os.path.join(company_path, md_filename)
                    
                    # 4. 이미 추출된 파일이 있다면 스킵 (API 호출 및 시간 절약)
                    if os.path.exists(output_md):
                        print(f"  ⏭️ 이미 추출된 파일이 존재하여 스킵합니다: {md_filename}")
                        continue
                        
                    print(f"  📄 PDF 발견: {file_name}")
                    
                    try:
                        doc = fitz.open(target_pdf)
                        total_pages = doc.page_count
                        print(f"  ▶ 총 {total_pages}페이지 추출 시작...")
                        
                        # 파일 열기 (덮어쓰기 모드)
                        with open(output_md, "w", encoding="utf-8") as f:
                            f.write(f"# {file_name} 전체 추출 결과\n\n")
                            
                            for i in range(total_pages):
                                print(f"    🔄 [{i+1}/{total_pages}] 페이지 처리 중...")
                                
                                page_image = pdf_page_to_image(doc, i, dpi=200)
                                md_text = extract_markdown_from_image(page_image)
                                
                                f.write(f"## Page {i+1}\n\n")
                                f.write(md_text + "\n\n")
                                f.write("---\n\n")
                                
                                # API 속도 제한(Rate Limit) 방지를 위해 3초 대기 (마지막 페이지 제외)
                                if i < total_pages - 1:
                                    time.sleep(3)
                        
                        doc.close()
                        print(f"  ✅ 추출 완료! 저장 위치: {output_md}")
                        
                    except Exception as e:
                        print(f"  ❌ 파일 처리 중 오류 발생 ({file_name}): {e}")

    print("\n🎉 모든 기업 폴더의 PDF 추출 작업이 완료되었습니다!")