import os
import glob
import json
import time
import matplotlib.pyplot as plt
import numpy as np
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# 커스텀 모듈 임포트
from parser import extract_text_from_pdf
from processor import refine_to_json 

# 한글 폰트 설정 (Mac/Windows 호환성 고려)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

def set_cell_background(cell, fill_color):
    """셀의 배경색을 설정하는 함수"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), fill_color)
    cell._element.get_or_add_tcPr().append(shading_elm)

def apply_center_alignment(target):
    """문단이나 표 셀의 텍스트를 중앙 정렬하는 함수"""
    if hasattr(target, 'paragraphs'): # 셀(Cell)인 경우
        for paragraph in target.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else: # 문단(Paragraph)인 경우
        target.alignment = WD_ALIGN_PARAGRAPH.CENTER

def apply_table_style(table, apply_header_color=True):
    """표 스타일 적용: 중앙 정렬, 헤더 Bold 및 배경색"""
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            apply_center_alignment(cell)
            if cell.paragraphs:
                p = cell.paragraphs[0]
                run = p.runs[0] if p.runs else p.add_run(cell.text)
                if i == 0 or j == 0:
                    run.bold = True
                if i == 0 and apply_header_color:
                    set_cell_background(cell, "FFFFCC")

def create_market_chart(market_data, file_name, target_dir):
    """타겟 시장 규모 비교 막대 그래프 생성"""
    if not market_data or len(market_data) < 2:
        return None
    try:
        labels = [str(row[0]) for row in market_data[1:]]
        # 수치 변환 (콤마 제거 등)
        values = []
        for row in market_data[1:]:
            try:
                val_str = str(row[1]).replace(',', '').strip()
                values.append(float(val_str))
            except:
                values.append(0.0)

        plt.figure(figsize=(8, 5))
        bars = plt.bar(labels, values, color='#4F81BD')
        plt.title(f'[{file_name}] 타겟 시장 규모 비교', fontsize=13, fontweight='bold', pad=15)
        
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, yval, f'{yval:,.0f}', ha='center', va='bottom', fontweight='bold')
        
        plt.tight_layout()
        chart_path = os.path.join(target_dir, f"{file_name}_market_chart.png")
        plt.savefig(chart_path, dpi=120)
        plt.close()
        return chart_path
    except Exception as e:
        print(f"   [Chart Error] 시장 차트 생성 실패: {e}")
        return None

def create_financial_chart(financial_table, file_name, target_dir):
    """매출액(Bar) + 영업이익(Line) 이중축 그래프 생성"""
    try:
        if not financial_table or len(financial_table) < 2: return None
        
        # 헤더에서 연도 추출 (YYYY 형태만)
        years = [str(y) for y in financial_table[0][1:]]
        
        rev_row = next((row for row in financial_table if "매출" in str(row[0])), None)
        opp_row = next((row for row in financial_table if "영업이익" in str(row[0])), None)
        
        if not rev_row: return None

        def to_float(val):
            v = str(val).replace(',', '').replace('(', '-').replace(')', '').replace('억', '').strip()
            try: return float(v)
            except: return 0.0

        rev_values = [to_float(v) for v in rev_row[1:]]
        
        fig, ax1 = plt.subplots(figsize=(10, 6))
        bars = ax1.bar(years, rev_values, color='#34495E', alpha=0.8, width=0.5, label='매출액')
        ax1.set_ylabel('매출액 (억 원)', fontsize=11, fontweight='bold')
        
        # 막대 위 값 표시
        max_val = max(rev_values) if rev_values else 100
        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + (max_val*0.02), f'{int(height):,}', ha='center', va='bottom', fontweight='bold')

        if opp_row:
            opp_values = [to_float(v) for v in opp_row[1:]]
            # 데이터 개수가 안 맞을 경우 보정
            if len(opp_values) != len(years):
                 opp_values = opp_values[:len(years)]
            
            ax2 = ax1.twinx()
            ax2.plot(years, opp_values, color='#E74C3C', marker='o', linewidth=2, label='영업이익')
            ax2.set_ylabel('영업이익 (억 원)', fontsize=11, fontweight='bold', color='#E74C3C')
            ax2.tick_params(axis='y', labelcolor='#E74C3C')
            
            lines, labels = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax1.legend(lines + lines2, labels + labels2, loc='upper left')
        else:
            ax1.legend(loc='upper left')

        plt.title(f'[{file_name}] 연도별 실적 추이 및 전망', fontsize=15, pad=20, fontweight='bold')
        chart_path = os.path.join(target_dir, f"{file_name}_chart.png")
        plt.tight_layout()
        plt.savefig(chart_path, dpi=150)
        plt.close()
        return chart_path
    except Exception as e:
        print(f"   [Chart Error] 재무 차트 생성 실패: {e}")
        return None

def save_as_word_report(data, file_name, target_dir):
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    doc = Document()
    header_info = data.get("Report_Header", {})
    company_name = header_info.get("Company_Name", file_name)
    
    # [표지]
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(f"\n{company_name}\n투자 대상기업 검토보고서")
    run.font.size = Pt(24); run.bold = True
    
    date_p = doc.add_paragraph(f"\n작성일: 2025년 00월 00일")
    date_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_page_break()

    # 1. Executive Summary
    h1 = doc.add_heading('Executive Summary', level=1)
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(data.get('Investment_Thesis_Summary', '내용 없음'))

    info_table = doc.add_table(rows=3, cols=2)
    info_table.style = 'Table Grid'
    rows_data = [
        ["기업명 / 대표자", f"{company_name} / {data.get('CEO_Name', '확인 필요')}"],
        ["담당 심사역", "LUCEN Investment Intelligence"],
        ["투자 등급(종합)", header_info.get("Investment_Rating", "N/A")]
    ]
    for i, row in enumerate(rows_data):
        info_table.rows[i].cells[0].text = row[0]
        info_table.rows[i].cells[1].text = str(row[1])
    apply_table_style(info_table, apply_header_color=False)
    doc.add_paragraph()

    # 2. 시장 동향 및 기술 분석
    h2 = doc.add_heading('시장 동향 및 기술 분석', level=1)
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('1. 최신 시장 동향 및 Pain Points', level=2)
    for point in data.get('Problem_and_Solution', {}).get('Market_Pain_Points', []):
        doc.add_paragraph(f"○ {str(point)}")

    doc.add_heading('2. 핵심 기술 및 독점적 해자 (Moat)', level=2)
    tech = data.get('Technology_and_Moat', {})
    p = doc.add_paragraph(f"핵심 기술명: {tech.get('Core_Technology_Name', 'N/A')}")
    p.runs[0].bold = True
    
    for detail in tech.get('Technical_Details', []):
        doc.add_paragraph(f"○ {str(detail)}")

    # 3. 재무 현황 및 성장 전망
    h3 = doc.add_heading('재무 현황 및 성장 전망', level=1)
    h3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('1. 타겟 시장 규모 및 성장성', level=2)
    market_opp = data.get('Market_Opportunity', {})
    
    # [시장 차트]
    m_chart_path = create_market_chart(market_opp.get('Market_Chart_Data'), file_name, target_dir)
    if m_chart_path and os.path.exists(m_chart_path):
        doc.add_picture(m_chart_path, width=Inches(5.0))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap = doc.add_paragraph(f"[그림 1] {company_name} 타겟 시장 규모 비교")
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(str(market_opp.get('TAM_SAM_SOM_Text', '')))

    # [재무 표 및 차트]
    financial_data = data.get('Table_Data_Preview', {}).get('Financial_Table', [])
    if financial_data:
        doc.add_heading('2. 주요 재무 실적 및 추정치', level=2)
        fin_table = doc.add_table(rows=len(financial_data), cols=len(financial_data[0]))
        fin_table.style = 'Table Grid'
        for i, row in enumerate(financial_data):
            for j, val in enumerate(row):
                fin_table.rows[i].cells[j].text = str(val)
        apply_table_style(fin_table)
        
        f_chart_path = create_financial_chart(financial_data, file_name, target_dir)
        if f_chart_path and os.path.exists(f_chart_path):
            doc.add_picture(f_chart_path, width=Inches(5.0))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap = doc.add_paragraph(f"[그림 2] 연도별 실적 추이")
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 4. 리스크 및 주요 질의
    h4 = doc.add_heading('리스크 분석 및 향후 과제', level=1)
    h4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('1. 주요 리스크 및 대응 전략', level=2)
    risks = data.get('Key_Risks_and_Mitigation', [])
    if risks:
        risk_table = doc.add_table(rows=1, cols=2)
        risk_table.style = 'Table Grid'
        risk_table.rows[0].cells[0].text = '리스크 요인'
        risk_table.rows[0].cells[1].text = '대응 전략'
        apply_table_style(risk_table) # 헤더 스타일 적용
        
        for r in risks:
            if isinstance(r, dict):
                row = risk_table.add_row().cells
                row[0].text = str(r.get('Risk_Factor', ''))
                row[1].text = str(r.get('Mitigation_Strategy', ''))
                # 내용 행은 중앙 정렬 강제하지 않음 (가독성)

    doc.add_heading('2. 주요 질의 및 사후 확인 사항', level=2)
    for q in data.get('Due_Diligence_Questions', []):
        doc.add_paragraph(f"○ {str(q)}")

    save_name = f"{file_name}_검토보고서.docx"
    report_path = os.path.join(target_dir, save_name)
    doc.save(report_path)
    return report_path

def main():
    input_dir = "data"
    output_dir = "output"
    report_dir = "output_report"

    for d in [output_dir, report_dir]:
        if not os.path.exists(d): os.makedirs(d)

    pdf_files = glob.glob(os.path.join(input_dir, "*.pdf"))
    if not pdf_files:
        print(f"[{input_dir}] 폴더에 PDF 파일이 없습니다.")
        return

    print(f"총 {len(pdf_files)}개의 파일을 처리를 시작합니다.")

    for i, file_path in enumerate(pdf_files):
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        json_path = os.path.join(output_dir, f"{file_name}_refined.json")

        print(f"\n>>> [{i+1}/{len(pdf_files)}] 처리 중: {file_name}")

        try:
            # 1. 기존 분석 결과 확인 (Skip 로직)
            if os.path.exists(json_path):
                with open(json_path, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                        if isinstance(data, dict) and "error" not in data:
                            print(f"   - [Skip] 기존 JSON 활용 -> Word 보고서 생성")
                            save_as_word_report(data, file_name, report_dir)
                            continue # 다음 파일로
                    except:
                        pass # JSON 깨졌으면 다시 분석
            
            # 2. 텍스트 추출
            raw_text = extract_text_from_pdf(file_path)
            if not raw_text:
                print("   - [Error] 텍스트 추출 실패")
                continue

            # 3. AI 분석 (4단계 프로세스 호출)
            # processor.py에서 텍스트 길이 조절하므로 여기서는 원본 전달
            refined_data = refine_to_json(raw_text)
            
            # 4. JSON 저장
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(refined_data, f, ensure_ascii=False, indent=2)

            # 5. 워드 보고서 생성
            if isinstance(refined_data, dict) and "error" not in refined_data:
                print(f"   - 보고서(Word) 생성 중...")
                report_path = save_as_word_report(refined_data, file_name, report_dir)
                print(f"   - [성공] {report_path}")
            else:
                err_msg = refined_data.get('message', 'Unknown Error')
                print(f"   - [실패] API 응답 에러: {err_msg[:100]}...")

        except Exception as e:
            print(f"   - [예외 발생] {str(e)}")
        
        # [Rate Limit 방어] 파일 간 충분한 휴식 (15초)
        print("   ⏳ API 한도 관리를 위해 15초 대기합니다...")
        time.sleep(15)

if __name__ == "__main__":
    main()