import os
import glob
import json
import matplotlib.pyplot as plt
import numpy as np
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
# processor.py에서 함수와 스키마 임포트
from processor import refine_pdf_to_json_onecall, JSON_SCHEMA

# =========================================================
# ✅ [설정] 디자인 및 색상 테마
# =========================================================
plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False

COLOR_RED = "indianred"
COLOR_YELLOW = "khaki"
COLOR_BLUE = "cornflowerblue"

# =========================================================
# 1. 문서 스타일링 헬퍼 함수
# =========================================================
def set_cell_background(cell, fill_color):
    """셀 배경색 설정"""
    shading_elm = OxmlElement("w:shd")
    shading_elm.set(qn("w:fill"), fill_color)
    cell._element.get_or_add_tcPr().append(shading_elm)

def apply_table_style(table, header_color="E7E6E6", center_all=True):
    """표 스타일: 헤더 강조, 정렬"""
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            if center_all:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i == 0: # 헤더 행
                set_cell_background(cell, header_color)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

# =========================================================
# 2. 그래프 생성 함수
# =========================================================
def clean_numeric(val):
    """문자열(쉼표, 단위 등)을 실수형(float)으로 안전하게 변환"""
    if isinstance(val, (int, float)):
        return float(val)
    if val is None:
        return 0.0
    
    s = str(val).replace(",", "").replace("억", "").replace("원", "").strip()
    if s in ["", "N/A", "-", "n/a", "null"]:
        return 0.0
        
    try:
        if s.startswith("(") and s.endswith(")"): # (500) -> -500
            s = "-" + s[1:-1]
        return float(s)
    except:
        return 0.0

def create_basic_bar_chart(data_list, title, color, file_name, suffix, target_dir):
    """단일 데이터 계열 막대 그래프"""
    if not data_list or len(data_list) < 2:
        return None

    try:
        years = []
        values = []
        for row in data_list[1:]:
            if isinstance(row, list) and len(row) >= 2:
                y = str(row[0]).replace("년","").strip()
                v = clean_numeric(row[1])
                years.append(y)
                values.append(v)
        
        if not years: return None

        plt.figure(figsize=(6, 4))
        
        if len(years) == 1:
            plt.bar(years, values, color=color, width=0.3)
            plt.xlim(-1, 1)
        else:
            plt.bar(years, values, color=color, width=0.5)

        plt.title(title, fontsize=13, fontweight="bold", pad=15)
        plt.grid(axis="y", linestyle="--", alpha=0.5)
        
        max_val = max(values) if values else 1
        for bar in plt.gca().patches: # patches로 접근
            height = bar.get_height()
            if height != 0:
                plt.text(bar.get_x() + bar.get_width()/2.0, height + (max_val*0.01),
                        f"{int(height):,}", ha='center', va='bottom', fontsize=9, fontweight='bold')

        save_path = os.path.join(target_dir, f"{file_name}_{suffix}.png")
        plt.tight_layout()
        plt.savefig(save_path, dpi=150)
        plt.close()
        return save_path
    except Exception as e:
        print(f"  [Graph Error] {title}: {e}")
        return None

def create_grouped_bar_chart_3(years, data1, data2, data3, labels, colors, title, file_name, suffix, target_dir):
    """3개 데이터 계열 그룹 막대 그래프 (길이 보정 포함)"""
    try:
        if not years or len(years) == 0: return None
        
        target_len = len(years)
        
        def pad_or_truncate(d_list):
            if d_list is None: d_list = []
            if len(d_list) < target_len:
                return d_list + [0] * (target_len - len(d_list))
            return d_list[:target_len]

        d1 = [clean_numeric(v) for v in pad_or_truncate(data1)]
        d2 = [clean_numeric(v) for v in pad_or_truncate(data2)]
        d3 = [clean_numeric(v) for v in pad_or_truncate(data3)]

        x = np.arange(len(years))
        width = 0.25

        plt.figure(figsize=(8, 5))
        plt.bar(x - width, d1, width, label=labels[0], color=colors[0])
        plt.bar(x,         d2, width, label=labels[1], color=colors[1])
        plt.bar(x + width, d3, width, label=labels[2], color=colors[2])

        plt.title(title, fontsize=14, fontweight="bold", pad=15)
        plt.xticks(x, years)
        plt.legend()
        plt.grid(axis="y", linestyle="--", alpha=0.3)
        plt.tight_layout()

        save_path = os.path.join(target_dir, f"{file_name}_{suffix}.png")
        plt.savefig(save_path, dpi=150)
        plt.close()
        return save_path
    except Exception as e:
        print(f"  [Graph Error] {title}: {e}")
        return None

# =========================================================
# 3. Word 보고서 작성 (✅ Null Safety 강화)
# =========================================================
def save_as_word_report(data, file_name, target_dir):
    if not data: return None

    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    doc = Document()
    
    # ✅ .get("key") or {} 패턴 사용 (JSON null -> None 방지)
    header = data.get("Report_Header") or {}
    comp_name = header.get("Company_Name", file_name)
    ceo_name = header.get("CEO_Name", "확인 필요")
    rating = header.get("Investment_Rating", "-")

    # [표지]
    doc.add_paragraph("\n\n\n")
    t = doc.add_heading(f"{comp_name}\n투자 검토 보고서", 0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("\n" * 5)
    p = doc.add_paragraph(f"Date: 2025. XX. XX\nAnalyst: LUCEN Investment Intelligence")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_page_break()

    # [Executive Summary]
    doc.add_heading("Executive Summary", level=1)
    table = doc.add_table(rows=3, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "기업명"
    table.rows[0].cells[1].text = str(comp_name)
    table.rows[1].cells[0].text = "대표자"
    table.rows[1].cells[1].text = str(ceo_name)
    table.rows[2].cells[0].text = "투자 등급"
    table.rows[2].cells[1].text = str(rating)
    apply_table_style(table)
    doc.add_paragraph("\n")
    doc.add_paragraph(data.get("Investment_Thesis_Summary") or "")

    # [1. 재무 현황]
    doc.add_heading("1. 재무 현황 및 투자 유치 현황", level=1)
    
    # ✅ Null Safety
    fin = data.get("Financial_Status") or {}

    # 1-1. 재무상태표
    doc.add_heading("1-1. 재무상태표 (Balance Sheet)", level=2)
    bs = fin.get("Detailed_Balance_Sheet") or {}
    years = bs.get("Years") or []
    
    row_keys = [
        ("유동자산", "Current_Assets"), ("비유동자산", "Non_Current_Assets"), ("자산총계", "Total_Assets"),
        ("유동부채", "Current_Liabilities"), ("비유동부채", "Non_Current_Liabilities"), ("부채총계", "Total_Liabilities"),
        ("자본금", "Capital_Stock"), ("이익잉여금외", "Retained_Earnings_Etc"), ("자본총계", "Total_Equity")
    ]

    if years:
        t_bs = doc.add_table(rows=len(row_keys)+1, cols=len(years)+1)
        t_bs.style = "Table Grid"
        t_bs.rows[0].cells[0].text = "구분"
        for i, y in enumerate(years):
            t_bs.rows[0].cells[i+1].text = str(y)
        
        for r_idx, (label, key) in enumerate(row_keys):
            vals = bs.get(key) or []
            t_bs.rows[r_idx+1].cells[0].text = label
            for c_idx, val in enumerate(vals):
                if c_idx < len(years):
                    t_bs.rows[r_idx+1].cells[c_idx+1].text = str(val)
        apply_table_style(t_bs, header_color="FFFFCC")

        # 재무상태표 그래프
        t_assets = bs.get("Total_Assets") or []
        t_liab = bs.get("Total_Liabilities") or []
        t_equity = bs.get("Total_Equity") or []

        if t_assets:
            doc.add_paragraph("\n■ 재무 상태 추이 (자산/부채/자본)")
            bs_img = create_grouped_bar_chart_3(
                years, t_assets, t_liab, t_equity,
                ["자산총계", "부채총계", "자본총계"],
                [COLOR_BLUE, COLOR_YELLOW, COLOR_RED],
                "재무상태 변동 추이", file_name, "bs_chart", target_dir
            )
            if bs_img: doc.add_picture(bs_img, width=Inches(6.0))
    else:
        doc.add_paragraph("재무상태표 데이터 없음")

    # 1-2. 손익계산서
    doc.add_heading("1-2. 수익성 분석 (Income Statement)", level=2)
    doc.add_paragraph(fin.get("Key_Financial_Commentary") or "")

    is_data = fin.get("Income_Statement_Summary") or {}
    is_years = is_data.get("Years") or []
    rev = is_data.get("Total_Revenue") or []
    op_prof = is_data.get("Operating_Profit") or []
    net_prof = is_data.get("Net_Profit") or []

    if is_years:
        doc.add_paragraph("\n■ 수익성 추이 (매출/영업이익/순이익)")
        is_img = create_grouped_bar_chart_3(
            is_years, rev, op_prof, net_prof,
            ["매출액", "영업이익", "당기순이익"],
            [COLOR_BLUE, COLOR_YELLOW, COLOR_RED],
            "수익성 변동 추이", file_name, "is_chart", target_dir
        )
        if is_img: doc.add_picture(is_img, width=Inches(6.0))

    # 1-3. 투자 유치 현황
    doc.add_heading("1-3. 투자 유치 현황", level=2)
    inv_hist = fin.get("Investment_History") or []
    
    if inv_hist:
        t_inv = doc.add_table(rows=1, cols=4)
        t_inv.style = "Table Grid"
        t_inv.rows[0].cells[0].text = "Date"
        t_inv.rows[0].cells[1].text = "Round"
        t_inv.rows[0].cells[2].text = "Amount"
        t_inv.rows[0].cells[3].text = "Investor"
        apply_table_style(t_inv)
        for item in inv_hist:
            row = t_inv.add_row().cells
            row[0].text = str(item.get("Date") or "")
            row[1].text = str(item.get("Round") or "")
            row[2].text = str(item.get("Amount") or "")
            row[3].text = str(item.get("Investor") or "")
    else:
        doc.add_paragraph("투자 이력 없음")

    # [2. 성장 가능성]
    doc.add_heading("2. 성장 가능성", level=1)
    # ✅ Null Safety
    growth = data.get("Growth_Potential") or {}

    doc.add_heading("2-1. 시장 동향 및 레퍼런스", level=2)
    trends = growth.get("Target_Market_Trends") or []
    for t_item in trends:
        p = doc.add_paragraph()
        p.add_run(f"[{t_item.get('Type','Info')}] ").bold = True
        p.add_run(f"{t_item.get('Content','')}\n- Source: {t_item.get('Source','')}")

    doc.add_heading("2-2. 주요 성장 지표", level=2)
    stats = growth.get("Export_and_Contract_Stats") or {}

    if stats.get("Export_Graph_Data"):
        doc.add_paragraph("■ 수출 실적")
        img = create_basic_bar_chart(stats["Export_Graph_Data"], "수출 실적 추이", COLOR_RED, file_name, "export", target_dir)
        if img: doc.add_picture(img, width=Inches(5.5))

    if stats.get("Contract_Count_Graph_Data"):
        doc.add_paragraph("\n■ 계약 건수")
        img = create_basic_bar_chart(stats["Contract_Count_Graph_Data"], "계약 체결 추이", COLOR_YELLOW, file_name, "contract", target_dir)
        if img: doc.add_picture(img, width=Inches(5.5))
        
    if stats.get("Sales_Graph_Data"):
        doc.add_paragraph("\n■ 매출 규모(성장성 관점)")
        img = create_basic_bar_chart(stats["Sales_Graph_Data"], "매출 성장 추이", COLOR_BLUE, file_name, "sales_growth", target_dir)
        if img: doc.add_picture(img, width=Inches(5.5))

    # [3. 기술 및 기타]
    doc.add_heading("3. 기술 경쟁력 및 해결 과제", level=1)
    
    # ✅ Null Safety
    prob_sol = data.get("Problem_and_Solution") or {}
    doc.add_heading("3-1. Market Pain Points", level=2)
    for mp in prob_sol.get("Market_Pain_Points") or []:
        doc.add_paragraph(f"• {mp}")

    doc.add_heading("3-2. Solution & Core Tech", level=2)
    tech = data.get("Technology_and_Moat") or {}
    doc.add_paragraph(f"핵심기술: {tech.get('Core_Technology_Name') or ''}")
    for td in tech.get("Technical_Details") or []:
        doc.add_paragraph(f"• {td}")

    # [4. 리스크 및 결론]
    doc.add_heading("4. 리스크 및 종합 의견", level=1)
    doc.add_heading("4-1. 주요 리스크", level=2)
    risks = data.get("Key_Risks_and_Mitigation") or []
    if risks:
        rt = doc.add_table(rows=1, cols=2)
        rt.style = "Table Grid"
        rt.rows[0].cells[0].text = "Risk Factor"
        rt.rows[0].cells[1].text = "Mitigation"
        apply_table_style(rt)
        for r in risks:
            row = rt.add_row().cells
            row[0].text = str(r.get("Risk_Factor") or "")
            row[1].text = str(r.get("Mitigation_Strategy") or "")

    doc.add_heading("4-2. 종합 결론", level=2)
    doc.add_paragraph(data.get("Final_Conclusion") or "")

    out_path = os.path.join(target_dir, f"{file_name}_검토보고서.docx")
    doc.save(out_path)
    return out_path

# =========================================================
# 4. Main Execution
# =========================================================
def main():
    input_dir = "data"
    output_dir = "output"
    report_dir = "output_report"
    
    for d in [output_dir, report_dir]:
        os.makedirs(d, exist_ok=True)

    pdf_files = glob.glob(os.path.join(input_dir, "*.pdf"))
    if not pdf_files:
        print("PDF 파일이 없습니다.")
        return

    print(f"총 {len(pdf_files)}개 파일 처리 시작...")

    for i, file_path in enumerate(pdf_files):
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        json_path = os.path.join(output_dir, f"{file_name}_refined.json")
        
        print(f"\n>>> [{i+1}/{len(pdf_files)}] {file_name}")

        try:
            # 1. JSON 추출
            data = refine_pdf_to_json_onecall(file_path)
            
            # None 에러 방지 (None 또는 error 키 체크)
            if data is None or (isinstance(data, dict) and data.get("error")):
                err = data.get("error") if isinstance(data, dict) else "Data is None"
                print(f"   [Error] JSON 생성 실패: {err}")
                continue

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            # 2. 보고서 생성
            doc_path = save_as_word_report(data, file_name, report_dir)
            if doc_path:
                print(f"   [완료] {doc_path}")
            else:
                print("   [Skip] 보고서 생성 실패 (데이터 없음)")

        except Exception as e:
            # traceback을 출력하여 정확한 위치 파악
            import traceback
            print(f"   [Fail] {e}")
            traceback.print_exc()

if __name__ == "__main__":
    main()