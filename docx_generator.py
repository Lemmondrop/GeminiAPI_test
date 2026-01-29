import os
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# =========================================================
# ✅ [설정] Matplotlib 한글 폰트 등
# =========================================================
plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False

COLOR_RED = "indianred"
COLOR_YELLOW = "khaki"
COLOR_BLUE = "cornflowerblue"

# =========================================================
# 1. 문서 스타일링 헬퍼
# =========================================================
def set_cell_background(cell, fill_color):
    shading_elm = OxmlElement("w:shd")
    shading_elm.set(qn("w:fill"), fill_color)
    cell._element.get_or_add_tcPr().append(shading_elm)

def apply_table_style(table, header_color="E7E6E6", center_all=True):
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            if center_all:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i == 0: 
                set_cell_background(cell, header_color)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

# =========================================================
# 2. 데이터 처리 및 단위 감지
# =========================================================
def clean_numeric(val):
    if isinstance(val, (int, float)): return float(val)
    if val is None: return 0.0
    s = str(val).strip()
    if s.startswith("(") and s.endswith(")"): s = "-" + s[1:-1]
    s = s.replace(",", "").replace("억", "").replace("원", "").replace("개", "").replace("만", "").replace("불", "").replace("$", "")
    if s in ["", "N/A", "-", "n/a", "null"]: return 0.0
    try: return float(s)
    except: return 0.0

def detect_unit_from_data(data_list, default_unit="(단위: 백만원)"):
    if not data_list: return default_unit
    if isinstance(data_list[0], list):
        flat_str = " ".join([str(row[1]) for row in data_list[1:] if len(row) > 1])
    else:
        flat_str = " ".join([str(v) for v in data_list])
    if "억" in flat_str: return "(단위: 억원)"
    if "조" in flat_str: return "(단위: 조원)"
    if "천만" in flat_str: return "(단위: 천만원)"
    if "백만" in flat_str: return "(단위: 백만원)"
    if "$" in flat_str or "달러" in flat_str: return "(단위: USD)"
    return default_unit

# =========================================================
# 3. 그래프 생성 함수
# =========================================================
def create_basic_bar_chart(data_list, title, color, file_name, suffix, target_dir, default_unit_label="(단위: 자료 수치 기준)"):
    if not data_list or len(data_list) < 2: return None
    unit_label = detect_unit_from_data(data_list, default_unit_label)
    try:
        years, values = [], []
        for row in data_list[1:]:
            if isinstance(row, list) and len(row) >= 2:
                y = str(row[0]).replace("년","").strip()
                v = clean_numeric(row[1])
                years.append(y)
                values.append(v)
        if not years: return None

        plt.figure(figsize=(7, 4.5))
        plt.bar(years, values, color=color, width=0.5)
        plt.title(f"{title}\n", fontsize=14, fontweight="bold", pad=10)
        plt.text(0, 1.02, unit_label, transform=plt.gca().transAxes, ha='left', fontsize=10, color='gray')
        plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
        plt.grid(axis="y", linestyle="--", alpha=0.5)
        
        max_val = max(values) if values else 1
        for bar in plt.gca().patches:
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

def create_grouped_bar_chart_3(years, data1, data2, data3, labels, colors, title, file_name, suffix, target_dir, unit_label="(단위: 백만원)"):
    try:
        if not years or len(years) == 0: return None
        target_len = len(years)
        def pad(d):
            if d is None: d = []
            if len(d) < target_len: return d + [0]*(target_len - len(d))
            return d[:target_len]

        d1, d2, d3 = [clean_numeric(v) for v in pad(data1)], [clean_numeric(v) for v in pad(data2)], [clean_numeric(v) for v in pad(data3)]
        x = np.arange(len(years))
        width = 0.25

        plt.figure(figsize=(8.5, 5))
        plt.bar(x - width, d1, width, label=labels[0], color=colors[0])
        plt.bar(x, d2, width, label=labels[1], color=colors[1])
        plt.bar(x + width, d3, width, label=labels[2], color=colors[2])

        plt.title(f"{title}\n", fontsize=15, fontweight="bold", pad=12)
        plt.text(0, 1.02, unit_label, transform=plt.gca().transAxes, ha='left', fontsize=10, color='gray')
        plt.xticks(x, years)
        plt.legend(loc='upper right')
        plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
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
# 4. Word 보고서 작성 함수 (메인 로직)
# =========================================================
def save_as_word_report(data, file_name, target_dir):
    if not data: return None
    if not os.path.exists(target_dir): os.makedirs(target_dir)
    
    doc = Document()
    header = data.get("Report_Header") or {}
    
    # [Cover]
    doc.add_paragraph("\n\n\n")
    t = doc.add_heading(f"{header.get('Company_Name', file_name)}\n투자 검토 보고서", 0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("\n" * 5)
    p = doc.add_paragraph(f"Date: 2025. XX. XX\nAnalyst: {header.get('Analyst', 'LUCEN')}")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_page_break()

    # [Executive Summary]
    doc.add_heading("Executive Summary", level=1)
    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "기업명"
    table.rows[0].cells[1].text = str(header.get("Company_Name", ""))
    table.rows[1].cells[0].text = "대표자"
    table.rows[1].cells[1].text = str(header.get("CEO_Name", ""))
    table.rows[2].cells[0].text = "산업 분야"
    table.rows[2].cells[1].text = str(header.get("Industry_Classification", ""))
    table.rows[3].cells[0].text = "투자 등급"
    table.rows[3].cells[1].text = str(header.get("Investment_Rating", ""))
    apply_table_style(table)
    doc.add_paragraph("\n")
    doc.add_paragraph(data.get("Investment_Thesis_Summary") or "")

    # [Section 1] 재무 현황
    doc.add_heading("1. 재무 현황 및 투자 유치 현황", level=1)
    fin = data.get("Financial_Status") or {}
    
    # 1-1. Balance Sheet
    doc.add_heading("1-1. 재무상태표 (Balance Sheet)", level=2)
    bs = fin.get("Detailed_Balance_Sheet") or {}
    years = bs.get("Years") or []
    if years:
        row_keys = [("자산총계", "Total_Assets"), ("부채총계", "Total_Liabilities"), ("자본총계", "Total_Equity")]
        t_bs = doc.add_table(rows=4, cols=len(years)+1)
        t_bs.style = "Table Grid"
        t_bs.rows[0].cells[0].text = "구분"
        for i, y in enumerate(years): t_bs.rows[0].cells[i+1].text = str(y)
        for r_idx, (label, key) in enumerate(row_keys):
            t_bs.rows[r_idx+1].cells[0].text = label
            vals = bs.get(key) or []
            for c_idx, v in enumerate(vals):
                if c_idx < len(years): t_bs.rows[r_idx+1].cells[c_idx+1].text = str(v)
        apply_table_style(t_bs, header_color="FFFFCC")
        
        chart_path = create_grouped_bar_chart_3(years, bs.get("Total_Assets"), bs.get("Total_Liabilities"), bs.get("Total_Equity"), ["자산", "부채", "자본"], [COLOR_BLUE, COLOR_YELLOW, COLOR_RED], "재무상태 추이", file_name, "bs", target_dir)
        if chart_path: doc.add_picture(chart_path, width=Inches(6.0))

    # 1-2. Income Statement
    doc.add_heading("1-2. 수익성 분석 (Income Statement)", level=2)
    doc.add_paragraph(fin.get("Key_Financial_Commentary") or "")
    is_data = fin.get("Income_Statement_Summary") or {}
    if is_data.get("Years"):
        chart_path = create_grouped_bar_chart_3(is_data.get("Years"), is_data.get("Total_Revenue"), is_data.get("Operating_Profit"), is_data.get("Net_Profit"), ["매출", "영업이익", "순이익"], [COLOR_BLUE, COLOR_YELLOW, COLOR_RED], "수익성 추이", file_name, "is", target_dir)
        if chart_path: doc.add_picture(chart_path, width=Inches(6.0))

    # 1-3. Investment History
    doc.add_heading("1-3. 투자 유치 현황", level=2)
    inv = fin.get("Investment_History") or []
    if inv:
        t_inv = doc.add_table(rows=1, cols=4)
        t_inv.style = "Table Grid"
        t_inv.rows[0].cells[0].text = "Date"
        t_inv.rows[0].cells[1].text = "Round"
        t_inv.rows[0].cells[2].text = "Amount"
        t_inv.rows[0].cells[3].text = "Investor"
        apply_table_style(t_inv)
        for item in inv:
            row = t_inv.add_row().cells
            row[0].text = str(item.get("Date") or "")
            row[1].text = str(item.get("Round") or "")
            row[2].text = str(item.get("Amount") or "")
            row[3].text = str(item.get("Investor") or "")

    # 1-4. Future Revenue Structure
    doc.add_heading("1-4. 미래 수익 구조", level=2)
    fr = fin.get("Future_Revenue_Structure") or {}
    doc.add_paragraph(f"■ 비즈니스 모델: {fr.get('Business_Model', '내용 없음')}")
    doc.add_paragraph(f"■ 향후 Cash Cow: {fr.get('Future_Cash_Cow', '내용 없음')}")

    # [Section 2] 시장성 및 성장 잠재력
    doc.add_heading("2. 시장성 및 성장 잠재력", level=1)
    mg = data.get("Growth_Potential") or {} 
    
    # 2-1. Target Market
    doc.add_heading("2-1. 타겟 시장 분석", level=2)
    tm = mg.get("Target_Market_Analysis") or {}
    doc.add_paragraph(f"• 타겟 영역: {tm.get('Target_Area', '')}")
    doc.add_paragraph(f"• 시장 특성: {tm.get('Market_Characteristics', '')}")
    doc.add_paragraph(f"• 포지셔닝: {tm.get('Competitive_Positioning', '')}")

    # 2-2. Trends
    doc.add_heading("2-2. 시장 동향 및 레퍼런스", level=2)
    for t in mg.get("Target_Market_Trends") or []:
        doc.add_paragraph(f"[{t.get('Type')}] {t.get('Content')} (Source: {t.get('Source')})")

    # 2-3. L/O & Exit
    doc.add_heading("2-3. L/O 및 Exit 전략", level=2)
    lo = mg.get("LO_Exit_Strategy") or {}
    doc.add_paragraph(f"• 검증된 시그널: {', '.join(lo.get('Verified_Signals') or [])}")
    doc.add_paragraph(f"• 적정 가치 범위: {lo.get('Valuation_Range', '확인 필요')}")
    if lo.get("Expected_LO_Scenarios"):
        t_lo = doc.add_table(rows=1, cols=3)
        t_lo.style = "Table Grid"
        t_lo.rows[0].cells[0].text = "구분"
        t_lo.rows[0].cells[1].text = "가능성"
        t_lo.rows[0].cells[2].text = "코멘트"
        apply_table_style(t_lo)
        for s in lo.get("Expected_LO_Scenarios"):
            r = t_lo.add_row().cells
            r[0].text = str(s.get("Category"))
            r[1].text = str(s.get("Probability"))
            r[2].text = str(s.get("Comment"))

    # 2-4. Growth Indicators
    doc.add_heading("2-4. 주요 성장 지표", level=2)
    stats = mg.get("Export_and_Contract_Stats") or {}
    
    if stats.get("Export_Graph_Data"):
        doc.add_paragraph("■ 수출 실적")
        img = create_basic_bar_chart(stats["Export_Graph_Data"], "수출 추이", COLOR_RED, file_name, "export", target_dir)
        if img: doc.add_picture(img, width=Inches(5.5))
        
    if stats.get("Contract_Count_Graph_Data"):
        doc.add_paragraph("■ 계약 건수")
        img = create_basic_bar_chart(stats["Contract_Count_Graph_Data"], "계약 건수", COLOR_YELLOW, file_name, "contract", target_dir)
        if img: doc.add_picture(img, width=Inches(5.5))

    if stats.get("Sales_Graph_Data"):
        doc.add_paragraph("■ 매출 규모(성장성)")
        img = create_basic_bar_chart(stats["Sales_Graph_Data"], "매출 성장", COLOR_BLUE, file_name, "sales", target_dir)
        if img: doc.add_picture(img, width=Inches(5.5))

    # [Section 3] 기술 경쟁력
    doc.add_heading("3. 기술 경쟁력 및 파이프라인", level=1)
    
    tp = data.get("Technology_and_Pipeline") or {}
    
    doc.add_heading("3-1. Market Pain Points", level=2)
    for p in tp.get("Market_Pain_Points") or []: doc.add_paragraph(f"• {p}")
    
    doc.add_heading("3-2. Solution & Core Tech", level=2)
    sol = tp.get("Solution_and_Core_Tech") or {}
    doc.add_paragraph(f"핵심기술: {sol.get('Technology_Name')}")
    for k in sol.get("Key_Features") or []: doc.add_paragraph(f"- {k}")

    doc.add_heading("3-3. 주요 파이프라인 개발 현황", level=2)
    pipe = tp.get("Pipeline_Development_Status") or {}
    doc.add_paragraph(f"• 플랫폼 상세: {pipe.get('Core_Platform_Details')}")
    doc.add_paragraph(f"• 위험도 분석: {pipe.get('Technical_Risk_Analysis')}")
    doc.add_paragraph(f"• 결론: {pipe.get('Technical_Conclusion')}")

    # [Section 4] 주요 인력 및 조직
    doc.add_heading("4. 주요 인력 및 조직", level=1)
    
    kp = data.get("Key_Personnel") or {}
    
    doc.add_heading("4-1. 대표이사 레퍼런스", level=2)
    ceo = kp.get("CEO_Reference") or {}
    doc.add_paragraph(f"■ 성명: {ceo.get('Name', '')}")
    doc.add_paragraph(f"■ 학력 및 경력:\n{ceo.get('Background_and_Education', '')}")
    doc.add_paragraph(f"■ 핵심 역량:\n{ceo.get('Core_Competency', '')}")
    doc.add_paragraph(f"■ 경영 철학:\n{ceo.get('Management_Philosophy', '')}")
    doc.add_paragraph(f"■ VC 관점 평가:\n{ceo.get('VC_Perspective_Evaluation', '')}")

    doc.add_heading("4-2. 조직 역량", level=2)
    team = kp.get("Team_Capability") or {}
    doc.add_paragraph("■ 핵심 임원진:")
    for ex in team.get("Key_Executives") or []: doc.add_paragraph(f"- {ex}")
    doc.add_paragraph(f"■ 조직 강점:\n{team.get('Organization_Strengths', '')}")
    doc.add_paragraph(f"■ 자문단:\n{team.get('Advisory_Board', '')}")

    # [Section 5] 리스크 및 종합 투자 판단
    doc.add_heading("5. 리스크 및 종합 투자 판단", level=1)
    
    doc.add_heading("5-1. 주요 리스크 및 대응", level=2)
    risks = data.get("Key_Risks_and_Mitigation") or []
    if risks:
        tr = doc.add_table(rows=1, cols=2)
        tr.style = "Table Grid"
        tr.rows[0].cells[0].text = "Risk"
        tr.rows[0].cells[1].text = "Mitigation"
        apply_table_style(tr)
        for r in risks:
            row = tr.add_row().cells
            row[0].text = str(r.get("Risk_Factor"))
            row[1].text = str(r.get("Mitigation_Strategy"))

    val_judge = data.get("Valuation_and_Judgment") or {}

    doc.add_heading("5-2. 밸류에이션 추정", level=2)
    val_table = val_judge.get("Valuation_Table") or []
    if val_table:
        vt = doc.add_table(rows=1, cols=4)
        vt.style = "Table Grid"
        vt.rows[0].cells[0].text = "Round"
        vt.rows[0].cells[1].text = "Pre-Money"
        vt.rows[0].cells[2].text = "Post-Money"
        vt.rows[0].cells[3].text = "Comment"
        apply_table_style(vt)
        for v in val_table:
            r = vt.add_row().cells
            r[0].text = str(v.get("Round"))
            r[1].text = str(v.get("Pre_Money"))
            r[2].text = str(v.get("Post_Money"))
            r[3].text = str(v.get("Comment"))
    
    # 2. 밸류에이션 산정 로직 (NEW - 상세 출력)
    logic = val_judge.get("Valuation_Logic_Detail") or {}
    if logic:
        doc.add_paragraph("") # Spacer
        p_logic = doc.add_paragraph()
        p_logic.add_run("■ 밸류에이션 산정 로직 (Analyst Opinion)\n").bold = True
        
        peers = ", ".join(logic.get("Peer_Group") or [])
        p_logic.add_run(f"• 비교 기업(Peer Group): {peers}\n")
        p_logic.add_run(f"• 적용 지표: {logic.get('Applied_Multiple', '-')}\n")
        p_logic.add_run(f"• 적용 이익: {logic.get('Target_Net_Income', '-')}\n")
        p_logic.add_run(f"• 산출 근거:\n{logic.get('Calculation_Rationale', '-')}")
    else:
        doc.add_paragraph("상세 밸류에이션 로직 데이터가 없습니다.")

    doc.add_heading("5-3. 종합 투자 판단", level=2)
    axes = val_judge.get("Three_Axis_Assessment") or {}
    doc.add_paragraph(f"• 기술성: {axes.get('Technology_Rating')}")
    doc.add_paragraph(f"• 성장성: {axes.get('Growth_Rating')}")
    doc.add_paragraph(f"• 회수성: {axes.get('Exit_Rating')}")
    doc.add_paragraph(f"• 적합 투자자: {val_judge.get('Suitable_Investor_Type')}")

    doc.add_heading("5-4. 종합 결론", level=2)
    doc.add_paragraph(data.get("Final_Conclusion") or "")

    out_path = os.path.join(target_dir, f"{file_name}_검토보고서.docx")
    doc.save(out_path)
    return out_path