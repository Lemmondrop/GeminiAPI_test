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
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# 커스텀 모듈 임포트
# from parser import extract_text_from_pdf  # 1단계가 LLM(PDF 첨부)이므로 사용하지 않음
from processor import refine_pdf_to_json_onecall, JSON_SCHEMA

# 한글 폰트 설정 (Mac/Windows 호환성 고려)
plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False

def add_hyperlink(paragraph, url, text):
    """
    python-docx에서 하이퍼링크를 추가하는 헬퍼
    """
    if not url:
        paragraph.add_run(text)
        return

    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    # 스타일: 파란색 + 밑줄
    c = OxmlElement("w:color")
    c.set(qn("w:val"), "0000FF")
    rPr.append(c)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def set_cell_background(cell, fill_color):
    """셀의 배경색을 설정하는 함수"""
    shading_elm = OxmlElement("w:shd")
    shading_elm.set(qn("w:fill"), fill_color)
    cell._element.get_or_add_tcPr().append(shading_elm)

def apply_center_alignment(target):
    """문단이나 표 셀의 텍스트를 중앙 정렬하는 함수"""
    if hasattr(target, "paragraphs"):  # 셀(Cell)인 경우
        for paragraph in target.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:  # 문단(Paragraph)인 경우
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
    """
    TAM/SAM/SOM Bar 차트 생성 (회의 결과 반영)
    - market_data는 기존처럼:
        [["구분","시장규모"],["TAM",..],["SAM",..],["SOM",..]]  (단일 비교)
      또는 연도별:
        [["연도","TAM","SAM","SOM"],["2023",100,80,50], ...]   (그룹 막대)
      또는 dict 스키마:
        {"TAM_SAM_SOM_Value": {"TAM":{"value":..,"year":..}, ...}} (단일 비교)
    """
    if market_data is None:
        return None

    # -----------------------------
    # 공통 숫자 변환
    # -----------------------------
    def to_float(x):
        try:
            s = str(x).replace(",", "").strip()
            return float(s)
        except:
            return 0.0

    # -----------------------------
    # 1) dict 스키마 지원: TAM_SAM_SOM_Value → 단일 비교용으로 정규화
    # -----------------------------
    unit = ""
    year_hint = None

    if isinstance(market_data, dict):
        v = market_data.get("TAM_SAM_SOM_Value", {}) or {}

        def pick_val(k):
            obj = v.get(k, {})
            if isinstance(obj, dict):
                nonlocal_year = obj.get("year", None)
                nonlocal_unit = obj.get("unit", "")
                return obj.get("value", 0), nonlocal_year, nonlocal_unit
            return obj, None, ""

        tam_v, tam_y, tam_u = pick_val("TAM")
        sam_v, sam_y, sam_u = pick_val("SAM")
        som_v, som_y, som_u = pick_val("SOM")

        # year/unit 힌트(있으면 타이틀/축에 반영)
        year_hint = tam_y or sam_y or som_y
        unit = tam_u or sam_u or som_u or ""

        # list 포맷으로 변환(단일 비교)
        market_data = [
            ["구분", "시장규모(단위 포함)"],
            ["TAM", tam_v],
            ["SAM", sam_v],
            ["SOM", som_v],
        ]

    # market_data는 이제 list일 가능성이 높음
    if not isinstance(market_data, list) or len(market_data) < 2:
        return None

    # -----------------------------
    # 2) 연도별 포맷 감지: ["연도","TAM","SAM","SOM"] 형태면 그룹 막대
    # -----------------------------
    header = market_data[0] if isinstance(market_data[0], list) else []
    header0 = str(header[0]).strip().lower() if header else ""

    is_timeseries = False
    if header and len(header) >= 4:
        # 헤더에 연도/year가 있고, TAM/SAM/SOM 컬럼이 있는지 확인
        cols = [str(c).strip().upper() for c in header]
        if ("연도" in header0 or "year" in header0) and all(k in cols for k in ["TAM", "SAM", "SOM"]):
            is_timeseries = True

    try:
        import matplotlib.pyplot as plt
        import numpy as np

        if is_timeseries:
            # -----------------------------
            # (B) 연도별 그룹 막대 그래프
            # -----------------------------
            cols = [str(c).strip().upper() for c in header]
            idx_year = 0
            idx_tam = cols.index("TAM")
            idx_sam = cols.index("SAM")
            idx_som = cols.index("SOM")

            years = []
            tam_list, sam_list, som_list = [], [], []

            for row in market_data[1:]:
                if not isinstance(row, list) or len(row) < max(idx_som, idx_sam, idx_tam) + 1:
                    continue
                y = str(row[idx_year]).strip()
                years.append(y)
                tam_list.append(to_float(row[idx_tam]))
                sam_list.append(to_float(row[idx_sam]))
                som_list.append(to_float(row[idx_som]))

            if not years:
                return None

            x = np.arange(len(years))
            width = 0.22

            plt.figure(figsize=(9.5, 5.5))
            plt.bar(x - width, tam_list, width, label="TAM", color = "cornflowerblue")
            plt.bar(x,         sam_list, width, label="SAM", color = "khaki")
            plt.bar(x + width, som_list, width, label="SOM", color = "indianred")

            plt.title("시장규모 (TAM/ SAM/ SOM)", fontsize=16, pad=18)
            plt.xticks(x, years)
            plt.grid(axis="y", alpha=0.3)
            plt.legend()

            # 단위 표시(있으면)
            if unit:
                plt.ylabel(unit)

            plt.tight_layout()
            chart_path = os.path.join(target_dir, f"{file_name}_market_chart.png")
            plt.savefig(chart_path, dpi=160)
            plt.close()
            return chart_path

        else:
            # -----------------------------
            # (A) 단일 비교 막대 그래프: TAM/SAM/SOM 3개
            # -----------------------------
            # 기존 포맷: [["구분","시장규모"],["TAM",..],...]
            vmap = {}
            for row in market_data[1:]:
                if not isinstance(row, list) or len(row) < 2:
                    continue
                k = str(row[0]).strip().upper()
                if k in ["TAM", "SAM", "SOM"]:
                    vmap[k] = to_float(row[1])

            tam = vmap.get("TAM", 0.0)
            sam = vmap.get("SAM", 0.0)
            som = vmap.get("SOM", 0.0)

            if max(tam, sam, som) <= 0:
                return None

            labels = ["TAM", "SAM", "SOM"]
            values = [tam, sam, som]

            plt.figure(figsize=(8, 5))
            plt.bar(labels, values)

            title = "시장규모 (TAM/ SAM/ SOM)"
            if year_hint:
                title += f" - {year_hint}"
            plt.title(title, fontsize=16, pad=18)

            if unit:
                plt.ylabel(unit)

            # 값 라벨
            for i, v in enumerate(values):
                plt.text(i, v, f"{v:,.2f}" if v != int(v) else f"{int(v):,}", ha="center", va="bottom", fontweight="bold")

            plt.grid(axis="y", alpha=0.3)
            plt.tight_layout()

            chart_path = os.path.join(target_dir, f"{file_name}_market_chart.png")
            plt.savefig(chart_path, dpi=160)
            plt.close()
            return chart_path

    except Exception as e:
        print(f"   [Chart Error] 시장 차트 생성 실패: {e}")
        return None

def create_financial_chart(financial_table, file_name, target_dir):
    """매출액(Bar) + 영업이익(Line) 이중축 그래프 생성"""
    try:
        if not financial_table or len(financial_table) < 2:
            return None

        years = [str(y) for y in financial_table[0][1:]]

        rev_row = next((row for row in financial_table if "매출" in str(row[0])), None)
        opp_row = next((row for row in financial_table if "영업이익" in str(row[0])), None)

        if not rev_row:
            return None

        def to_float(val):
            v = str(val).replace(",", "").replace("(", "-").replace(")", "").replace("억", "").strip()
            try:
                return float(v)
            except:
                return 0.0

        rev_values = [to_float(v) for v in rev_row[1:]]

        fig, ax1 = plt.subplots(figsize=(10, 6))
        bars = ax1.bar(years, rev_values, color="#34495E", alpha=0.8, width=0.5, label="매출액")
        ax1.set_ylabel("매출액 (억 원)", fontsize=11, fontweight="bold")

        max_val = max(rev_values) if rev_values else 100
        for bar in bars:
            height = bar.get_height()
            ax1.text(
                bar.get_x() + bar.get_width() / 2.0,
                height + (max_val * 0.02),
                f"{int(height):,}",
                ha="center",
                va="bottom",
                fontweight="bold",
            )

        if opp_row:
            opp_values = [to_float(v) for v in opp_row[1:]]
            if len(opp_values) != len(years):
                opp_values = opp_values[: len(years)]

            ax2 = ax1.twinx()
            ax2.plot(years, opp_values, color="#E74C3C", marker="o", linewidth=2, label="영업이익")
            ax2.set_ylabel("영업이익 (억 원)", fontsize=11, fontweight="bold", color="#E74C3C")
            ax2.tick_params(axis="y", labelcolor="#E74C3C")

            lines, labels = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax1.legend(lines + lines2, labels + labels2, loc="upper left")
        else:
            ax1.legend(loc="upper left")

        plt.title(f"[{file_name}] 연도별 실적 추이 및 전망", fontsize=15, pad=20, fontweight="bold")
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
    ceo_name = header_info.get("CEO_Name", data.get("CEO_Name", "확인 필요"))

    # [표지]
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(f"\n{company_name}\n투자 대상기업 검토보고서")
    run.font.size = Pt(24)
    run.bold = True

    date_p = doc.add_paragraph(f"\n작성일: 2025년 00월 00일")
    date_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_page_break()

    # 1. Executive Summary
    h1 = doc.add_heading("Executive Summary", level=1)
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(data.get("Investment_Thesis_Summary", "내용 없음"))

    info_table = doc.add_table(rows=3, cols=2)
    info_table.style = "Table Grid"
    rows_data = [
        ["기업명 / 대표자", f"{company_name} / {ceo_name}"],
        ["담당 심사역", header_info.get("Analyst", "LUCEN Investment Intelligence")],
        ["투자 등급(종합)", header_info.get("Investment_Rating", "N/A")],
    ]
    for i, row in enumerate(rows_data):
        info_table.rows[i].cells[0].text = row[0]
        info_table.rows[i].cells[1].text = str(row[1])
    apply_table_style(info_table, apply_header_color=False)
    doc.add_paragraph()

    # 2. 시장 동향 및 기술 분석
    h2 = doc.add_heading("시장 동향 및 기술 분석", level=1)
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("1. 최신 시장 동향 및 Pain Points", level=2)
    for point in data.get("Problem_and_Solution", {}).get("Market_Pain_Points", []):
        doc.add_paragraph(f"○ {str(point)}")

    doc.add_heading("2. 핵심 기술 및 독점적 해자 (Moat)", level=2)
    tech = data.get("Technology_and_Moat", {})
    p = doc.add_paragraph(f"핵심 기술명: {tech.get('Core_Technology_Name', 'N/A')}")
    if p.runs:
        p.runs[0].bold = True

    for detail in tech.get("Technical_Details", []):
        doc.add_paragraph(f"○ {str(detail)}")

    # 3. 재무 현황 및 성장 전망
    h3 = doc.add_heading("재무 현황 및 성장 전망", level=1)
    h3.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("1. 타겟 시장 규모 및 성장성", level=2)
    market_opp = data.get("Market_Opportunity", {}) or {}

    # ---- (1) Market Definition (VC 제출용 강화) ----
    mdef = market_opp.get("Market_Definition", {}) or {}
    primary_market = (mdef.get("Primary_Market") or "").strip()

    doc.add_heading("1-1. 시장 정의 (Definition)", level=3)
    if primary_market:
        p = doc.add_paragraph()
        r = p.add_run(f"Primary Market: {primary_market}")
        r.bold = True
    else:
        doc.add_paragraph("Primary Market: N/A")

    included = mdef.get("Included_Segments", []) or []
    excluded = mdef.get("Excluded_Segments", []) or []

    if included:
        doc.add_paragraph("Included Segments:", style=None)
        for seg in included[:8]:
            doc.add_paragraph(f"• {str(seg)}")
    else:
        doc.add_paragraph("Included Segments: N/A")

    if excluded:
        doc.add_paragraph("Excluded Segments:", style=None)
        for seg in excluded[:8]:
            doc.add_paragraph(f"• {str(seg)}")
    else:
        doc.add_paragraph("Excluded Segments: N/A")

    # ---- (2) TAM/SAM/SOM 요약 + 차트 ----
    doc.add_heading("1-2. 시장 규모 (TAM/SAM/SOM)", level=3)

    m_chart_path = create_market_chart(market_opp.get("Market_Chart_Data"), file_name, target_dir)
    if m_chart_path and os.path.exists(m_chart_path):
        doc.add_picture(m_chart_path, width=Inches(5.5))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap = doc.add_paragraph(f"[그림 1] {company_name} 시장 규모 비교 (TAM/SAM/SOM)")
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    tam_text = (market_opp.get("TAM_SAM_SOM_Text") or "").strip()
    doc.add_paragraph(tam_text if tam_text else "시장 규모 요약: N/A")

    # ---- (3) Sources (VC 레벨 핵심) ----
    doc.add_heading("1-3. 시장 근거 자료 (Sources)", level=3)
    sources = market_opp.get("Market_Sources", []) or []

    if not sources:
        doc.add_paragraph("시장 근거 자료가 없습니다. (Market_Sources: N/A)")
    else:
        # 표: Claim / Source / Year / Evidence / Link
        src_table = doc.add_table(rows=1, cols=5)
        src_table.style = "Table Grid"
        hdr = src_table.rows[0].cells
        hdr[0].text = "Claim"
        hdr[1].text = "Source"
        hdr[2].text = "Year"
        hdr[3].text = "Evidence"
        hdr[4].text = "Link/ID"
        apply_table_style(src_table)

        for s in sources[:10]:  # 너무 길어지지 않게 상한
            if not isinstance(s, dict):
                continue
            row = src_table.add_row().cells
            row[0].text = str(s.get("Claim", "")).strip()
            row[1].text = str(s.get("Source", "")).strip()
            row[2].text = str(s.get("Year", "")).strip()
            row[3].text = str(s.get("Evidence", "")).strip()

            link_val = str(s.get("URL_or_Identifier", "")).strip()
            # 링크가 URL처럼 보이면 하이퍼링크로, 아니면 텍스트로
            if link_val.startswith("http://") or link_val.startswith("https://"):
                # 셀의 첫 문단에 링크 삽입
                p = row[4].paragraphs[0]
                add_hyperlink(p, link_val, "Open")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                row[4].text = link_val

        doc.add_paragraph("※ 위 자료는 google_search 기반 보강 결과이며, VC 검토를 위해 출처/연도/근거를 표로 명시했습니다.")

    financial_data = data.get("Table_Data_Preview", {}).get("Financial_Table", [])
    if financial_data:
        doc.add_heading("2. 주요 재무 실적 및 추정치", level=2)
        fin_table = doc.add_table(rows=len(financial_data), cols=len(financial_data[0]))
        fin_table.style = "Table Grid"
        for i, row in enumerate(financial_data):
            for j, val in enumerate(row):
                fin_table.rows[i].cells[j].text = str(val)
        apply_table_style(fin_table)

        f_chart_path = create_financial_chart(financial_data, file_name, target_dir)
        if f_chart_path and os.path.exists(f_chart_path):
            doc.add_picture(f_chart_path, width=Inches(5.0))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap = doc.add_paragraph("[그림 2] 연도별 실적 추이")
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 4. 리스크 및 주요 질의
    h4 = doc.add_heading("리스크 분석 및 향후 과제", level=1)
    h4.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("1. 주요 리스크 및 대응 전략", level=2)
    risks = data.get("Key_Risks_and_Mitigation", [])
    if risks:
        risk_table = doc.add_table(rows=1, cols=2)
        risk_table.style = "Table Grid"
        risk_table.rows[0].cells[0].text = "리스크 요인"
        risk_table.rows[0].cells[1].text = "대응 전략"
        apply_table_style(risk_table)

        for r in risks:
            if isinstance(r, dict):
                row = risk_table.add_row().cells
                row[0].text = str(r.get("Risk_Factor", ""))
                row[1].text = str(r.get("Mitigation_Strategy", ""))

    doc.add_heading("2. 주요 질의 및 사후 확인 사항", level=2)
    for q in data.get("Due_Diligence_Questions", []):
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
        if not os.path.exists(d):
            os.makedirs(d)

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
            # 1) 기존 분석 결과 확인 (Skip 로직)
            if os.path.exists(json_path):
                with open(json_path, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                        if isinstance(data, dict) and "error" not in data:
                            print("   - [Skip] 기존 JSON 활용 -> Word 보고서 생성")
                            save_as_word_report(data, file_name, report_dir)
                            continue
                    except:
                        pass

            # 2) PDF 1개당 LLM 호출 1회로 JSON 생성
            refined_data = refine_pdf_to_json_onecall(
                pdf_path=file_path,
                json_schema=JSON_SCHEMA,
                max_output_tokens=16384,
                retry_max_output_tokens=32768
            )
            
            # 실패 처리
            if not isinstance(refined_data, dict) or refined_data.get("error"):
                print(f"   - [실패] JSON 생성 실패: {refined_data.get('message', str(refined_data))[:200]}...")
                continue

            # 3) JSON 저장
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(refined_data, f, ensure_ascii=False, indent=2)

            # 4) 워드 보고서 생성 (JSON 기반)
            print("   - 보고서(Word) 생성 중...")
            report_path = save_as_word_report(refined_data, file_name, report_dir)
            print(f"   - [성공] {report_path}")

        except Exception as e:
            print(f"   - [예외 발생] {str(e)}")

if __name__ == "__main__":
    main()