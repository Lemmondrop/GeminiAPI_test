import os
import glob
import json
import traceback

# [Modules]
try:
    from docx_generator import save_as_word_report  # 보고서 생성기
except ImportError:
    save_as_word_report = None
    print("⚠️ [Warning] docx_generator.py가 없거나 오류가 있습니다. 보고서 생성이 건너뛰어질 수 있습니다.")

# [Agents] - 5인의 전문가 호출
try:
    from agents import financial_agent
    from agents import market_agent
    from agents import tech_agent
    from agents import personnel_agent
    from agents import valuation_agent
except ImportError as e:
    print(f"⚠️ [Critical Error] 에이전트 파일을 찾을 수 없습니다: {e}")
    exit()

def merge_dictionaries(dicts):
    """여러 딕셔너리를 하나로 병합 (중복 키는 덮어씀)"""
    result = {}
    for d in dicts:
        if d:
            result.update(d)
    return result

def main():
    input_dir = "data"
    output_dir = "output"
    report_dir = "output_report"
    
    for d in [output_dir, report_dir]:
        os.makedirs(d, exist_ok=True)

    pdf_files = glob.glob(os.path.join(input_dir, "*.pdf"))
    if not pdf_files:
        print("📂 PDF 파일이 'data' 폴더에 없습니다.")
        return

    print(f"🚀 총 {len(pdf_files)}개 기업 분석 시작 (Full Multi-Agent System)")

    for i, pdf_path in enumerate(pdf_files):
        file_name = os.path.splitext(os.path.basename(pdf_path))[0]
        json_path = os.path.join(output_dir, f"{file_name}_final.json")
        
        print(f"\n==================================================")
        print(f">>> [{i+1}/{len(pdf_files)}] {file_name} 분석 시작")
        print(f"==================================================")

        try:
            # ------------------------------------------------------------
            # 1. [Financial Agent] 재무 및 기초 정보 추출 (가장 먼저 실행)
            # ------------------------------------------------------------
            print("   [1/5] 💰 Financial Agent (재무/기초정보 분석 중)...")
            fin_data = financial_agent.analyze(pdf_path)
            
            # 메타 데이터 추출 (다른 에이전트에 넘겨줄 정보)
            header = fin_data.get("Report_Header", {})
            company_name = header.get("Company_Name", file_name.split("_")[0])
            ceo_name = header.get("CEO_Name", "")
            industry = header.get("Industry_Classification", "IT/제조/바이오") # 산업분야 추출
            
            print(f"         → 식별된 기업: {company_name} / 대표: {ceo_name} / 산업: {industry}")

            # ------------------------------------------------------------
            # 2. [Market Agent] 시장성 분석
            # ------------------------------------------------------------
            print("   [2/5] 🌍 Market Agent (시장/경쟁사 분석 중)...")
            mkt_data = market_agent.analyze(pdf_path, company_name, industry)

            # ------------------------------------------------------------
            # 3. [Tech Agent] 기술성 분석
            # ------------------------------------------------------------
            print("   [3/5] 🔬 Tech Agent (기술/파이프라인 분석 중)...")
            tech_data = tech_agent.analyze(pdf_path)

            # ------------------------------------------------------------
            # 4. [Personnel Agent] 인력/조직 분석
            # ------------------------------------------------------------
            print("   [4/5] 👥 Personnel Agent (경영진/조직 역량 분석 중)...")
            human_data = personnel_agent.analyze(pdf_path, ceo_name)

            # ------------------------------------------------------------
            # 5. [Valuation Agent] 밸류에이션 심층 분석 (가장 중요)
            # ------------------------------------------------------------
            print("   [5/5] 💎 Valuation Agent (Peer 선정 및 가치 산출 중)...")
            
            # ✅ [수정됨] industry 인자를 추가하여 호출
            val_data = valuation_agent.analyze(pdf_path, company_name, ceo_name, industry)

            # ------------------------------------------------------------
            # [Final Merge] 모든 데이터 병합
            # ------------------------------------------------------------
            final_data = merge_dictionaries([
                fin_data, 
                mkt_data, 
                tech_data, 
                human_data, 
                val_data
            ])
            
            # ------------------------------------------------------------
            # [Save & Report]
            # ------------------------------------------------------------
            # JSON 저장
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(final_data, f, ensure_ascii=False, indent=2)
            print(f"   💾 JSON 데이터 저장 완료")

            # Word 보고서 생성
            if save_as_word_report:
                doc_path = save_as_word_report(final_data, file_name, report_dir)
                print(f"   📝 Word 보고서 생성 완료: {doc_path}")
            else:
                print("   ⚠️ Word 생성기 모듈 오류로 보고서 생성을 건너뜁니다.")

        except Exception as e:
            print(f"   ❌ [Fail] 분석 중 오류 발생: {e}")
            traceback.print_exc()

if __name__ == "__main__":
    main()