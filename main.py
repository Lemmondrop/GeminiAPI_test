import os
import json
import traceback
import concurrent.futures
from collections import defaultdict
from parser import parse_any_file

# [Modules]
try:
    from docx_generator import save_as_word_report
except ImportError:
    save_as_word_report = None
    print("⚠️ [Warning] docx_generator.py가 없거나 오류가 있습니다.")

# [Agents]
try:
    from agents import financial_agent
    from agents import market_agent
    from agents import tech_agent
    from agents import personnel_agent
    from agents import valuation_agent
except ImportError as e:
    print(f"⚠️ [Critical Error] 에이전트 파일을 찾을 수 없습니다: {e}")
    exit()

# =========================================================
# 💡 [설정] 사용자가 자유롭게 추가/수정할 수 있는 파일 키워드 및 예외 폴더
# =========================================================
TARGET_DOC_TYPES = [
    "IR",       # IR자료, IR_Deck 등
    "재무",     # 재무제표, 감사보고서, Financial 등
    "홍보",     # 홍보브로셔, PR 등
    "기타"      # 키워드에 매칭되지 않는 나머지 파일들
]
SUPPORTED_EXTS = ['.pdf', '.pptx', '.docx', '.xlsx', '.xls', '.csv', '.md']

# 🚫 분석에서 아예 제외할 시스템/작업 폴더 이름들
IGNORE_FOLDERS = ['metadata', 'test', '밸류추정 인수인의 의견', '.DS_Store']

def gather_company_data(base_dir="data"):
    """data 폴더 안의 기업명 폴더 또는 단일 파일을 스캔하여 카테고리별로 수집"""
    company_files = defaultdict(lambda: defaultdict(list))
    
    if not os.path.exists(base_dir):
        print(f"📂 '{base_dir}' 폴더가 없습니다. 폴더를 생성합니다.")
        os.makedirs(base_dir)
        return {}

    for item_name in os.listdir(base_dir):
        item_path = os.path.join(base_dir, item_name)

        if item_name in IGNORE_FOLDERS:
            continue

        if os.path.isdir(item_path):
            company_name = item_name
            for file_name in os.listdir(item_path):
                ext = os.path.splitext(file_name)[1].lower()
                if ext not in SUPPORTED_EXTS: continue

                file_path = os.path.join(item_path, file_name)
                matched = False
                for doc_type in TARGET_DOC_TYPES[:-1]:
                    if doc_type.lower() in file_name.lower():
                        company_files[company_name][doc_type].append(file_path)
                        matched = True
                        break
                if not matched:
                    company_files[company_name]["기타"].append(file_path)

        elif os.path.isfile(item_path):
            ext = os.path.splitext(item_name)[1].lower()
            if ext not in SUPPORTED_EXTS: continue

            raw_name = os.path.splitext(item_name)[0]
            if " IR" in raw_name: company_name = raw_name.split(" IR")[0].strip()
            elif "_" in raw_name: company_name = raw_name.split("_")[0].strip()
            else: company_name = raw_name.split(" ")[0].strip()

            matched = False
            for doc_type in TARGET_DOC_TYPES[:-1]:
                if doc_type.lower() in item_name.lower():
                    company_files[company_name][doc_type].append(item_path)
                    matched = True
                    break

            if not matched:
                if ext == '.pdf': company_files[company_name]["IR"].append(item_path)
                else: company_files[company_name]["기타"].append(item_path)

    return company_files

def merge_dictionaries(dicts):
    result = {}
    for d in dicts:
        if d: result.update(d)
    return result

# 🚨 [신규] 보충 문서 병렬 파싱을 위한 헬퍼 함수
def parse_extra_file(file_path, main_pdf_path):
    if file_path == main_pdf_path: return ""
    parsed_text = ""
    if file_path.lower().endswith('.md'):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                parsed_text = f.read()
        except Exception as e:
            print(f"      ⚠️ Markdown 읽기 실패 ({os.path.basename(file_path)}): {e}")
    else:
        parsed_text = parse_any_file(file_path)
    
    if parsed_text:
        return f"\n\n==== [보충 문서: {os.path.basename(file_path)}] ====\n{parsed_text}"
    return ""

def main():
    output_dir = "output"
    report_dir = "output_report"
    for d in [output_dir, report_dir]: os.makedirs(d, exist_ok=True)

    print(f"🔍 데이터 폴더 스캔 중...")
    company_data_map = gather_company_data("data")
    
    if not company_data_map:
        print("📂 처리할 기업 데이터(폴더)가 없습니다.")
        return

    print(f"🚀 총 {len(company_data_map)}개 기업 멀티-에이전트 병렬 분석 시작\n")

    for i, (company_name, files) in enumerate(company_data_map.items()):
        json_path = os.path.join(output_dir, f"{company_name}_final.json")
        
        print(f"==================================================")
        print(f">>> [{i+1}/{len(company_data_map)}] '{company_name}' 분석 시작")
        print(f"==================================================")

        try:
            # 1. 메인 PDF 추출
            ir_pdfs = [f for f in files.get("IR", []) if f.lower().endswith('.pdf')]
            main_pdf_path = ir_pdfs[0] if ir_pdfs else None

            # 2. 추가 문서 병렬 파싱 (속도 극대화)
            print("   [1/3] 📑 추가 문서 병렬 파싱 중 (Excel, PPT, Word, Markdown 등)...")
            extra_texts = []
            parse_futures = []
            with concurrent.futures.ThreadPoolExecutor() as parse_executor:
                for doc_type, file_list in files.items():
                    for file_path in file_list:
                        parse_futures.append(parse_executor.submit(parse_extra_file, file_path, main_pdf_path))
                
                for f in concurrent.futures.as_completed(parse_futures):
                    res = f.result()
                    if res: extra_texts.append(res)
            
            combined_extra_text = "".join(extra_texts)

            if not main_pdf_path and not combined_extra_text:
                print(f"   ⚠️ [Skip] 분석 가능한 문서가 없습니다.")
                continue

            print(f"        → 확보된 데이터: 메인 PDF({'O' if main_pdf_path else 'X'}), 보충 텍스트({len(combined_extra_text)} bytes)")
            
            # ------------------------------------------------------------
            # [비동기 병렬 Agent 실행부] - 의존성 그래프에 기반한 Phase 제어
            # ------------------------------------------------------------
            print("   [2/3] 🤖 멀티 에이전트 병렬 분석 시작...")
            
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as agent_executor:
                
                # ▶ Phase 1: 의존성이 없는 Financial, Tech 에이전트 동시 실행
                print("      ⚡ [Phase 1] Financial & Tech Agent 병렬 실행 중...")
                future_fin = agent_executor.submit(financial_agent.analyze, main_pdf_path, extra_text=combined_extra_text)
                future_tech = agent_executor.submit(tech_agent.analyze, main_pdf_path, extra_text=combined_extra_text)

                # Financial 결과 대기 (Phase 2를 위한 필수 기초 데이터)
                fin_data = future_fin.result()
                print("      ✅ Financial Agent 완료 (CEO 및 산업군 정보 확보)")

                header = fin_data.get("Report_Header", {})
                ceo_name = header.get("CEO_Name", "")
                industry = header.get("Industry_Classification", "IT/제조/바이오")

                # ▶ Phase 2: Financial에 의존하는 Valuation, Personnel 동시 실행
                print("      ⚡ [Phase 2] Valuation & Personnel Agent 병렬 실행 중...")
                future_val = agent_executor.submit(valuation_agent.analyze, main_pdf_path, company_name, ceo_name, industry, extra_text=combined_extra_text)
                future_human = agent_executor.submit(personnel_agent.analyze, main_pdf_path, ceo_name, extra_text=combined_extra_text)

                # Valuation 결과 대기 (Phase 3을 위한 필수 동기화 데이터)
                val_data = future_val.result()
                print("      ✅ Valuation Agent 완료 (Peer Group 상장사 명단 확보)")

                val_judge = val_data.get("Valuation_and_Judgment", {})
                logic = val_judge.get("Valuation_Logic_Detail", {})
                peer_list = logic.get("Step5_Final_Peers", [])
                if not peer_list: peer_list = val_judge.get("Step5_Final_Peers", [])
                if not peer_list: peer_list = logic.get("stage4_final_peers") or logic.get("Step4_Final_Peers") or []
                
                peer_names = ", ".join(peer_list) if peer_list else "관련 산업 상장사"
                market_sync_instruction = f"\n\n🚨 [필독 - 분석 지시]: 이번 분석의 경쟁사 비교표에는 반드시 다음 Peer Group 기업 중 일부를 포함하십시오: {peer_names}"
                market_extra_text = combined_extra_text + market_sync_instruction

                # ▶ Phase 3: Valuation에 의존하는 Market 에이전트 단독 실행
                print("      ⚡ [Phase 3] Market Agent 실행 중...")
                future_mkt = agent_executor.submit(market_agent.analyze, main_pdf_path, company_name, industry, extra_text=market_extra_text)

                # 모든 스레드의 결과물 최종 수집
                tech_data = future_tech.result()
                human_data = future_human.result()
                mkt_data = future_mkt.result()
                
                print("      ✅ 모든 Agent 분석 완료!")

            # ------------------------------------------------------------
            # [데이터 병합 및 저장]
            # ------------------------------------------------------------
            print("   [3/3] 💾 데이터 병합 및 보고서 생성 중...")
            final_data = merge_dictionaries([fin_data, mkt_data, tech_data, human_data, val_data])
            
            # Valuation 에이전트가 만든 종합 등급을 Report_Header로 안전하게 이동
            if "Investment_Rating" in val_data:
                if "Report_Header" not in final_data: final_data["Report_Header"] = {}
                final_data["Report_Header"]["Investment_Rating"] = val_data["Investment_Rating"]

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(final_data, f, ensure_ascii=False, indent=2)

            if save_as_word_report:                
                orig_name = os.path.basename(main_pdf_path) if main_pdf_path else ""
                doc_path = save_as_word_report(final_data, company_name, report_dir, orig_name)
                print(f"   🎉 분석 및 Word 보고서 생성 완료: {doc_path}")

        except Exception as e:
            print(f"   ❌ [Fail] 분석 중 오류 발생: {e}")
            traceback.print_exc()

if __name__ == "__main__":
    main()