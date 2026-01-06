import os
import glob
import json
from main import save_as_word_report  # 기존 main.py의 함수 재사용

def recover_reports():
    json_dir = "output"        # JSON이 저장된 폴더
    report_dir = "output_report" # 워드가 저장될 폴더
    
    # 1. output 폴더 내의 모든 _refined.json 파일 탐색
    json_files = glob.glob(os.path.join(json_dir, "*_refined.json"))
    
    print(f"총 {len(json_files)}개의 JSON 파일을 발견했습니다. 워드 생성을 시작합니다.")

    for json_path in json_files:
        # 파일명에서 _refined 제외하고 원본 이름 추출
        base_name = os.path.basename(json_path).replace("_refined.json", "")
        
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                refined_data = json.load(f)
            
            # 에러 메시지가 담긴 JSON인지 확인
            if isinstance(refined_data, dict) and "error" not in refined_data:
                print(f">>> [생성 중] {base_name}")
                save_as_word_report(refined_data, base_name, report_dir)
                print(f"    - [완료]")
            else:
                print(f"--- [건너뜀] {base_name}: 유효한 데이터가 아닙니다.")
                
        except Exception as e:
            print(f"!!! [실패] {base_name}: {str(e)}")

if __name__ == "__main__":
    recover_reports()