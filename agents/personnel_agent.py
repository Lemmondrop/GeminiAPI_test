### CEO, 조직도, 맨파워 PROMPT
import json
from utils import call_gemini, safe_json_loads

PERSONNEL_SCHEMA = """
{
  "Key_Personnel": {
      "CEO_Reference": {
          "Name": "성명",
          "Background_and_Education": "학력 및 주요 경력 (연도별 상세)",
          "Core_Competency": "핵심 역량 (기술/경영/네트워크)",
          "Management_Philosophy": "경영 철학 및 비전",
          "VC_Perspective_Evaluation": "VC 관점에서의 정성적 평가 (리더십, 성공 가능성)"
      },
      "Team_Capability": {
          "Key_Executives": ["주요 임원 상세 이력"],
          "Organization_Strengths": "팀워크, 연구 인력 비중, 조직 문화",
          "Advisory_Board": "자문위원단 및 외부 네트워크"
      }
  }
}
"""

def analyze(pdf_path: str, ceo_name: str) -> dict:
    print(f"   [Personnel Agent] 경영진 및 조직 역량 분석 중...")

    # 1. RAG: CEO 평판 조회
    if ceo_name:
        rag_res = call_gemini(f"'{ceo_name}' 대표이사의 프로필, 인터뷰, 과거 경력을 검색하십시오.", tools=[{"google_search": {}}])
        rag_context = rag_res.get("text", "") if rag_res.get("ok") else ""
    else:
        rag_context = ""

    # 2. Analysis
    prompt = f"""
    당신은 벤처캐피탈(VC)의 인사 검증 담당자입니다.
    자료와 검색된 정보(RAG Context)를 바탕으로 경영진과 조직의 역량을 평가하십시오.

    [RAG Context]
    {rag_context}

    [분석 지침]
    1. **CEO 분석**: 단순 약력이 아니라, 과거의 경험이 현재 사업에 어떻게 기여하는지 '해석'하여 평가를 작성하십시오.
    2. **팀 역량**: C-Level 임원들의 전문성과 팀의 조화를 중점적으로 서술하십시오.

    [Output Schema]
    {PERSONNEL_SCHEMA}
    """
    
    res = call_gemini(prompt, pdf_path=pdf_path)
    if res.get("ok"):
        return safe_json_loads(res["text"])
    else:
        print(f"   [Personnel Agent] Error: {res.get('error')}")
        return {}