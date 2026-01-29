### 기술성, 파이프라인, 특허 PROMPT
import json
from utils import call_gemini, safe_json_loads

TECH_SCHEMA = """
{
  "Technology_and_Pipeline": {
      "Market_Pain_Points": ["기존 시장의 문제점 1 (상세)", "문제점 2 (상세)"], 
      "Solution_and_Core_Tech": {
          "Technology_Name": "핵심 솔루션 명칭",
          "Key_Features": ["기술적 특징 1 (차별점)", "기술적 특징 2 (차별점)"]
      },
      "Pipeline_Development_Status": {
          "Core_Platform_Details": "플랫폼 기술의 원리 및 구조 상세 설명",
          "Technical_Risk_Analysis": "기술적 난관 및 극복 방안",
          "Technical_Conclusion": "기술성 종합 평가 (경쟁 우위 요소)"
      }
  }
}
"""

def analyze(pdf_path: str) -> dict:
    print(f"   [Tech Agent] 기술 경쟁력 및 파이프라인 분석 중...")
    
    prompt = f"""
    당신은 기술 특례 상장 심사역(CTO)입니다.
    제공된 자료를 바탕으로 회사의 기술력과 파이프라인을 심층 분석하십시오.

    [작성 지침]
    1. **Pain Point & Solution**: 시장의 어떤 문제를 우리 기술이 어떻게 해결하는지 인과관계가 명확하게 서술하십시오.
    2. **Core Tech**: 기술의 작동 원리를 비전문가도 이해할 수 있되, 전문 용어를 사용하여 구체적으로 설명하십시오.
    3. **Pipeline**: 각 파이프라인의 개발 단계(TRL, 임상 단계 등)를 명확히 구분하여 적으십시오.

    [Output Schema]
    {TECH_SCHEMA}
    """
    
    res = call_gemini(prompt, pdf_path=pdf_path)
    if res.get("ok"):
        return safe_json_loads(res["text"])
    else:
        print(f"   [Tech Agent] Error: {res.get('error')}")
        return {}