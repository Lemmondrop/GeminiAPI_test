import json
from pydantic import BaseModel, Field
from typing import List, Optional
from utils import call_gemini, safe_json_loads

# =========================================================================
# 1. [Pydantic 스키마 정의] - 기존 TECH_SCHEMA 문자열을 완벽하게 대체
# =========================================================================
class SolutionAndCoreTech(BaseModel):
    Technology_Name: str = Field(description="핵심 솔루션 명칭")
    Key_Features: List[str] = Field(description="기술적 특징 및 차별점 배열 (예: ['차별점 1', '차별점 2'])")

class PipelineDevelopmentStatus(BaseModel):
    Core_Platform_Details: str = Field(description="플랫폼 기술의 원리 및 구조 상세 설명 (각 파이프라인의 TRL, 임상 단계 등 명시할 것)")
    Technical_Risk_Analysis: str = Field(description="기술적 난관 및 극복 방안")
    Technical_Conclusion: str = Field(description="기술성 종합 평가 (경쟁 우위 요소)")

class TechnologyAndPipeline(BaseModel):
    Market_Pain_Points: List[str] = Field(description="기존 시장의 문제점 배열 (예: ['문제점 1 상세', '문제점 2 상세'])")
    Solution_and_Core_Tech: SolutionAndCoreTech
    Pipeline_Development_Status: PipelineDevelopmentStatus

class TechResponseSchema(BaseModel):
    Technology_and_Pipeline: TechnologyAndPipeline

# =========================================================================
# 2. [분석 엔진(Agent) 메인 실행부]
# =========================================================================
def analyze(pdf_path: str, extra_text: str = "") -> dict:
    print(f"   [Tech Agent] 기술 경쟁력 및 파이프라인 분석 중...")
    
    # 🚨 프롬프트 지시사항은 그대로 유지하고 하드코딩된 Output Schema만 제거
    prompt = f"""
    당신은 기술 특례 상장 심사역(CTO)입니다.
    제공된 메인 자료(PDF)와 **[보충 문서 데이터]**를 바탕으로 회사의 기술력과 파이프라인을 심층 분석하십시오.

    [보충 문서 데이터]
    {extra_text}

    [작성 지침]
    1. **Pain Point & Solution**: 시장의 어떤 문제를 우리 기술이 어떻게 해결하는지 인과관계가 명확하게 서술하십시오.
    2. **Core Tech**: 기술의 작동 원리를 비전문가도 이해할 수 있되, 전문 용어를 사용하여 구체적으로 설명하십시오.
    3. **Pipeline**: 각 파이프라인의 개발 단계(TRL, 임상 단계 등)를 명확히 구분하여 적으십시오.
    """
    
    # 🚨 [핵심 변경] response_schema 파라미터를 넘겨 JSON 출력을 100% 통제합니다.
    res = call_gemini(prompt, pdf_path=pdf_path, response_schema=TechResponseSchema)
    
    if res.get("ok"):
        return safe_json_loads(res["text"])
    else:
        print(f"   [Tech Agent] Error: {res.get('error')}")
        return {}