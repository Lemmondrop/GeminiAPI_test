import google.generativeai as genai
import os

genai.configure(api_key=os.environ["GEMINI_API_KEY"])

for m in genai.list_models():
    # generateContent 지원하는 모델만 보기
    if "generateContent" in getattr(m, "supported_generation_methods", []):
        print(m.name, m.supported_generation_methods)