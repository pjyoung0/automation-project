import os
import google.generativeai as genai

genai.configure(
    api_key=os.environ["GEMINI_API_KEY"]
)

model = genai.GenerativeModel(
    "gemini-2.5-flash"
)

response = model.generate_content(
    "삼성전자 투자포인트 3줄"
)

print(response.text)
