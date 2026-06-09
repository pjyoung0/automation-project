import os
import feedparser
import pandas as pd
import yfinance as yf
import google.generativeai as genai
import smtplib

from email.message import EmailMessage

# =========================
# Gemini 설정
# =========================

genai.configure(
    api_key=os.environ["GEMINI_API_KEY"]
)

model = genai.GenerativeModel(
    "gemini-2.5-flash"
)

# =========================
# 종목 설정
# =========================

stocks = {
    "삼성전자": "005930.KS",
    "SK하이닉스": "000660.KS",
    "LG전자": "066570.KS",
    "현대차": "005380.KS",
    "두산로보틱스": "454910.KS",
    "코리아써키트": "007810.KS",
    "브이엠": "089970.KQ"
}

# =========================
# 뉴스 수집
# =========================

def get_news(keyword):

    url = (
        f"https://news.google.com/rss/search?q={keyword}"
    )

    feed = feedparser.parse(url)

    results = []

    for item in feed.entries[:10]:

        results.append({
            "company": keyword,
            "title": item.title,
            "link": item.link
        })

    return results

all_news = []

for stock in stocks.keys():

    news = get_news(stock)

    print(stock, len(news))

    all_news.extend(news)

# =========================
# 뉴스 텍스트 생성
# =========================

news_text = ""

for item in all_news:

    news_text += (
        f"[{item['company']}] "
        f"{item['title']}\n"
    )

# =========================
# 주가 정보
# =========================

stock_info = []

stock_info_text = ""

for name, ticker in stocks.items():

    try:

        stock = yf.Ticker(ticker)

        hist = stock.history(period="1y")

        current_price = hist["Close"].iloc[-1]

        high_52w = hist["High"].max()

        drawdown = (
            (current_price - high_52w)
            / high_52w
            * 100
        )

        stock_info.append({
            "종목": name,
            "현재가": round(current_price, 2),
            "52주고점": round(high_52w, 2),
            "52주고점대비(%)": round(drawdown, 2)
        })

        stock_info_text += (
            f"{name}\n"
            f"현재가: {current_price:.0f}\n"
            f"52주 고점: {high_52w:.0f}\n"
            f"52주 고점 대비: {drawdown:.1f}%\n\n"
        )

    except Exception as e:

        print(name, e)

# =========================
# Gemini 분석
# =========================

prompt = f"""
당신은 국내 주식형 펀드를 운용하는 펀드매니저입니다.

아래 뉴스와 주가 정보를 바탕으로 분석하세요.

각 종목별로:

■ 핵심 뉴스
- 내용

■ 긍정 포인트
- 내용

■ 리스크
- 내용

■ 투자 시사점
한 문단

=================

마지막에:

■ 투자 시사점
한 문단

을 작성하세요.

=================

반드시 위 형식을 유지하세요.

뉴스

{news_text}

=================

주가 정보

{stock_info_text}
"""

response = model.generate_content(prompt)

summary = response.text

print(summary)

# =========================
# 엑셀 저장
# =========================

news_df = pd.DataFrame(all_news)

stock_df = pd.DataFrame(stock_info)

summary_df = pd.DataFrame({
    "summary": [summary]
})

with pd.ExcelWriter(
    "daily_report.xlsx"
) as writer:

    news_df.to_excel(
        writer,
        sheet_name="News",
        index=False
    )

    stock_df.to_excel(
        writer,
        sheet_name="StockInfo",
        index=False
    )

    summary_df.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

# =========================
# 이메일 발송
# =========================

EMAIL = os.environ["EMAIL_ADDRESS"]

PASSWORD = os.environ["EMAIL_PASSWORD"]

msg = EmailMessage()

msg["Subject"] = "Daily Stock Brief"

msg["From"] = EMAIL

msg["To"] = EMAIL

msg.set_content(
f"""
[Daily Stock Brief]

{summary}
"""
)

with open(
    "daily_report.xlsx",
    "rb"
) as f:

    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype="octet-stream",
        filename="daily_report.xlsx"
    )

with smtplib.SMTP_SSL(
    "smtp.gmail.com",
    465
) as smtp:

    smtp.login(
        EMAIL,
        PASSWORD
    )

    smtp.send_message(msg)

print("메일 발송 완료")
