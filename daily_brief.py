import os
import feedparser
import pandas as pd
import yfinance as yf
from pykrx import stock
from datetime import datetime, timedelta
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
# KOSPI200 전일 등락률 TOP10
# =========================

try:

    market_day = stock.get_nearest_business_day_in_a_week()

    kospi200 = stock.get_index_portfolio_deposit_file(
        "1028"
    )

    change_df = stock.get_market_price_change(
        market_day,
        market_day,
        market="KOSPI"
    )

    change_df = change_df.loc[
        change_df.index.intersection(kospi200)
    ]

    top_up = (
        change_df
        .sort_values(
            "등락률",
            ascending=False
        )
        .head(10)
    )

    top_down = (
        change_df
        .sort_values(
            "등락률",
            ascending=True
        )
        .head(10)
    )

    market_text = ""

    market_text += (
        "[KOSPI200 상승률 TOP10]\n"
    )

    for _, row in top_up.iterrows():

        market_text += (
            f"- {row['종목명']} "
            f"{row['등락률']:.2f}%\n"
        )

    market_text += "\n"

    market_text += (
        "[KOSPI200 하락률 TOP10]\n"
    )

    for _, row in top_down.iterrows():

        market_text += (
            f"- {row['종목명']} "
            f"{row['등락률']:.2f}%\n"
        )

except Exception as e:

    print(e)

    market_text = ""

# =========================
# Gemini 분석
# =========================

prompt = f"""
당신은 국내 자산운용사 펀드매니저입니다.

아래 뉴스, 주가 정보, 시장 데이터를 참고하여
장 마감 브리핑을 작성하세요.

==================
관심 종목 뉴스
==================

{news_text}

==================
관심 종목 주가 정보
==================

{stock_info_text}

==================
시장 데이터
==================

{market_text}

==================

다음 형식으로 작성하세요.

# 관심 종목 분석

종목별로

■ 핵심 뉴스

■ 긍정 포인트

■ 리스크

■ 투자 시사점

==================

# 시장 분석

■ 상승 종목 공통점

3~5개

■ 하락 종목 공통점

3~5개

■ 강세 섹터

3개

■ 약세 섹터

3개

==================

# 내일 체크포인트

■ 내일 주목할 섹터

■ 내일 체크할 종목

■ 내일 체크할 이벤트

■ 펀드매니저 코멘트

5~10줄

실제 운용역이 작성하는
장 마감 리포트처럼 작성하세요.

마크다운 기호(**, ##)는 사용하지 말고
읽기 쉬운 일반 텍스트 형식으로 작성하세요.

아래 형식으로 작성할 것.

━━━━━━━━━━━━━━━━━━

[삼성전자]

핵심 뉴스
- ...

긍정 포인트
- ...

리스크
- ...

투자 시사점
- ...

━━━━━━━━━━━━━━━━━━

[시장 분석]

상승 종목 공통점
- ...

하락 종목 공통점
- ...

━━━━━━━━━━━━━━━━━━

[내일 체크포인트]

주목할 섹터
- ...

주목할 종목
- ...

체크할 이벤트
- ...

펀드매니저 코멘트
- ...

━━━━━━━━━━━━━━━━━━

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

today_str = datetime.now().strftime(
    "%Y-%m-%d"
)

msg = EmailMessage()

msg["Subject"] = (
    f"[장마감 브리핑] {today_str}"
)

msg["From"] = EMAIL

msg["To"] = EMAIL

# Gemini 결과를 HTML로 변환

html_summary = summary.replace(
    "\n",
    "<br>"
)

# HTML 메일 작성

msg.add_alternative(
    f"""
    <html>
    <body style="
        font-family: Arial;
        line-height: 1.6;
        padding: 20px;
    ">

    <h2>📈 Daily Market Brief</h2>

    {html_summary}

    <br><br>

    <hr>

    <p>
    자동 생성된 장마감 브리핑입니다.
    </p>

    </body>
    </html>
    """,
    subtype="html"
)

# 엑셀 첨부

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

# 메일 발송

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
