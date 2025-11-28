import requests
from bs4 import BeautifulSoup
import datetime
import os

# -------------------- [설정값] --------------------

# 검색할 키워드 목록
TOPICS = [
    "산업부", "반도체", "모빌리티", "통신", "헬스케어",
    "기계로봇", "군사", "AI", "우주항공", "ESG"
]

# ✨ 추가됨: 검색에서 제외할 키워드 목록
KEYWORDS_TO_EXCLUDE = ["투자", "MOU", "취임", "협력", "선정", "징역"]

# 키워드별로 가져올 기사 수
ARTICLES_PER_TOPIC = 1

# 최종 저장될 HTML 파일 이름
OUTPUT_HTML_FILENAME = "keyword_news.html"

# -------------------- [날짜 입력 함수] --------------------
def get_date_input(prompt, default):
    """사용자로부터 날짜를 YYYY-MM-DD 형식으로 입력받습니다."""
    while True:
        date_str = input(f"{prompt} (예: {default}) [기본값: {default}]: ").strip()
        if not date_str:
            return default
        try:
            datetime.datetime.strptime(date_str, "%Y-%m-%d")
            return date_str
        except ValueError:
            print("❌ 날짜 형식이 올바르지 않습니다. YYYY-MM-DD 형식으로 다시 입력해주세요.")

# -------------------- [뉴스 검색 함수 (✨수정됨)] --------------------
def search_google_news_rss(topic, count, start_date, end_date):
    """지정된 기간과 키워드로 구글 뉴스 RSS를 검색합니다."""
    
    exclude_query = " ".join([f'-"{keyword}"' for keyword in KEYWORDS_TO_EXCLUDE])
    search_query = f'"{topic}" "기술" {exclude_query} after:{start_date} before:{end_date}'
    
    print(f"-> '{topic} 기술' 관련 뉴스를 검색합니다... ({start_date}~{end_date})")
    
    encoded_query = requests.utils.quote(search_query)
    url = f"https://news.google.com/rss/search?q={encoded_query}&hl=ko&gl=KR&ceid=KR:ko"
    
    results = []
    try:
        res = requests.get(url, timeout=10)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, "xml")
        
        items = soup.find_all("item", limit=count)
        for item in items:
            # ✨ 수정됨: 제목에서 ' - 언론사' 부분 제거
            raw_title = item.title.text if item.title else "제목 없음"
            if ' - ' in raw_title:
                title = raw_title.rsplit(' - ', 1)[0].strip()
            else:
                title = raw_title

            link = item.link.text if item.link else "#"
            press = item.source.text if item.source else "언론사 불명"
            pubdate = item.pubDate.text if item.pubDate else ""
            
            news_date = ""
            if pubdate:
                try:
                    dt = datetime.datetime.strptime(pubdate.replace(" GMT", ""), "%a, %d %b %Y %H:%M:%S")
                    news_date = f"{dt.month}/{dt.day}"
                except ValueError:
                    news_date = "날짜 오류"
            
            results.append({
                "topic": topic,
                "title": title,
                "link": link,
                "press": press,
                "date": news_date
            })
    except Exception as e:
        print(f"오류: '{topic}' 뉴스 검색 중 오류 발생: {e}")

    return results

# -------------------- [HTML 생성 함수 (최종 수정)] --------------------
def generate_table_html(news_list):
    """뉴스 목록으로 제목을 포함한 HTML 테이블을 생성합니다."""

    # ✨ 수정됨: 제목과 전체 틀을 포함하는 외부 테이블 구조 추가
    # --- HTML 헤더 부분 ---
    html_content = """
<table width="800" border="0" cellpadding="0" cellspacing="0" align="center">
    <tbody>
        <tr>
            <td height="40" style="background-color: #389c92;color:#fff;font-size:16px;font-weight: 700;padding-left:20px">
                국내외 임베디드 산업 동향
            </td>
        </tr>
        <tr>
            <td valign="top">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-bottom:1px solid #e2e2e2">
"""

    # --- HTML 본문 (뉴스 목록) 부분 ---
    for news in news_list:
        html_content += f"""
<tr>
    <td width="100" style="background-color: #f9f7ff;text-align: center;font-size:13px;color:#305eb3;font-weight:700;border-top:1px solid #e2e2e2; padding: 10px 0;">
        {news['topic']}
    </td>
    <td style="padding:10px;border-top:1px solid #e2e2e2">
        <a href="{news['link']}" target="_blank" style="text-decoration: none;color:#222;font-size:13px;">{news['title']}</a>
    </td>
    <td width="100" style="background-color: #edfff5;text-align: center;font-size:13px;color:#222;border-top:1px solid #e2e2e2">
        {news['press']}
    </td>
    <td width="60" style="background-color: #f5f5f5;color:#222222;text-align: center;font-size:13px;border-top:1px solid #e2e2e2">
        {news['date']}
    </td>
</tr>
"""

    # ✨ 수정됨: 외부 테이블 구조를 닫는 태그 추가
    # --- HTML 푸터 부분 ---
    html_content += """
                </table>
            </td>
        </tr>
    </tbody>
</table>
"""
    
    return html_content

    # --- HTML 테이블 종료 ---
    html_content += "</table>"
    
    return html_content

# -------------------- [메인 실행 부분] --------------------
def main():
    """스크립트의 메인 실행 함수"""
    today = datetime.date.today()
    default_start = (today - datetime.timedelta(days=7)).strftime("%Y-%m-%d")
    default_end = today.strftime("%Y-%m-%d")
    
    start_date = get_date_input("시작 날짜를 입력하세요", default_start)
    end_date = get_date_input("종료 날짜를 입력하세요", default_end)
    print("-" * 20)

    all_news = []
    for topic in TOPICS:
        news = search_google_news_rss(topic, ARTICLES_PER_TOPIC, start_date, end_date)
        all_news.extend(news)
    
    # 최종 HTML 생성 (✨수정됨)
    final_html_content = generate_table_html(all_news)

    # 파일로 저장
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(script_dir, OUTPUT_HTML_FILENAME)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_html_content)
        print(f"\n🎉 성공! '{output_path}' 파일이 생성되었습니다.")
    except IOError as e:
        print(f"\n❌ 오류: HTML 파일을 저장할 수 없습니다. {e}")


if __name__ == "__main__":
    main()