import os
import requests
from bs4 import BeautifulSoup
import openpyxl
from datetime import datetime, timedelta
from difflib import SequenceMatcher # ✨ 추가됨: 유사도 측정을 위한 라이브러리

# -------------------- [설정값] --------------------

MEMBER_XLSX_FILENAME = "memberlist.xlsx"
MAX_NEWS_PER_COMPANY = 5
STOCK_KEYWORDS_TO_EXCLUDE = ["주가", "증시", "코스피", "코스닥", "목표주가", "투자의견", "매수", "매도", "상한가", "하한가", "특징주", "증권"]
OUTPUT_HTML_FILENAME = "member_news.html"

# -------------------- [✨ 새로운 제목 유사도 비교 함수] --------------------
def is_similar_by_words(title1, title2, threshold=0.5):
    """단어 집합의 유사도(자카드 유사도)를 계산하여 중복 여부를 판단합니다."""
    words1 = set(title1.split())
    words2 = set(title2.split())
    
    if not words1 or not words2:
        return False
        
    intersection = len(words1.intersection(words2))
    union = len(words1.union(words2))
    
    similarity = intersection / union if union > 0 else 0
    
    return similarity >= threshold

# -------------------- [날짜 입력 함수] --------------------
def get_date_input(prompt, default_date):
    """사용자로부터 날짜를 YYYY-MM-DD 형식으로 입력받습니다."""
    while True:
        date_str = input(f"{prompt} (예: {default_date}) [기본값: {default_date}]: ").strip()
        if not date_str:
            return default_date
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            return date_str
        except ValueError:
            print("❌ 날짜 형식이 올바르지 않습니다. YYYY-MM-DD 형식으로 다시 입력해주세요.")

# -------------------- [1단계: 엑셀에서 회원사 이름 읽기] --------------------
def get_member_names(filename):
    """지정된 엑셀 파일의 C열에서 회원사 목록을 읽어옵니다."""
    try:
        workbook = openpyxl.load_workbook(filename)
        sheet = workbook.active
        names = [row[2].value for row in sheet.iter_rows(min_row=2) if row[2].value and row[1].value]
        print(f"✅ 엑셀 파일에서 총 {len(names)}개의 회원사를 찾았습니다.")
        return names
    except FileNotFoundError:
        print(f"❌ 오류: '{filename}'을 찾을 수 없습니다. 파이썬 파일과 같은 폴더에 있는지 확인하세요.")
        return None
    except Exception as e:
        print(f"❌ 엑셀 파일 처리 중 오류 발생: {e}")
        return None

# -------------------- [2단계: 회사 이름으로 구글 뉴스 검색 (✨수정됨)] --------------------
def search_google_news(company_name, count, start_date, end_date):
    """뉴스 검색 후, 핵심 단어 기반으로 중복을 제거하고 최신순으로 정렬합니다."""
    print(f"-> '{company_name}' 관련 뉴스를 검색합니다... ({start_date}~{end_date})")
    
    exclude_query = " ".join([f'-"{keyword}"' for keyword in STOCK_KEYWORDS_TO_EXCLUDE])
    search_query = f'"{company_name}" {exclude_query} after:{start_date} before:{end_date}'
    encoded_query = requests.utils.quote(search_query)
    url = f"https://news.google.com/rss/search?q={encoded_query}&hl=ko&gl=KR&ceid=KR:ko"
    
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, "xml")
        items = soup.find_all("item", limit=count * 3) # 중복 제거를 위해 3배수 검색
        
        candidate_news = []
        for item in items:
            raw_title = item.title.text if item.title else ""
            title = raw_title.rsplit(' - ', 1)[0].strip() if ' - ' in raw_title else raw_title
            if not title: continue

            link = item.link.text if item.link else "#"
            press = item.source.text if item.source else "언론사 불명"
            pub_date_str = item.pubDate.text if item.pubDate else ""
            
            dt_obj = None
            if pub_date_str:
                try:
                    dt_obj = datetime.strptime(pub_date_str.replace(" GMT", ""), "%a, %d %b %Y %H:%M:%S")
                    date_formatted = dt_obj.strftime("%m/%d")
                except ValueError:
                    date_formatted = "날짜 오류"
            
            candidate_news.append({
                "title": title, "link": link, "press": press,
                "date": date_formatted, "datetime_obj": dt_obj
            })
        
        # ✨ 수정됨: 새로운 중복 제거 로직
        unique_news = []
        for news_item in candidate_news:
            is_duplicate = False
            # 이미 추가된 уникальный 기사들과 제목 비교
            for unique_item in unique_news:
                if is_similar_by_words(news_item["title"], unique_item["title"]):
                    is_duplicate = True
                    break
            if not is_duplicate:
                unique_news.append(news_item)
            
            if len(unique_news) >= count:
                break
        
        # 날짜 최신순으로 최종 정렬
        unique_news.sort(key=lambda x: x["datetime_obj"] or datetime.min, reverse=True)
        return unique_news
        
    except requests.exceptions.RequestException as e:
        print(f"오류: '{company_name}' 뉴스 검색 중 네트워크 오류 발생: {e}")
        return []
    except Exception as e:
        print(f"오류: '{company_name}' 뉴스 파싱 중 오류 발생: {e}")
        return []

# -------------------- [3단계: HTML 테이블 생성] --------------------
def generate_member_news_html(all_news_data):
    """전체 뉴스 데이터를 받아 동적 rowspan을 적용한 HTML 테이블을 생성합니다."""
    html_content = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>회원사 이슈</title>
</head>
<body>
<table width="800" border="0" cellpadding="0" cellspacing="0" align="center">
    <tbody>
        <tr>
            <td colspan="4" height="50" style="background-color: #f8f9fa; color:#333; font-size:16px; font-weight: 700; padding-left:15px; border-top: 2px solid #305eb3;">
                표2. 회원사 이슈
            </td>
        </tr>
        <tr>
            <td colspan="4" valign="top">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-bottom:1px solid #e2e2e2">
    """

    for company_name, articles in all_news_data.items():
        if not articles:
            articles = [{"title": "해당 기간에 관련 기사가 없습니다.", "link": "#", "press": "", "date": ""}]
        
        rowspan = len(articles)
        
        for i, article in enumerate(articles):
            html_content += "<tr>\n"
            if i == 0:
                html_content += f'''
    <td rowspan="{rowspan}" style="background-color: #f9f7ff;text-align: center;font-size:13px;color:#305eb3;font-weight:700;border-top:1px solid #e2e2e2" width="120" valign="middle">
        {company_name}
    </td>
'''
            if "관련 기사가 없습니다" in article["title"]:
                 html_content += f'''
    <td colspan="3" style="padding:10px;border-top:1px solid #e2e2e2;color:#777;font-size:13px;">
        {article["title"]}
    </td>
'''
            else:
                html_content += f'''
    <td style="padding:10px;border-top:1px solid #e2e2e2">
        <a href="{article["link"]}" target="_blank" style="text-decoration: none;color:#222;font-size:13px;">{article["title"]}</a>
    </td>
    <td width="100" style="background-color: #f5f5f5;text-align: center;font-size:13px;color:#222;border-top:1px solid #e2e2e2">
        {article["press"]}
    </td>
    <td width="60" style="background-color: #f5f5f5;color:#222222;text-align: center;font-size:13px;border-top:1px solid #e2e2e2">
        {article["date"]}
    </td>
'''
            html_content += "</tr>\n"
            
    html_content += """
                </table>
            </td>
        </tr>
    </tbody>
</table>
</body>
</html>
    """
    return html_content

# -------------------- [메인 실행 부분] --------------------
def main():
    """스크립트의 메인 실행 함수"""
    today = datetime.now().date()
    default_start_date = (today - timedelta(days=7)).strftime("%Y-%m-%d")
    default_end_date = today.strftime("%Y-%m-%d")

    print("--- 뉴스 검색 기간 설정 ---")
    start_date = get_date_input("시작 날짜를 입력하세요", default_start_date)
    end_date = get_date_input("종료 날짜를 입력하세요", default_end_date)
    print("--------------------------\n")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    member_xlsx_path = os.path.join(script_dir, MEMBER_XLSX_FILENAME)
    
    company_names = get_member_names(member_xlsx_path)
    
    if company_names is None:
        print("프로세스를 종료합니다.")
        return

    all_news_data = {}
    for name in company_names:
        news = search_google_news(name, MAX_NEWS_PER_COMPANY, start_date, end_date)
        all_news_data[name] = news
        
    final_html = generate_member_news_html(all_news_data)
    
    try:
        output_path = os.path.join(script_dir, OUTPUT_HTML_FILENAME)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_html)
        print(f"\n🎉 성공! '{output_path}' 파일이 생성되었습니다.")
    except IOError as e:
        print(f"\n❌ 오류: HTML 파일을 저장할 수 없습니다. {e}")

if __name__ == "__main__":
    main()