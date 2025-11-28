import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urlparse
import re
from datetime import datetime
import time

def extract_news_info(url):
    """
    뉴스 기사 URL에서 제목, 날짜, 언론사 정보를 추출합니다.
    """
    try:
        # User-Agent 헤더 추가 (일부 사이트에서 봇 차단 방지)
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        response = requests.get(url.strip(), headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # 제목 추출 (여러 패턴 시도)
        title = None
        title_selectors = [
            'h1.headline',  # 일반적인 헤드라인
            'h1.title',
            'h1',
            '.article-head h1',
            '.article-title',
            'title'
        ]
        
        for selector in title_selectors:
            element = soup.select_one(selector)
            if element:
                title = element.get_text().strip()
                break
        
        if not title:
            title = soup.title.string if soup.title else "제목 없음"
        
        # 언론사 추출
        press = None
        
        # 도메인 기반으로 언론사 추출
        domain = urlparse(url).netloc
        domain_to_press = {
            'news.naver.com': '네이버뉴스',
            'www.chosun.com': '조선일보',
            'www.donga.com': '동아일보',
            'www.joongang.co.kr': '중앙일보',
            'www.hani.co.kr': '한겨레',
            'www.khan.co.kr': '경향신문',
            'www.yna.co.kr': '연합뉴스',
            'news.kbs.co.kr': 'KBS',
            'imnews.imbc.com': 'MBC',
            'news.sbs.co.kr': 'SBS'
        }
        
        press = domain_to_press.get(domain)
        
        # 메타 태그에서 언론사 정보 추출
        if not press:
            press_selectors = [
                'meta[property="og:site_name"]',
                'meta[name="author"]',
                '.press',
                '.source',
                '.media'
            ]
            
            for selector in press_selectors:
                element = soup.select_one(selector)
                if element:
                    if element.name == 'meta':
                        press = element.get('content', '').strip()
                    else:
                        press = element.get_text().strip()
                    if press:
                        break
        
        if not press:
            press = domain
        
        # 날짜 추출
        date = None
        date_selectors = [
            'meta[property="article:published_time"]',
            'meta[name="article:published_time"]',
            'time',
            '.date',
            '.publish-date',
            '.article-date'
        ]
        
        for selector in date_selectors:
            element = soup.select_one(selector)
            if element:
                if element.name == 'meta':
                    date_text = element.get('content', '')
                elif element.name == 'time':
                    date_text = element.get('datetime', '') or element.get_text()
                else:
                    date_text = element.get_text()
                
                # 날짜 형식 파싱
                if date_text:
                    # ISO 형식 (2024-01-15T10:30:00+09:00)
                    iso_match = re.search(r'(\d{4}-\d{2}-\d{2})', date_text)
                    if iso_match:
                        date = iso_match.group(1)
                        break
                    
                    # 한국어 형식 (2024년 1월 15일, 2024.01.15 등)
                    korean_match = re.search(r'(\d{4})[년\-\.](\d{1,2})[월\-\.](\d{1,2})', date_text)
                    if korean_match:
                        year, month, day = korean_match.groups()
                        date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                        break
        
        if not date:
            date = "날짜 없음"
        
        return {
            'url': url.strip(),
            'title': title,
            'date': date,
            'press': press
        }
    
    except Exception as e:
        print(f"Error processing {url}: {str(e)}")
        return {
            'url': url.strip(),
            'title': f"오류: {str(e)}",
            'date': "날짜 없음",
            'press': "언론사 없음"
        }

def process_news_links(txt_file_path, output_excel_path):
    """
    TXT 파일에서 뉴스 링크를 읽어와 정보를 추출하고 엑셀 파일로 저장합니다.
    """
    try:
        # TXT 파일 읽기
        with open(txt_file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        
        # 쉼표로 구분된 URL들 추출
        urls = [url.strip() for url in content.split(',') if url.strip()]
        
        if not urls:
            print("유효한 URL이 없습니다.")
            return
        
        print(f"총 {len(urls)}개의 URL을 처리합니다...")
        
        # 각 URL에서 정보 추출
        news_data = []
        for i, url in enumerate(urls, 1):
            print(f"처리 중... ({i}/{len(urls)}) {url[:50]}...")
            
            info = extract_news_info(url)
            news_data.append(info)
            
            # 서버 부하 방지를 위한 딜레이
            time.sleep(1)
        
        # DataFrame 생성
        df = pd.DataFrame(news_data, columns=['url', 'title', 'date', 'press'])
        df.columns = ['링크', '기사제목', '기사날짜', '언론사명']
        
        # 날짜순 정렬 (오래된 순서부터)
        def parse_date_for_sorting(date_str):
            """정렬을 위한 날짜 파싱"""
            if date_str == "날짜 없음" or not date_str:
                return datetime.min
            try:
                # YYYY-MM-DD 형식 파싱
                return datetime.strptime(date_str, '%Y-%m-%d')
            except:
                return datetime.min
        
        df['정렬용_날짜'] = df['기사날짜'].apply(parse_date_for_sorting)
        df = df.sort_values('정렬용_날짜').drop('정렬용_날짜', axis=1)
        df = df.reset_index(drop=True)
        
        # 엑셀 파일로 저장
        df.to_excel(output_excel_path, index=False, engine='openpyxl')
        print(f"\n완료! 결과가 '{output_excel_path}' 파일로 저장되었습니다.")
        
        # 결과 미리보기
        print("\n=== 추출된 데이터 미리보기 ===")
        print(df.head())
        
        return df
    
    except Exception as e:
        print(f"파일 처리 중 오류 발생: {str(e)}")
        return None

# 사용 예시
if __name__ == "__main__":
    import os
    
    # 현재 스크립트 파일이 있는 디렉토리 가져오기
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 파일 경로 설정 (스크립트와 같은 폴더)
    txt_file = os.path.join(script_dir, "news_link.txt")  # 입력 TXT 파일 경로
    excel_file = os.path.join(script_dir, "news_data.xlsx")  # 출력 엑셀 파일 경로
    
    # 파일 존재 확인
    if not os.path.exists(txt_file):
        print(f"파일을 찾을 수 없습니다: {txt_file}")
        print("현재 디렉토리의 txt 파일들:")
        for file in os.listdir(script_dir):
            if file.endswith('.txt'):
                print(f"  - {file}")
        
        # txt 파일이 하나만 있다면 자동으로 사용
        txt_files = [f for f in os.listdir(script_dir) if f.endswith('.txt')]
        if len(txt_files) == 1:
            txt_file = os.path.join(script_dir, txt_files[0])
            print(f"자동으로 선택된 파일: {txt_files[0]}")
        else:
            exit()
    
    # 뉴스 링크 처리 실행
    result = process_news_links(txt_file, excel_file)
    
    if result is not None:
        print(f"\n총 {len(result)}개의 기사 정보가 추출되었습니다.")
    
    # 단일 URL 테스트용 함수
    def test_single_url(url):
        """단일 URL 테스트용"""
        print(f"테스트 URL: {url}")
        info = extract_news_info(url)
        print("추출된 정보:")
        for key, value in info.items():
            print(f"  {key}: {value}")
    
    # 테스트 예시 (주석 해제하여 사용)
    # test_single_url("https://www.example-news.com/article/123")