import os
import time
import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# -------------------- [설정값] --------------------
# 1. 크롤링 관련 설정
NTIS_URL = "https://www.ntis.go.kr/rndgate/eg/un/ra/mng.do"
TARGET_DEPARTMENTS = ["산업통상자원부", "과학기술정보통신부", "중소벤처기업부"]
DEPT_ALIAS = {
    "산업통상자원부": "산업부",
    "과학기술정보통신부": "과기부",
    "중소벤처기업부": "중기부"
}
DEADLINE_THRESHOLD_DAYS = 7

# 2. 파일 경로 설정
DOWNLOAD_DIR = r"다운로드 파일을 저장할 폴더 경로"
EXCEL_FILENAME = "공고목록.xls"
EXCEL_FILE_PATH = os.path.join(DOWNLOAD_DIR, EXCEL_FILENAME)

OUTPUT_HTML_FILENAME = "ntis_projects.html"
OUTPUT_DIR = r"html 파일을 저장할 경로로"
FULL_OUTPUT_PATH = os.path.join(OUTPUT_DIR, OUTPUT_HTML_FILENAME)

# -------------------- [1단계: 엑셀 파일 다운로드 함수] --------------------
def download_excel_file():
    print("1단계: 엑셀 파일 다운로드를 시작합니다...")
    
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": DOWNLOAD_DIR}
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    try:
        driver = webdriver.Chrome(options=options)
    except Exception as e:
        print(f"❌ 크롬 드라이버 실행 오류: {e}")
        return False

    try:
        if os.path.exists(EXCEL_FILE_PATH):
            os.remove(EXCEL_FILE_PATH)
            print(f"기존 '{EXCEL_FILENAME}' 파일을 삭제했습니다.")

        driver.get(NTIS_URL)

        try:
            close_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='popup_footer']//button[text()='닫기']"))
            )
            close_button.click()
            print("팝업창 '닫기' 버튼을 클릭했습니다.")
        except TimeoutException:
            print("팝업창이 발견되지 않았습니다.")

        download_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '리스트 다운로드')]"))
        )
        download_button.click()
        print("'리스트 다운로드' 버튼을 클릭했습니다.")

        for i in range(60):
            if os.path.exists(EXCEL_FILE_PATH):
                print(f"✅ '{EXCEL_FILENAME}' 다운로드 완료!")
                return True
            time.sleep(1)
        
        print("❌ 오류: 60초 내에 파일 다운로드가 완료되지 않았습니다.")
        return False

    finally:
        driver.quit()

# -------------------- [2단계: 엑셀 파일 분석 함수 (정렬 기능 추가)] --------------------
def process_excel_file():
    print("\n2단계: 다운로드한 엑셀 파일 분석을 시작합니다...")
    
    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
    except FileNotFoundError:
        print(f"❌ 오류: '{EXCEL_FILE_PATH}' 파일을 찾을 수 없습니다.")
        return {}

    # 날짜 형식 변환
    df['마감일'] = pd.to_datetime(df['마감일'], errors='coerce')
    df.dropna(subset=['마감일'], inplace=True)

    # 1. 부처명으로 필터링
    filtered_df = df[df['부처명'].isin(TARGET_DEPARTMENTS)].copy()
    
    # 2. 마감일로 필터링
    today = datetime.datetime.now()
    filtered_df = filtered_df[(filtered_df['마감일'] - today).dt.days >= DEADLINE_THRESHOLD_DAYS]

    # --- ✨ 여기가 추가된 부분입니다 ✨ ---
    # 3. 마감일 기준으로 내림차순 정렬 (많이 남은 순)
    filtered_df = filtered_df.sort_values(by='마감일', ascending=False)
    # ------------------------------------

    print(f"총 {len(filtered_df)}개의 유효한 공고를 찾았습니다.")

    # HTML 생성을 위해 데이터 형식 맞추기
    all_announcements = {alias: [] for alias in DEPT_ALIAS.values()}
    for index, row in filtered_df.iterrows():
        dept_name = row['부처명']
        alias = DEPT_ALIAS.get(dept_name, dept_name)
        
        post = {
            "title": row['공고명'],
            "link": row['공고문 바로가기(URL)'],
            "deadline": f"~{row['마감일'].strftime('%m/%d')}"
        }
        all_announcements[alias].append(post)
        
    return all_announcements

# -------------------- [HTML 생성 함수 (수정됨)] --------------------
def generate_html_file(all_data):
    # (HTML 헤더 부분은 동일)
    html_content = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>신규 정부과제 안내</title>
</head>
<body>
<table width="800" border="0" cellpadding="0" cellspacing="0" align="center">
    <tbody>
        <tr>
            <td rowspan="2" valign="top" style="background-color:#634abb;"></td>
            <td colspan="8" height="50" style="background-color: #634abb;color:#fff;font-size:16px;font-weight: 700; padding-left:15px;">
                신규 정부과제 안내
            </td>
            <td rowspan="2" valign="top" style="background-color:#634abb;"></td>
        </tr>
        <tr>
            <td colspan="8" valign="top">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #fff;">
    """

    # --- ✨ 여기가 수정된 부분입니다 ✨ ---
    # 루프를 도는 순서를 지정
    ordered_aliases = ["산업부", "과기부", "중기부"]
    is_first_department = True

    for alias in ordered_aliases:
        # all_data 딕셔너리에서 현재 부처(alias)의 공고 목록을 가져옴
        dept_posts = all_data.get(alias, [])
        
        # 공고가 없으면 다음 부처로 넘어감
        if not dept_posts:
            continue
        
        rowspan = len(dept_posts)
        border_style = "" if is_first_department else 'border-top:1px solid #e2e2e2;'
        
        for i, post in enumerate(dept_posts):
            html_content += "<tr>\n"
            # 첫 번째 행에만 부처명과 '본공고' 셀을 추가 (rowspan 적용)
            if i == 0:
                html_content += f'''
    <td rowspan="{rowspan}" style="background-color: #f0f0f0;color:#305eb3;text-align: center;font-size:13px;font-weight:700;padding:10px 0;{border_style}" width="75" valign="top">
        [{alias}]
    </td>
    <td rowspan="{rowspan}" style="background-color: #fdfff4;color:#305eb3;text-align: center;font-size:13px;font-weight:700;padding:10px 0;{border_style}" width="63" valign="top">
        본공고
    </td>
'''
            # 모든 행에 공고 제목과 마감일 셀 추가
            # 첫 번째 행에만 상단 테두리 스타일 적용
            td_style = f'padding:10px;{border_style}' if i == 0 else 'padding:10px;'
            deadline_style = f'background-color: #f5f5f5;color:#222222;text-align: center;font-size:13px;{border_style}' if i == 0 else 'background-color: #f5f5f5;color:#222222;text-align: center;font-size:13px;'
            
            html_content += f'''
    <td style="{td_style}">
        <a href="{post["link"]}" target="_blank" style="text-decoration: none;color:#222;font-size:13px;">{post["title"]}</a>
    </td>
    <td style="{deadline_style}" width="75">{post["deadline"]}</td>
'''
            html_content += "</tr>\n"
        
        is_first_department = False
    
    # (HTML 푸터 부분은 동일)
    html_content += """
                    <tr>
                        <td height="29" style="background-color: #634abb;border-top:1px solid #e2e2e2;" colspan="4"></td>
                    </tr>
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
    if download_excel_file():
        all_data = process_excel_file()
        final_html = generate_html_file(all_data)
        
        try:
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            with open(FULL_OUTPUT_PATH, "w", encoding="utf-8") as f:
                f.write(final_html)
            print(f"\n✅ 최종 HTML 파일 생성 완료! '{FULL_OUTPUT_PATH}'")
        except IOError as e:
            print(f"\n❌ 오류: HTML 파일을 저장할 수 없습니다. {e}")

if __name__ == "__main__":
    main()
