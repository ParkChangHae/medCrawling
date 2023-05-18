import csv
import json
import random
import re
import time
import traceback
from collections import defaultdict
from datetime import datetime, timedelta

import requests
from openpyxl import Workbook
from openpyxl import load_workbook
from selenium import webdriver
from selenium.common import StaleElementReferenceException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from seleniumwire import webdriver
# A package to have a chromedriver always up-to-date.
from webdriver_manager.chrome import ChromeDriverManager
from fake_useragent import UserAgent



FINAL_SAVE_XLSX_FILE_NAME = './경쟁사_QnA게시판_현황.xlsx'
DEFAULT_THRESHOLD_DATE = datetime(2022,12,1)
TEACHER_URL_INFO = './강사 정보.csv'



def chrome_proxy():

    """
    


    proxy_list = [
        "dc.pr.oxylabs.io:10000",
        "dc.jp-pr.oxylabs.io:12000",
        "dc.kr-pr.oxylabs.io:14000",
        "dc.hk-pr.oxylabs.io:16000",
        "dc.tw-pr.oxylabs.io:18000",
        "dc.vn-pr.oxylabs.io:20000",
        "dc.th-pr.oxylabs.io:22000",
        "dc.my-pr.oxylabs.io:24000",
        "dc.us-pr.oxylabs.io:30000",
        "dc.mx-pr.oxylabs.io:32000",
        "dc.ca-pr.oxylabs.io:34000",
        "dc.de-pr.oxylabs.io:40000",
        "dc.fr-pr.oxylabs.io:42000",
        "dc.nl-pr.oxylabs.io:44000",
        "dc.gb-pr.oxylabs.io:46000",
        "dc.ro-pr.oxylabs.io:48000",
    ]

    ENDPOINT = random.choice(proxy_list)

    wire_options = {
        "proxy": {
            "http": f"http://{USERNAME}:{PASSWORD}@{ENDPOINT}",
            "https": f"http://{USERNAME}:{PASSWORD}@{ENDPOINT}",
        }
    }
    """


def execute_driver():
    options = Options()
    options.add_argument('--headless')

    for i in range (0, 5):
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options, seleniumwire_options=chrome_proxy() )

        driver.get("https://ip.oxylabs.io/")

        print(f'\nYour IP is: {driver.find_element(By.CSS_SELECTOR, "pre").text}')

        driver.quit()


def load_from_xslx():

    try:
        wb = load_workbook(FINAL_SAVE_XLSX_FILE_NAME)

        last_update_date = wb['metadata']['A1'].value.split(':')[1].strip()

        today_date = datetime.now()
        threshold_date = datetime.strptime(last_update_date, '%Y-%m-%d')

        time_gap = (today_date - threshold_date).days

        if (time_gap > 0):
            wb['QnA게시판 현황'].insert_rows(4, time_gap)  # 새롭게 생성해야하는 만큼 Row insert!
            wb.save(FINAL_SAVE_XLSX_FILE_NAME)

        elif (time_gap == 0):
            pass
        else:
            pass
            # ERROR!!! 발생 불가능 케이스
    except Exception:
        return DEFAULT_THRESHOLD_DATE
        #파일 존재하지 않을 때

    return threshold_date


def save_to_xslx(teacher_info_dict, threshold_date):

    try:
        wb = load_workbook(FINAL_SAVE_XLSX_FILE_NAME)
        data_sheet = wb['Qna게시판 현황']
        sheet_metadata = wb['metadata']

    except Exception:
        #파일이 없는 경우
        wb = Workbook()
        data_sheet = wb.active
        data_sheet.title = "QnA게시판 현황"
        sheet_metadata = wb.create_sheet('metadata')


    # Set the start and end dates
    start_date = datetime.now()

    # Start from the cell A4
    row = 4
    col = 1

    current_date = start_date
    while current_date >= threshold_date:
        data_sheet.cell(row=row, column=col).value = current_date.strftime("%Y/%m/%d")
        current_date -= timedelta(days=1)
        row += 1  # move to the next row

    j = 2
    for site_name in teacher_info_dict.keys():

        for subject in teacher_info_dict[site_name].keys():
            for teacher in teacher_info_dict[site_name][subject].keys():

                # 가로줄  #세로줄
                data_sheet.cell(row=1, column=j).value = site_name
                data_sheet.cell(row=2, column=j).value = subject
                data_sheet.cell(row=3, column=j).value = teacher

                for date in teacher_info_dict[site_name][subject][teacher]['qna_cnt'].keys():
                    current_teacher_dict = teacher_info_dict[site_name][subject][teacher]

                    current_teacher_dict['qna_cnt'][date]['총질문수'] = current_teacher_dict['qna_cnt'][date]['답변완료'] + \
                                                                    current_teacher_dict['qna_cnt'][date]['답변대기']

                    offset = (start_date - datetime.strptime(date, '%Y-%m-%d')).days + 4

                    data_sheet.cell(row=offset, column=j).value = \
                    teacher_info_dict[site_name][subject][teacher]['qna_cnt'][date]['총질문수']

                j += 1

    sheet_metadata.sheet_state = 'hidden'
    sheet_metadata['A1'] = 'final_update : ' + datetime.now().strftime('%Y-%m-%d')

    # Save the workbook to a file
    wb.save(FINAL_SAVE_XLSX_FILE_NAME)

def week_of_month(dt):
    first_day_of_month = dt.replace(day=1)
    day_of_month = dt.day
    adjusted_dom = day_of_month + first_day_of_month.weekday()

    return (adjusted_dom - 1) // 7 + 1

def load_teacher_info(file_name = '강사 정보.csv'):
    teacher_info_dict = defaultdict( lambda : defaultdict( lambda  : defaultdict (lambda : defaultdict (list))))

    with open(file_name, 'r', encoding='utf-8-sig') as file:
        reader = csv.reader(file)
        for row in reader:
            teacher_info_dict[row[0]][row[1]][row[2]] = {'url' : row[3], 'qna_cnt' : defaultdict( lambda: defaultdict(int))}

    return teacher_info_dict

def is_contains_date(text):
    # 정규 표현식 패턴 정의
    date_pattern = re.compile(r'\d{4}/\d{2}/\d{2}')
    date_pattern2 = re.compile(r'\d{4}.\d{2}.\d{2}')

    # 텍스트에서 날짜를 찾음
    match = date_pattern.search(text)
    match2 = date_pattern2.search(text)

    # 날짜가 발견되면 True 반환, 그렇지 않으면 False 반환
    return (match is not None) or (match2 is not None)

def is_match_date(text):
    # 정규 표현식 패턴 정의
    date_pattern = re.compile(r'\d{4}/\d{2}/\d{2}')
    date_pattern2 = re.compile(r'\d{4}.\d{2}.\d{2}')

    # 텍스트에서 날짜를 찾음
    match = date_pattern.match(text)
    match2 = date_pattern2.match(text)

    # 날짜가 발견되면 True 반환, 그렇지 않으면 False 반환
    return (match is not None) or (match2 is not None)

def contains_keyword(text, keyword):
    pattern = r'\b' + keyword + r'\b'
    if re.search(pattern, text):
        return True
    return False

def save_to_dict(current_teacher_dict, date, status="답변완료"):

    current_teacher_dict['qna_cnt'][date][status] += 1


class NextTeacherRaise(Exception):
    pass

def crawling_qna(teacher_info_dict):

    # 브라우저 옵션 설정
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # 옵션: 창이 나타나지 않게 함
    chrome_options.add_argument("--enable-javascript")
    #browser2 = Chrome(chrome_options=chrome_options)

    MAX_RETRIES = 10

    threshold_date = load_from_xslx()

    # 웹 드라이버 시작
    browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    for site_name in teacher_info_dict.keys():
        for subject in teacher_info_dict[site_name].keys():
            for teacher_name in teacher_info_dict[site_name][subject].keys():

                current_teacher_dict = teacher_info_dict[site_name][subject][teacher_name]
                original_qna_url = current_teacher_dict['url']

                for i in range(85, 9999999):
                    try:
                        print("i : " + str(i))
                        retries = 0
                        while retries < MAX_RETRIES:
                            try:

                                if(site_name == "메가스터디"):

                                    bring_amount = 500

                                    target_url = original_qna_url + "startRowNum="+str(bring_amount*(i-1))+"&itemPerPage="+str(bring_amount)
                                    #target_url = original_qna_url + "startRowNum="+str(100000)+"&itemPerPage=100"

                                    start_time = time.time()  # 코드 실행 전의 시간

                                    #response = requests.get(target_url, headers={'User-agent': 'Mozilla/5.0'})

                                    #if response.status_code == 200:
                                    #    elements = BeautifulSoup(response.content, "html.parser")

                                    browser.get(target_url)
                                    browser.implicitly_wait(10)

                                    end_time = time.time()  # 코드 실행 후의 시간
                                    execution_time = end_time - start_time  # 실행 시간 계산
                                    print(f"데이터 크롤링 시간 : {execution_time}초")

                                    elements = WebDriverWait(browser, 20).until( EC.presence_of_element_located((By.TAG_NAME, 'body')))

                                    try :
                                        match = re.search(r'\"aData\":(.*?), \"bData\"', elements.text, re.DOTALL)
                                        raw_json_data = match.group(1)
                                        fixed_elements_text = re.sub(r'"BT_TITLE":.*?"QNAFLG":".*?",', '',
                                                                     raw_json_data)
                                        json_data = json.loads(fixed_elements_text)
                                    except Exception:
                                        raise StaleElementReferenceException

                                    for data_element in json_data:

                                        if(int(data_element['CNT']) > 0):
                                            status = "답변완료"
                                        else:
                                            status = "답변대기"

                                        date = data_element['REG_DT']

                                        if (datetime.strptime(date, "%Y-%m-%d") < threshold_date):
                                            raise NextTeacherRaise
                                            # 다음 선생님으로 넘어가야함

                                        save_to_dict(current_teacher_dict, date, status)

                                    if( len(json_data) <=0 ):
                                        raise NextTeacherRaise

                                elif(site_name == "이투스"):

                                    target_url = original_qna_url + str(i)

                                    start_time = time.time()  # 코드 실행 전의 시간

                                    browser.get(target_url)
                                    browser.implicitly_wait(10)

                                    end_time = time.time()  # 코드 실행 후의 시간
                                    execution_time = end_time - start_time  # 실행 시간 계산
                                    print(f"데이터 크롤링 시간 : {execution_time}초")

                                    """
                                    target_url = original_qna_url + str(i)

                                    browser.implicitly_wait(10)
                                    browser.get(target_url)
                                    browser.implicitly_wait(10)

                                    elements = WebDriverWait(browser, 10).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, 'table.tbl.subcomm_tbl_board')))

                                    img_elements = elements.find_elements(By.TAG_NAME,'img')

                                    target_image_src = "https://img.etoos.com/sub2016/common/ico_a.png"

                                    matching_images = 0
                                    for img_element in img_elements:
                                        if img_element.get_attribute('src') == target_image_src:
                                            matching_images += 1

                                    origin_text_list = str(elements.text).split("\n")

                                    for entry in origin_text_list:
                                        if ('[강좌내용]' in entry or '[학습방법]' in entry or '[교재내용]' in entry ):
                                            date = entry.split()[-1:][0].replace('.','-')

                                            if( datetime.strptime(date, "%Y-%m-%d") < THRESHOLD_DATE ):
                                                raise NextTeacherRaise
                                                #다음 선생님으로 넘어가야함

                                            save_to_dict(current_teacher_dict, date)

                                    
                                    """

                                    elements = WebDriverWait(browser, 20).until(
                                        EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, 'ul.board_list')))

                                    origin_text_list = str(elements.text).split("\n")

                                    for j in range(len(origin_text_list)-1, -1, -1):

                                        if( origin_text_list[j]  == '대기' or origin_text_list[j] == '완료' ) :
                                            pass

                                        elif( contains_keyword( origin_text_list[j], '공지' )  or not is_contains_date(origin_text_list[j]) or is_match_date(origin_text_list[j]) ) :
                                            origin_text_list.pop(j)

                                    for j in range(0, len(origin_text_list), 2):

                                        if( origin_text_list[j] == '완료' ):
                                            status = "답변완료"
                                        else:
                                            status = "답변대기"

                                        date = origin_text_list[j+1].split()[-1:][0].replace('.','-')

                                        if (len(origin_text_list) <= 0 or datetime.strptime(date, "%Y-%m-%d") < threshold_date):
                                            raise NextTeacherRaise
                                            # 다음 선생님으로 넘어가야함

                                        save_to_dict(current_teacher_dict, date, status)

                                elif(site_name == "대성마이맥"):

                                    ua = UserAgent()
                                    userAgnet = ua.random

                                    USERNAME = "fpdlwj4869"
                                    PASSWORD = "Wlrnrhkgkr1!"

                                    entry = ('http://customer-%s:%s@dc.pr.oxylabs.io:10000' %
                                             (USERNAME, PASSWORD))

                                    proxy_dict = {
                                        "http": entry,
                                        "https": entry
                                    }

                                    headers = {
                                        "Origin": "https://m.mimacstudy.com",
                                        "Referer" : "https://m.mimacstudy.com/mobile/tcher/studyQna.ds?coType=M&tcd="+current_teacher_dict['url'],
                                        "User-Agent" : userAgnet
                                    }

                                    params = {
                                        "pid" : "",
                                        "tcd" : current_teacher_dict['url'],
                                        "currPage" : str(i),
                                        "pagePerCnt" : "200",
                                        "myQna" : "N",
                                        "srchWordTxt" : "",
                                        "isScrtN" : "N"
                                    }

                                    ipaddress = requests.get(url="https://ip.oxylabs.io/", proxies=proxy_dict)
                                    print("IP : " + ipaddress.text + " User Agent : " + userAgnet)

                                    start_time = time.time()  # 코드 실행 전의 시간

                                    while(True):
                                        res = requests.get(url="https://m.mimacstudy.com/mobile/tcher/getStudyQnaListByAjax.ds", headers= headers, params=params, proxies=proxy_dict)
                                        json_data = json.loads(res.text)

                                        if (json_data['code'] == 'success'):
                                            json_data = json_data['data']
                                            break

                                    end_time = time.time()  # 코드 실행 후의 시간
                                    execution_time = end_time - start_time  # 실행 시간 계산
                                    print(f"데이터 크롤링 시간 : {execution_time}초")

                                    for data_element in json_data:

                                        if(data_element['qnaType'] == 'Q' ):
                                            status = "답변완료"
                                        else:
                                            status = "답변완료"
                                            continue

                                        date = data_element['regDate']

                                        if (len(json_data) <= 0 or datetime.strptime(date, "%Y/%m/%d") < threshold_date):
                                            raise NextTeacherRaise
                                            # 다음 선생님으로 넘어가야함

                                        save_to_dict(current_teacher_dict, date, status)


                                    print("xx")


                                    """
                                    
                                    

                                    if(i == 1):
                                        target_url = original_qna_url

                                        browser.implicitly_wait(10)
                                        browser.get(target_url)
                                        browser.implicitly_wait(10)

                                        elements = WebDriverWait(browser, 10).until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.tbltype_list')))

                                    elif( i > 1 ):
                                        wait.until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, "div.tbltype_list")))
                                        browser.execute_script("javascript:pageMove("+str(i)+")")

                                        elements = WebDriverWait(browser, 10).until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.tbltype_list')))

                                    origin_text_list = str(elements.text).split("\n")

                                    for j in range(len(origin_text_list)-1, -1, -1):
                                        if('공지' in origin_text_list[j] or not is_contains_date(origin_text_list[j])):
                                            origin_text_list.pop(j)

                                    while(j < len(origin_text_list)):
                                        qna_title = origin_text_list[j]

                                        if( j+1 < len(origin_text_list) and '답변' in origin_text_list[j+1]):
                                            status = "답변완료"
                                            j += 2
                                        else:
                                            status = "답변대기"
                                            j += 1

                                        date = qna_title.split()[-2:-1][0].replace('/','-')

                                        if (datetime.strptime(date, "%Y-%m-%d") < threshold_date):
                                            raise NextTeacherRaise
                                            # 다음 선생님으로 넘어가야함

                                        save_to_dict(current_teacher_dict, date, status)
                                    """

                                break

                            except StaleElementReferenceException:
                                retries += 1
                                print("Exception : " + str(retries))
                                continue

                        if(i % 10 == 0):
                            print(i)
                            for date, status_counts in current_teacher_dict['qna_cnt'].items():
                                print(f"{site_name}_{subject}_{teacher_name} | {date}: 답변완료 {status_counts['답변완료']}개, 답변대기 {status_counts['답변대기']}개, 총 질문 수 {str(int(status_counts['답변완료']) + int(status_counts['답변대기']))}개")

                        #save_to_xslx(teacher_info_dict)

                    except NextTeacherRaise:
                        break
                    except Exception as e:

                        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                        with open(f"{current_time}_Exception.txt", "w") as file:
                            file.write(str(e))
                            file.write(traceback.format_exc())

                        print(e)

    save_to_xslx(teacher_info_dict, threshold_date)

    # 웹 드라이버 종료
    browser.quit()


def main():

    teacher_info_dict = load_teacher_info('강사 정보.csv')
    crawling_qna(teacher_info_dict)


if __name__ == '__main__':
    main()
