# 버전업(1.3) - 선택 없이(이전자료 자동 삭제) 정보획득(quant_gsi.py)과 리벨런싱간 현재가 판단하여 순위 재산정(quant_ts) 프로그램 분리

import os
import random
import requests  # 설치필요
import sqlite3
import telepot # 설치필요
from pykiwoom.kiwoom import *   # 설치필요
from selenium import webdriver   # 설치필요

# 전역 변수 선언
DBPath = 'quantDB.db'  # DB 파일위치
nowDateTime = datetime.datetime.now().strftime('%Y%m%d%H%M')
myAccount = ''  # 사용자 기본 정보
kiwoom = ''  # 키움 API사용

def resetDB():
    connect = sqlite3.connect(DBPath, isolation_level=None)
    sqlite3.Connection
    cursor = connect.cursor()
    cursor.execute("DELETE FROM StockList;")
    cursor.execute("DELETE FROM StockRank;")
    cursor.execute("DELETE FROM StockHaving;")
    cursor.execute("DELETE FROM QuantList;")
    connect.close()

def getCodeList():
    print("파일 다운로드 : ", end='')
    dataFolder = "C:\\Users\\lpureall\\PycharmProjects\\QuantInvest\\down"
    filelist = os.listdir(dataFolder) # 파일 다운로드 전에 모든 파일 삭제
    for filename in filelist:
        filePath = dataFolder + "\\" + filename
        if os.path.isfile(filePath):
            os.remove(filePath)

    options = webdriver.ChromeOptions()
    # options.add_argument('headless') # 화면 숨기면 다운안될 가능성있음
    # options.add_argument('window-size=1920x1080')
    # options.add_argument("disable-gpu")

    options.add_experimental_option("prefs", {
        "download.default_directory": dataFolder,
        "download.prompt_for_download": False,
        "download.directory_upgrade" : True,
        "safebrowsing.enabled" : True,
        "plugins.always_open_pdf_externally": True
    })

    path = "chromedriver.exe"
    driver = webdriver.Chrome(path, options=options)

    driver.get("http://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd?menuId=MDC0201020101")
    time.sleep(5)  #  창이 모두 열릴 때 까지 5초 기다립니다.
    driver.find_element_by_xpath('//*[@id="MDCSTAT015_FORM"]/div[2]/div/p[2]/button[2]/img').click()
    time.sleep(5)  #  다운로드가 될때 까지 5초 기다립니다.
    driver.find_element_by_xpath('//*[@id="ui-id-1"]/div/div[1]/a').click()
    time.sleep(5)  #  다운로드가 될때 까지 5초 기다립니다.

    filelist = os.listdir(dataFolder)
    while len(filelist) == 0: # 5초 이후에도 다운로드 안되면 5 더 기다림
        print(".", end='')
        time.sleep(5)
        filelist = os.listdir(dataFolder)
    filelist.sort(reverse=True)  # 파일이 여러개일 경우 가장 큰 하나만 불러옴
    filePath = dataFolder + '\\' + filelist[0]

    driver.quit()

    print("완료")
    print("파일명 : %s" % filelist[0])

    print("파일 업로드 : ", end='')
    def changeCode(code):
        code = str(code)
        code = '0' * (6 - len(code)) + code
        return code

    temp_df = pd.read_excel(filePath)
    info_df = temp_df[['종목코드', '종목명', '종가', '시가총액', '상장주식수', '거래량', '시장구분']].copy()
    info_df['종목코드'] = info_df['종목코드'].apply(changeCode)
    info_df = info_df.sort_values(by=['시가총액'])
    info_df.reset_index(drop=True, inplace=True) # 정렬로 인덱스 변경에 따른 인덱스 재설정

    # 조건를 충족하지 않는 데이터를 필터링하여 새로운 변수에 저장합니다.
    noPrice = info_df['거래량'] == 0
    info_df = info_df[~noPrice]  # ~ : 틸데라고 하며 반대조건 증 지금은 거래량에 데이터가 0인 경우 제외됨

    konex = info_df['시장구분'] == "KONEX"
    info_df = info_df[~konex]  # ~ : 틸데라고 하며 반대조건. KONEX제외(개인은 거래가 제한, 3천만원 예수금 필요함)

    cnt = len(info_df) * 0.2  # 시가총액 하위 20%만 선택
    info_df = info_df.loc[:cnt]

    print("완료")

    print("StockList 테이블에 종목 Insert : ", end='')

    connect = sqlite3.connect(DBPath, isolation_level=None)
    sqlite3.Connection
    cursor = connect.cursor()

    info_df.to_sql('TempStockList', connect, if_exists='replace') # 종목 전체 Updata는 시간이 많이 걸림

    sql = "INSERT INTO StockList (Code, Name, Price, MarketCap, Date) SELECT 종목코드, 종목명, 종가, 시가총액, '%s' FROM TempStockList;" % (nowDateTime)
    cursor.execute(sql)
    connect.close()

    print("완료")

def getCodeInfo():
    print("종목별 투자정보 Insert")

    connect = sqlite3.connect(DBPath, isolation_level=None)
    sqlite3.Connection
    cursor = connect.cursor()

    sql = "SELECT ID, Name, Code FROM StockList WHERE Date = '%s' and EPS IS NULL ORDER BY MarketCap;" % (nowDateTime)
    cursor.execute(sql)
    rows = cursor.fetchall()

    cnt = 1

    for row in rows:
        print("(%s / %s) %s 주식 정보 가져오기 : " % (cnt, len(rows), row[1]), end='')

        finance_url = 'http://comp.fnguide.com/SVO2/ASP/SVD_Invest.asp?pGB=1&cID=&MenuYn=Y&ReportGB=B&NewMenuID=105&stkGb=701&gicode=A' + str(row[2])
        finance_page = requests.get(finance_url, verify=False)
        time.sleep(2)
        if finance_page.text.find('error2.htm') == -1:  # 일부 주식은 투자지표가 오류로 되어 안나타남 예) 094800 맵스리얼티1
            finance_text = finance_page.text.replace('(원)', '') # 일부 주식은 (원) 이 없음 예) 096300 베트남개발1
            finance_tables = pd.read_html(finance_text)
            temp_df = finance_tables[1]
            temp_df = temp_df.set_index(temp_df.columns[0])
        else:
            temp_df = [0]

        if len(temp_df) >= 23 :  # 일부 주식은 CFPS, SPS가 없으므로 조회하지 않음  예) 900290 GRT
            temp_df = temp_df.loc[['EPS계산에 참여한 계정 펼치기', 'BPS계산에 참여한 계정 펼치기', 'CFPS계산에 참여한 계정 펼치기', 'SPS계산에 참여한 계정 펼치기']]
            temp_df.index = ['EPS', 'BPS', 'CFPF', 'SPS']
            temp_df.drop(temp_df.columns[0:4], axis=1, inplace=True)

            if str(temp_df.loc['EPS'][0]) != 'nan': eps = int(temp_df.loc['EPS'][0])
            else: eps = 0

            if str(temp_df.loc['BPS'][0]) != 'nan': bps = int(temp_df.loc['BPS'][0])
            else: bps = 0

            if str(temp_df.loc['CFPF'][0]) != 'nan': cfpf = int(temp_df.loc['CFPF'][0])
            else: cfpf = 0

            if str(temp_df.loc['SPS'][0]) != 'nan': sps = int(temp_df.loc['SPS'][0])
            else: sps = 0

        else:
            eps = 0
            bps = 0
            cfpf = 0
            sps = 0

        sql_update_value = "UPDATE StockList SET EPS = %s, BPS = %s, CFPF = %s, SPS = %s, Date = '%s' WHERE ID = %s;"
        sql_update_value = sql_update_value % (eps, bps, cfpf, sps, nowDateTime, row[0])
        cursor.execute(sql_update_value)

        if eps == 0 or bps == 0 or cfpf == 0 or sps == 0:
            print("일부 기업가치 지표 미추출로 0 처리")
        else:
            print("완료")

        delay_time = random.uniform(1, 5)
        time.sleep(random.uniform(0.5, delay_time))

        cnt += 1

    connect.close()

    print("종목별 투자정보 Insert : 완료")

def send_msg(msg):
    with open('telepot.txt', 'r') as file:
        keys = file.readlines()
        apiToken = keys[0].strip()
        chatId = keys[1].strip()
    bot = telepot.Bot(apiToken)
    bot.sendMessage(chatId, msg)

# 프로그램 시작
sendText = "포트폴리오를 생성합니다."
print(sendText)
resetDB()  # DB 초기화
getCodeList()  # 종목 리스트 획득
getCodeInfo()  # 종목별 정보 획득

print("포토폴리오 생성을 완료하였습니다.")
send_msg("포토폴리오 생성을 완료하였습니다.")

exit()
