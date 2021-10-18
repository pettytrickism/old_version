# 버전업(1.2) - 정보획득(quant_selectitem.py)과 리벨런싱(quant_ts) 프로그램 분리
# 파일이름을 quant.py로 수정하여 사용
import os
import random
import requests  # 설치필요
import sqlite3
import telepot # 설치필요
from pykiwoom.kiwoom import *   # 설치필요
from selenium import webdriver   # 설치필요

class KiwoomPy():
    def __init__(self):
        super().__init__()

        # 전역 변수 선언
        self.DBPath = 'quantDB.db'  # DB 파일위치
        self.nowDateTime = datetime.datetime.now().strftime('%Y%m%d_%H%M')
        self.myAccount = ''   # 사용자 기본 정보
        self.kiwoom = ''  # 키움 API사용

        select_com = input('1:포트폴리오 생성, 2:포트폴리오 업데이트, 3:DB초기화 실행번호는? ')
        if select_com == '1':
            sendText = "포트폴리오를 생성합니다."
            print(sendText)
            self.getCodeList()  # 종목 리스트 획득
            self.getCodeInfo()  # 종목별 정보 획득
            self.getStockInfo()  # 지표별 값 및 종합 순위 결정

            print("포토폴리오 생성을 완료하였습니다.")
            self.send_msg("포토폴리오 생성을 완료하였습니다.")

        elif select_com == '2':
            sendText = "포트폴리오를 업데이트 합니다."
            print(sendText)
            print("공시정보가 업데이트되면 종목정보를 다시 받아야 합니다.")
            self.kiwoomLogin()  # 키움증권 로그인
            selectDateTime = self.chkDateList()  # Date List 선택
            self.updateNowPrice(selectDateTime)  # 현재가 업데이트
            self.getStockInfo()  # 지표별 값 및 종합 순위 결정

            print("포토폴리오 업데이트를 완료하였습니다.")
            self.send_msg("포토폴리오 업데이트를 완료하였습니다.")

        elif select_com == '3':
            sendText = "DB를 초기화합니다."
            print(sendText)
            self.resetDB()  # DB 초기화

        else:
            sendText = "잘못 입력했습니다. 다시 입력하세요."
            print(sendText)

        exit()

    def getCodeList(self):
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

        connect = sqlite3.connect(self.DBPath, isolation_level=None)
        sqlite3.Connection
        cursor = connect.cursor()

        info_df.to_sql('TempStockList', connect, if_exists='replace') # 종목 전체 Updata는 시간이 많이 걸림

        sql = "INSERT INTO StockList (Code, Name, Price, MarketCap, Date) SELECT 종목코드, 종목명, 종가, 시가총액, '%s' FROM TempStockList;" % (self.nowDateTime,)
        cursor.execute(sql)
        connect.close()

        print("완료")

    def getCodeInfo(self):
        print("종목별 투자정보 Insert")

        connect = sqlite3.connect(self.DBPath, isolation_level=None)
        sqlite3.Connection
        cursor = connect.cursor()

        sql = "SELECT ID, Name, Code FROM StockList WHERE Date = '%s' and EPS IS NULL ORDER BY MarketCap;" % (self.nowDateTime)
        cursor.execute(sql)
        rows = cursor.fetchall()

        for row in rows:
            print("%s 주식 정보 가져오기 : " % row[1], end='')

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
                temp_df.index = ['EPS', 'BPS', 'CFPS', 'SPS']
                temp_df.drop(temp_df.columns[0:4], axis=1, inplace=True)

                if str(temp_df.loc['EPS'][0]) != 'nan': eps = int(temp_df.loc['EPS'][0])
                else: eps = 0

                if str(temp_df.loc['BPS'][0]) != 'nan': bps = int(temp_df.loc['BPS'][0])
                else: bps = 0

                if str(temp_df.loc['CFPS'][0]) != 'nan': cfps = int(temp_df.loc['CFPS'][0])
                else: cfps = 0

                if str(temp_df.loc['SPS'][0]) != 'nan': sps = int(temp_df.loc['SPS'][0])
                else: sps = 0

            else:
                eps = 0
                bps = 0
                cfps = 0
                sps = 0

            sql_update_value = "UPDATE StockList SET EPS = %s, BPS = %s, CFPS = %s, SPS = %s, Date = '%s' WHERE ID = %s;"
            sql_update_value = sql_update_value % (eps, bps, cfps, sps, self.nowDateTime, row[0])
            cursor.execute(sql_update_value)

            if eps == 0 or bps == 0 or cfps == 0 or sps == 0:
                print("일부 기업가치 지표 미추출로 0 처리")
            else:
                print("완료")

            delay_time = random.uniform(1, 5)
            time.sleep(random.uniform(0.5, delay_time))

        connect.close()

        print("종목별 투자정보 Insert : 완료")

    def kiwoomLogin(self):
        print("키움 로그인 : ", end='')
        self.kiwoom = Kiwoom()
        self.kiwoom.CommConnect(block=True)  # 블록킹 처리
        print("완료")

    def getStockInfo(self):
        connect = sqlite3.connect(self.DBPath, isolation_level=None)
        sqlite3.Connection
        cursor = connect.cursor()

        print("RER 순위 추가 : ", end='')
        cursor.execute("SELECT Code, Name, Price / EPS FROM StockList WHERE Date = '%s';" % (self.nowDateTime))
        rows = cursor.fetchall()
        df_PER = pd.DataFrame(rows)
        df_PER.columns = ['Code', 'Name', 'PER']
        df_PER.set_index('Code', drop=True, append=False, inplace=True)
        df_PER = df_PER.dropna()
        df_PER = df_PER[df_PER['PER'] > 0]
        df_PER = df_PER.sort_values(by='PER')  # 내림차순 정렬은 ascending=False를 () 안에 추가
        df_PER['rankPER'] = df_PER['PER'].rank()
        print("완료")

        print("PBR 순위 추가 : ", end='')
        cursor.execute("SELECT Code, Price / BPS FROM StockList WHERE Date = '%s';" % (self.nowDateTime))
        rows = cursor.fetchall()
        df_PBR = pd.DataFrame(rows)
        df_PBR.columns = ['Code', 'PBR']
        df_PBR.set_index('Code', drop=True, append=False, inplace=True)
        df_PBR = df_PBR.dropna()
        df_PBR = df_PBR[df_PBR['PBR'] > 0]
        df_PBR = df_PBR.sort_values(by='PBR')
        df_PBR['rankPBR'] = df_PBR['PBR'].rank()
        print("완료")

        print("PCR 순위 추가 : ", end='')
        cursor.execute("SELECT Code, Price / CFPS FROM StockList WHERE Date = '%s';" % (self.nowDateTime))
        rows = cursor.fetchall()
        df_PCR = pd.DataFrame(rows)
        df_PCR.columns = ['Code', 'PCR']
        df_PCR.set_index('Code', drop=True, append=False, inplace=True)
        df_PCR = df_PCR.dropna()
        df_PCR = df_PCR[df_PCR['PCR'] > 0]
        df_PCR = df_PCR.sort_values(by='PCR')
        df_PCR['rankPCR'] = df_PCR['PCR'].rank()
        print("완료")

        print("PSR 순위 추가 : ", end='')
        cursor.execute("SELECT Code, Price / SPS FROM StockList WHERE Date = '%s';" % (self.nowDateTime))
        rows = cursor.fetchall()
        df_PSR = pd.DataFrame(rows)
        df_PSR.columns = ['Code', 'PSR']
        df_PSR.set_index('Code', drop=True, append=False, inplace=True)
        df_PSR = df_PSR.dropna()
        df_PSR = df_PSR[df_PSR['PSR'] > 0]
        df_PSR = df_PSR.sort_values(by='PSR')
        df_PSR['rankPSR'] = df_PSR['PSR'].rank()
        print("완료")

        print("종합 순위 DB Update : ", end='')

        result_df = pd.merge(df_PER, df_PBR, how='inner', left_index=True, right_index=True)
        result_df = pd.merge(result_df, df_PCR, how='inner', left_index=True, right_index=True)
        result_df = pd.merge(result_df, df_PSR, how='inner', left_index=True, right_index=True)

        result_df['RankTotal'] = (result_df['rankPER'] + result_df['rankPBR'] + result_df['rankPCR'] + result_df['rankPSR']).rank()
        result_df = result_df.sort_values(by='RankTotal')
        result_df['Date'] = self.nowDateTime
        result_df.to_sql('StockRank', connect, if_exists='replace')
        connect.close()

        print("완료")
        print("정보입력 완료")

    def chkDateList(self):
        connect = sqlite3.connect(self.DBPath, isolation_level=None)
        sqlite3.Connection
        cursor = connect.cursor()
        sql = "SELECT DISTINCT Date FROM StockList;"
        cursor.execute(sql)
        dateList = cursor.fetchall()
        print("저장된 Date List : ")

        cnt = 1
        for row in dateList:
            print("%s : %s" % (cnt, row))
            cnt = cnt + 1

        selectList = input("매매(리밸런싱)할 Date List ? ")
        return dateList[int(selectList) - 1][0]

        connect.close()

    def updateNowPrice(self, selectDateTime):
        sendText = "현재가를 업데이트 합니다."
        print(sendText)
        connect = sqlite3.connect(self.DBPath, isolation_level=None)
        sqlite3.Connection
        cursor = connect.cursor()
        sql = "SELECT * FROM StockList WHERE Date = '%s';" % selectDateTime
        cursor.execute(sql)
        rows = cursor.fetchall()

        self.nowDateTime = selectDateTime + "_up_" + datetime.datetime.now().strftime('%m%d%H%M')

        for row in rows:
            result_df = self.kiwoom.block_request("opt10001", 종목코드=row[1], output="주식기본정보", next=0)
            price = abs(int(result_df['현재가']))
            sql = "INSERT INTO StockList (Code, Name, MarketCap, Price, EPS, BPS, CFPS, SPS, Date) VALUES ('%s', '%s', %s, %s, %s, %s, %s, %s, '%s')" % (row[1], row[2], row[3], price, row[4], row[5], row[6], row[7], self.nowDateTime,)
            cursor.execute(sql)
            print("%s의 현재가를 업데이트 했습니다." % row[2])
            time.sleep(0.5) # 1초에 5번은 가능하나 0.25초일경우 176번 업데이트 후 차단되어 0.5초로 수정

        print("%s의 현재가 업데이트를 완료하였습니다." % selectDateTime)

    def resetDB(self):
        connect = sqlite3.connect(self.DBPath, isolation_level=None)
        sqlite3.Connection
        cursor = connect.cursor()
        cursor.execute("DELETE FROM StockList;")
        cursor.execute("DELETE FROM StockRank;")
        cursor.execute("DELETE FROM StockHaving;")
        cursor.execute("DELETE FROM QuantList;")
        connect.close()

    def send_msg(self, msg):
        with open('telepot.txt', 'r') as file:
            keys = file.readlines()
            apiToken = keys[0].strip()
            chatId = keys[1].strip()
        bot = telepot.Bot(apiToken)
        bot.sendMessage(chatId, msg)


if __name__ == "__main__":
    kiwoom = KiwoomPy()
