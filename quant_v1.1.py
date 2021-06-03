# 버전업(1.1) - 텔레그램 메시지 전송 기능 추가, 주식 정보 재활용(현재가 업데이트로 정보획득 시간 절약)
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

        select_com = input('1:구성/매매, 2:리밸런싱, 3:업데이트/매매, 4:DB초기화 실행번호는? ')
        if select_com == '1':
            sendText = "포트폴리오 구성 및 매매 합니다."
            print(sendText)
            self.send_msg(sendText)
            self.getCodeList()  # 종목 리스트 획득
            self.getCodeInfo()  # 종목별 정보 획득
            self.getStockInfo()  # 지표별 값 및 종합 순위 결정

            selectDateTime = self.nowDateTime  # 현재 구성된 포트폴리오
            nowTime = int(datetime.datetime.now().strftime('%H%M'))
            print("현재시간 : %s" % datetime.datetime.now().strftime('%H:%M'))

            while nowTime not in range(900, 1530):
                sendText = "매매 대기중.(현재시간 : %s)" % nowTime
                print(sendText)
                self.send_msg(sendText)
                time.sleep(3600)  # 60분 대기
                nowTime = int(datetime.datetime.now().strftime('%H%M'))

            self.kiwoomLogin()  # 키움증권 로그인
            self.getQuantList(selectDateTime)  # 포트폴리오 작성
            self.runTrading(selectDateTime)  # 주식 매매


        elif select_com == '2':
            sendText = "리밸런싱을 합니다."
            print(sendText)
            self.send_msg(sendText)
            selectDateTime = self.chkDateList()  # Date List 선택
            nowTime = int(datetime.datetime.now().strftime('%H%M'))
            print("현재시간 : %s" % datetime.datetime.now().strftime('%H:%M'))

            while nowTime not in range(900, 1530):
                sendText = "매매 대기중.(현재시간 : %s)" % nowTime
                print(sendText)
                self.send_msg(sendText)

                print("매매 대기중.(현재시간 : %s)" % nowTime)
                time.sleep(3600)  # 30분 대기
                nowTime = int(datetime.datetime.now().strftime('%H%M'))

            self.kiwoomLogin()  # 키움증권 로그인
            self.getQuantList(selectDateTime)  # 포트폴리오 작성
            self.runTrading(selectDateTime)  # 주식 매매

        elif select_com == '3':
            sendText = "포트폴리오 업데이트 및 매매를 합니다."
            print(sendText)
            self.send_msg(sendText)
            print("공시정보가 업데이트되면 종목정보를 다시 받아야 합니다.")
            selectDateTime = self.chkDateList()  # Date List 선택

            self.kiwoomLogin()  # 키움증권 로그인
            self.updateNowPrice(selectDateTime)  # 현재가 업데이트
            self.getStockInfo()  # 지표별 값 및 종합 순위 결정

            nowTime = int(datetime.datetime.now().strftime('%H%M'))
            print("현재시간 : %s" % datetime.datetime.now().strftime('%H:%M'))

            while nowTime not in range(900, 1530):
                sendText = "매매 대기중.(현재시간 : %s)" % nowTime
                print(sendText)
                self.send_msg(sendText)
                state = self.kiwoom.GetConnectState()
                if state == 1:
                    print("로그인 중")
                    self.send_msg("로그인 중")
                else:
                    print("로그아웃됨")
                    self.send_msg("로그아웃됨")

                time.sleep(1800)  # 30분 대기
                nowTime = int(datetime.datetime.now().strftime('%H%M'))

            state = self.kiwoom.GetConnectState()
            if state == 0:
                self.kiwoomLogin()  # 키움증권 로그인

            self.getQuantList(selectDateTime)  # 포트폴리오 작성
            self.runTrading(selectDateTime)  # 주식 매매

        elif select_com == '4':
            sendText = "DB를 초기화합니다."
            print(sendText)
            self.resetDB()  # DB 초기화

        else:
            sendText = "잘못 입력했습니다. 다시 입력하세요."
            print(sendText)

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
            sql_update_value = sql_update_value % (eps, bps, cfpf, sps, self.nowDateTime, row[0])
            cursor.execute(sql_update_value)

            if eps == 0 or bps == 0 or cfpf == 0 or sps == 0:
                print("일부 기업가치 지표 미추출로 0 처리")
            else:
                print("완료")

            delay_time = random.uniform(1, 5)
            time.sleep(random.uniform(0.5, delay_time))

        connect.close()

        print("종목별 투자정보 Insert : 완료")

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
        cursor.execute("SELECT Code, Price / CFPF FROM StockList WHERE Date = '%s';" % (self.nowDateTime))
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
        self.send_msg(sendText)
        connect = sqlite3.connect(self.DBPath, isolation_level=None)
        sqlite3.Connection
        cursor = connect.cursor()
        sql = "SELECT * FROM StockList WHERE Date = '%s';" % selectDateTime
        cursor.execute(sql)
        rows = cursor.fetchall()

        self.nowDateTime = selectDateTime + "_up_" + datetime.datetime.now().strftime('%m%d%H%M')

        for row in rows:
            result_df = self.kiwoom.block_request("opt10001", 종목코드=row[1], output="예수금상세현황", next=0)
            price = abs(int(result_df['현재가']))
            sql = "INSERT INTO StockList (Code, Name, MarketCap, Price, EPS, BPS, CFPF, SPS, Date) VALUES ('%s', '%s', %s, %s, %s, %s, %s, %s, '%s')" % (row[1], row[2], row[3], price, row[4], row[5], row[6], row[7], self.nowDateTime,)
            cursor.execute(sql)
            print("%s의 현재가를 업데이트 했습니다." % row[2])
            time.sleep(0.5) # 1초에 5번은 가능하나 0.25초일경우 176번 업데이트 후 차단되어 0.5초로 수정

        print("%s의 현재가 업데이트를 완료하였습니다." % selectDateTime)
        self.send_msg("%s의 현재가 업데이트를 완료하였습니다." % selectDateTime)

    def kiwoomLogin(self):
        print("키움 로그인 : ", end='')
        self.kiwoom = Kiwoom()
        self.kiwoom.CommConnect(block=True)  # 블록킹 처리
        userInfoAccno = self.kiwoom.GetLoginInfo("ACCNO")  # 전체 계좌 리스트
        self.myAccount = userInfoAccno[1]  # 주식거래 계좌로 선택
        print("완료")
        self.send_msg("키움 로그인 완료")
        print("계좌번호 : %s" % self.myAccount)

    def getQuantList(self, selectDateTime):
        print("주식 보유 현황 DB Update : ", end='')
        result_df = self.kiwoom.block_request("opw00001", 계좌번호=self.myAccount, 비밀번호="", 비밀번호입력매체구분="00", 조회구분=3, output="예수금상세현황", next=0)
        Deposit = int(result_df['d+2추정예수금'])

        result_df = self.kiwoom.block_request("opw00018", 계좌번호=self.myAccount, 비밀번호="", 비밀번호입력매체구분="00", 조회구분=1, output="계좌평가결과", next=0)
        TotalPurchase = int(result_df['총평가금액'][0])
        StockCount = int(result_df['조회건수'][0])

        myTotalAssets = Deposit + TotalPurchase
        print('완료 (가능 금액 %s 원, 주식수 %s)' % (myTotalAssets, StockCount))
        self.send_msg('보유 현황 (가능 금액 %s 원, 주식수 %s)' % (myTotalAssets, StockCount))

        connect = sqlite3.connect(self.DBPath, isolation_level=None)
        sqlite3.Connection
        cursor = connect.cursor()

        if StockCount == 0:
            print("종목 보유 없음.")
            self.send_msg("종목 보유 없음.")
        else:
            print("보유종목 현황 확인 : ", end='')
            result_df = self.kiwoom.block_request("opw00018", 계좌번호=self.myAccount, 비밀번호="", 비밀번호입력매체구분="00", 조회구분=1, output="계좌평가잔고개별합산", next=0)
            result_df.to_sql('TempStockHaving', connect, if_exists='replace')
            cursor.execute("INSERT INTO StockHaving (Code, Name, ProfitLoss, PurchasePrice, HavingCount, Price, Date) SELECT replace(종목번호, 'A', ''), 종목명, 평가손익, 매입가, 보유수량, 현재가, '%s' FROM TempStockHaving;" % (selectDateTime))
            print("완료(%s건)" % StockCount)
            self.send_msg("보유종목 : %s건" % StockCount)

        print("종목별 매수량, 매도량 확인 : 시작")
        self.send_msg("종목별 매수량, 매도량 확인 : 시작")
        # 선정된 종목에서 매수, 매도 확인

        cursor.execute("SELECT Code, Name FROM StockRank WHERE Date = '%s' ORDER BY RankTotal LIMIT %s;" % (selectDateTime, 20)) # 전역변수로 변경 가능
        rows = cursor.fetchall()

        for row in rows:
            time.sleep(1)
            Code = row[0]
            Name = row[1]

            df = self.kiwoom.block_request("opt10001", 종목코드=Code, output="주식기본정보", next=0)
            Price = abs(int(df['현재가'][0]))
            if Price == 0:
                print('거래정지(주가 0원)로 제외 : %s(%s)' % (df['종목명'], df['종목코드']))
                continue
            Quota = int(myTotalAssets / 20) # 전역변수로 변경 가능
            BuyingCount = Quota // Price

            # 보유중인 종목과 비교하여 적으면 매수, 많으면 매도
            cursor.execute("SELECT HavingCount FROM StockHaving WHERE Date = '%s' AND Code = '%s';" % (selectDateTime, Code))
            tempCount = cursor.fetchall()

            if tempCount == []:
                HavingCount = 0
            else:
                HavingCount = int(tempCount[0][0])

            if BuyingCount > HavingCount:
                Buy = BuyingCount - HavingCount
                Cell = 0
            elif BuyingCount < HavingCount:
                Buy = 0
                Cell = HavingCount - BuyingCount
            else:
                Buy = 0
                Cell = 0

            cursor.execute("INSERT INTO QuantList VALUES ('%s', '%s', %s, %s, %s, %s, %s, %s, '%s')" % (Code, Name, Price, Quota, BuyingCount, HavingCount, Buy, Cell, selectDateTime))

        # 보유중인 종목에서 List에 없는 종목 매도(List에 없으면 통과)
        cursor.execute("SELECT Code, Name, Price, HavingCount FROM StockHaving WHERE Date = '%s' AND Code NOT IN (SELECT Code FROM StockRank WHERE Date = '%s' ORDER BY RankTotal LIMIT %s);" % (selectDateTime, selectDateTime, 20))
        rows = cursor.fetchall()

        for row in rows:
            Code = row[0]
            Name = row[1]
            Price = row[2]
            Quota = 0
            BuyingCount = 0
            HavingCount = row[3]
            Buy = 0
            Cell = HavingCount
            Date = selectDateTime

            cursor.execute("INSERT INTO QuantList VALUES ('%s', '%s', %s, %s, %s, %s, %s, %s, '%s')" % (Code, Name, Price, Quota, BuyingCount, HavingCount, Buy, Cell, Date))

        connect.close()
        print("종목별 매수량, 매도량 확인 : 완료")

    def runTrading(self, selectDateTime):
        connect = sqlite3.connect(self.DBPath, isolation_level=None)
        sqlite3.Connection
        cursor = connect.cursor()

        # 대상종목 매도
        cursor.execute("SELECT Code, Cell, Name FROM QuantList WHERE Date = '%s' AND Cell > 0;" % (selectDateTime))
        rows = cursor.fetchall()

        for row in rows:
            time.sleep(2)
            self.kiwoom.SendOrder("매도거래", "1000", self.myAccount, 2, row[0], row[1], 0, "03", "")
            print("%s, %s주 매도 요청" % (row[2], row[1]))
            self.send_msg("%s, %s주 매도 요청" % (row[2], row[1]))

        df = self.kiwoom.block_request("opt10075", 계좌번호=self.myAccount, 전체종목구분=0, 매매구분=1, 종목코드="", 체결구분=1, output="미체결", next=0)
        while df.loc[0][1] != "":
            time.sleep(30)
            df = self.kiwoom.block_request("opt10075", 계좌번호=self.myAccount, 전체종목구분=0, 매매구분=1, 종목코드="", 체결구분=1, output="미체결", next=0)
            print("%s건 매도 미채결" % len(df))
            self.send_msg("%s건 매도 미채결" % len(df))

        print("%s건 매도 완료" %len(rows))
        self.send_msg("%s건 매도 완료" %len(rows))

        # 대상 종목 매수
        cursor.execute("SELECT Code, Buy, Name FROM QuantList WHERE Date = '%s' AND Buy > 0;" % (selectDateTime))    # 마지막에 , 확인
        rows = cursor.fetchall()

        for row in rows:
            time.sleep(2)
            self.kiwoom.SendOrder("매수거래", "1001", self.myAccount, 1, row[0], row[1], 0, "03", "")
            print("%s, %s주 매수 요청" % (row[2], row[1]))
            self.send_msg("%s, %s주 매수 요청" % (row[2], row[1]))

        df = self.kiwoom.block_request("opt10075", 계좌번호=self.myAccount, 전체종목구분=0, 매매구분=2, 종목코드="", 체결구분=1, output="미체결", next=0)
        while df.loc[0][1] != "":
            time.sleep(30)
            df = self.kiwoom.block_request("opt10075", 계좌번호=self.myAccount, 전체종목구분=0, 매매구분=2, 종목코드="", 체결구분=1, output="미체결", next=0)
            print("%s건 매수 미채결" % len(df))
            self.send_msg("%s건 매수 미채결" % len(df))

        print("%s건 매수 완료" %len(rows))
        self.send_msg("%s건 매수 완료" %len(rows))
        connect.close()

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
        # 텔레그램 토큰, 쳇ID 얻는법
        # 1. 텔레그램에서 BotFather 검색해서 들어감
        # 2. 아래의 "시작" 버튼 클릭
        # 3. "/newbot" 선택 후 봇이름 입력 (봇이름은 원하는 이름으로 입력)
        # 4. 봇 아이디 입력 (원하는 이름 입력, 마지막은 꼭 bot 추가) 하면 HTTP API 토큰 확인
        # 5. 텔레그램에서 봇 아이디 검색해서 들어간다음 "시작" 버튼 클릭
        # 6. 크롬 주소창에 "https://api.telegram.org/bot'HTTP API 토큰'/getUpdates" 입력하여 접속
        #    (홑따움표는 필요없음 예: https://api.telegram.org/bot1005626992:AAIIwBBByAJd9JLoZWFGKvsY9a4gO4IIIIg/getUpdates )
        # 7. 텔레그램에 아무 메세지 발송
        # 7. 브라우저 새로고침 한다음 표시되는 두번째 줄에 있는 ID 확인(쳇ID로 사용)

        apiToken = "HTTP API 토큰"
        chatId = "쳇ID"
        bot = telepot.Bot(apiToken)
        bot.sendMessage(chatId, msg)

if __name__ == "__main__":
    kiwoom = KiwoomPy()
