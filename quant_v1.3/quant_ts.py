# 버전업(1.2) - 정보획득(quant_selectitem.py)과 리벨런싱(quant_ts) 프로그램 분리
# 리벨런싱(quant_ts) 프로그램은 출근 후 텔레그램으로 주식계좌와 주식 정보를 텔레그램으로 원하는 번호로 선택하여 진행

import time
import sqlite3
import telepot # 설치필요
from pykiwoom.kiwoom import *   # 설치필요

DBPath = 'quantDB.db'  # DB 파일위치

with open('telepot.txt', 'r') as file:
    keys = file.readlines()
    apiToken = keys[0].strip()
    myId = keys[1].strip()
bot = telepot.Bot(apiToken)

def selectNum(msg):
    chk = True
    while chk:
        if msg == 0:
            messageText = '남길 금액은 얼마인가요?'
        else:
            messageText = ''
            for i in range(1, len(msg)+1):
                messageText = messageText + '%s: %s \n' % (i, msg[i-1])

        print(messageText)
        bot.sendMessage(myId, messageText)
        last_mes_all = bot.getUpdates()
        if last_mes_all == [] :
            last_mesTime = 0
        else :
            last_mes = bot.getUpdates()[-1]['message']
            last_mesTime = last_mes['date']

        last_mes_all = bot.getUpdates()
        while last_mes_all == [] :
            time.sleep(3)
            last_mes_all = bot.getUpdates()

        now_mes = bot.getUpdates()[-1]['message']
        while last_mesTime >= now_mes['date']:
            time.sleep(3)
            now_mes = bot.getUpdates()[-1]['message']

        now_mesText = now_mes['text']
        if msg == 0:
            now_mesText = int(now_mesText)
            chk = False
        elif int(now_mesText) in range(1, len(msg) + 1):
            now_mesText = int(now_mesText) - 1
            chk = False

    return now_mesText

def getQuantList(selectDateTime, myAccount, cache):
    print("주식 보유 현황 DB Update : ", end='')
    result_df = kiwoom.block_request("opw00001", 계좌번호=myAccount, 비밀번호="", 비밀번호입력매체구분="00", 조회구분=3, output="예수금상세현황", next=0)
    while result_df['d+2추정예수금'][0] == '':
        result_df = kiwoom.block_request("opw00001", 계좌번호=myAccount, 비밀번호="", 비밀번호입력매체구분="00", 조회구분=3, output="예수금상세현황", next=0)
        time.sleep(3.6)

    Deposit = int(result_df['d+2추정예수금'][0])

    result_df = kiwoom.block_request("opw00018", 계좌번호=myAccount, 비밀번호="", 비밀번호입력매체구분="00", 조회구분=1, output="계좌평가결과", next=0)
    while result_df['총평가금액'][0] == '' or result_df['조회건수'][0] == '':
        result_df = kiwoom.block_request("opw00018", 계좌번호=myAccount, 비밀번호="", 비밀번호입력매체구분="00", 조회구분=1, output="계좌평가결과", next=0)
        time.sleep(3.6)

    TotalPurchase = int(result_df['총평가금액'][0])
    StockCount = int(result_df['조회건수'][0])

    myTotalAssets = (Deposit + TotalPurchase) - cache
    print('완료 (가능 금액 %s 원, 주식수 %s)' % (myTotalAssets, StockCount))
    bot.sendMessage(myId, '보유 현황 (가능 금액 %s 원, 주식수 %s)' % (myTotalAssets, StockCount))

    connect = sqlite3.connect(DBPath, isolation_level=None)
    sqlite3.Connection
    cursor = connect.cursor()
    cursor.execute("DELETE FROM StockHaving;")
    cursor.execute("DELETE FROM QuantList;")

    if StockCount == 0:
        print("종목 보유 없음.")
        bot.sendMessage(myId, "종목 보유 없음.")
    else:
        print("보유종목 현황 확인 : ", end='')
        result_df = kiwoom.block_request("opw00018", 계좌번호=myAccount, 비밀번호="", 비밀번호입력매체구분="00", 조회구분=1, output="계좌평가잔고개별합산", next=0)
        result_df.to_sql('TempStockHaving', connect, if_exists='replace')
        cursor.execute("INSERT INTO StockHaving (Code, Name, ProfitLoss, PurchasePrice, HavingCount, Price, Date) SELECT replace(종목번호, 'A', ''), 종목명, 평가손익, 매입가, 보유수량, 현재가, '%s' FROM TempStockHaving;" % (selectDateTime))
        print("완료(%s건)" % StockCount)
        bot.sendMessage(myId, "보유종목 : %s건" % StockCount)

    print("종목별 매수량, 매도량 확인 : 시작")
    bot.sendMessage(myId, "종목별 매수량, 매도량 확인 : 시작")

    # 선정된 종목에서 매수, 매도 확인
    cursor.execute("SELECT Code, Name FROM StockRank WHERE Date = '%s' ORDER BY RankTotal LIMIT %s;" % (selectDateTime, 20))
    rows = cursor.fetchall()

    for row in rows:
        time.sleep(1)
        Code = row[0]
        Name = row[1]

        result_df = kiwoom.block_request("opt10001", 종목코드=Code, output="주식기본정보", next=0)
        while result_df['종목코드'][0] != row[0]:
            result_df = kiwoom.block_request("opt10001", 종목코드=Code, output="주식기본정보", next=0)
            time.sleep(3.6)
        Price = abs(int(result_df['현재가'][0]))
        if Price == 0:
            print('거래정지(주가 0원)로 제외 : %s(%s)' % (result_df['종목명'], result_df['종목코드']))
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

def runTrading(selectDateTime, myAccount):
    connect = sqlite3.connect(DBPath, isolation_level=None)
    sqlite3.Connection
    cursor = connect.cursor()

    # 대상종목 매도
    cursor.execute("SELECT Code, Cell, Name FROM QuantList WHERE Date = '%s' AND Cell > 0;" % (selectDateTime))
    rows = cursor.fetchall()

    for row in rows:
        time.sleep(2)
        kiwoom.SendOrder("매도거래", "1000", myAccount, 2, row[0], row[1], 0, "03", "")
        print("%s, %s주 매도 요청" % (row[2], row[1]))
        bot.sendMessage(myId, "%s, %s주 매도 요청" % (row[2], row[1]))

    df = kiwoom.block_request("opt10075", 계좌번호=myAccount, 전체종목구분=0, 매매구분=1, 종목코드="", 체결구분=1, output="미체결", next=0)
    while len(df) > 1 :#  df.loc[0][1] != "":
        time.sleep(30)
        df = kiwoom.block_request("opt10075", 계좌번호=myAccount, 전체종목구분=0, 매매구분=1, 종목코드="", 체결구분=1, output="미체결", next=0)
        print("%s건 매도 미채결" % len(df))
        bot.sendMessage(myId, "%s건 매도 미채결" % len(df))

    print("%s건 매도 완료" %len(rows))
    bot.sendMessage(myId, "%s건 매도 완료" %len(rows))

    # 대상 종목 매수
    cursor.execute("SELECT Code, Buy, Name FROM QuantList WHERE Date = '%s' AND Buy > 0;" % (selectDateTime))    # 마지막에 , 확인
    rows = cursor.fetchall()

    for row in rows:
        time.sleep(2)
        kiwoom.SendOrder("매수거래", "1001", myAccount, 1, row[0], row[1], 0, "03", "")
        print("%s, %s주 매수 요청" % (row[2], row[1]))
        bot.sendMessage(myId, "%s, %s주 매수 요청" % (row[2], row[1]))

    df = kiwoom.block_request("opt10075", 계좌번호=myAccount, 전체종목구분=0, 매매구분=2, 종목코드="", 체결구분=1, output="미체결", next=0)
    while len(df) > 1 :#  df.loc[0][1] != "":
        time.sleep(30)
        df = kiwoom.block_request("opt10075", 계좌번호=myAccount, 전체종목구분=0, 매매구분=2, 종목코드="", 체결구분=1, output="미체결", next=0)
        print("%s건 매수 미채결" % len(df))
        bot.sendMessage(myId, "%s건 매수 미채결" % len(df))

    print("%s건 매수 완료" %len(rows))
    bot.sendMessage(myId, "%s건 매수 완료" %len(rows))
    connect.close()

def updateNowPrice(selectDateTime):
    print(sendText)
    connect = sqlite3.connect(DBPath, isolation_level=None)
    sqlite3.Connection
    cursor = connect.cursor()
    sql = "SELECT ID, Code, Name FROM StockList WHERE Date = '%s';" % selectDateTime
    cursor.execute(sql)
    rows = cursor.fetchall()

    for row in rows:
        result_df = kiwoom.block_request("opt10001", 종목코드=row[1], output="주식기본정보", next=0)
        while result_df['종목코드'][0] != row[1]:
            result_df = kiwoom.block_request("opt10001", 종목코드=row[1], output="주식기본정보", next=0)
            time.sleep(3.6)
        price = abs(int(result_df['현재가']))
        # sql = "INSERT INTO StockList (Code, Name, MarketCap, Price, EPS, BPS, CFPS, SPS, Date) VALUES ('%s', '%s', %s, %s, %s, %s, %s, %s, '%s')" % (row[1], row[2], row[3], price, row[4], row[5], row[6], row[7], nowDateTime,)
        sql = "UPDATE StockList SET Price = %s WHERE ID = %s AND Date = %s;" % (price, row[0], selectDateTime)
        cursor.execute(sql)
        print("%s의 현재가를 업데이트 했습니다." % row[2])
        time.sleep(0.7) # 1초에 5번은 가능하나 0.25초일경우 176번 업데이트 후 차단되어 0.5초로 수정

    print("%s의 현재가 업데이트를 완료하였습니다." % selectDateTime)

def getStockInfo(selectDateTime):
    connect = sqlite3.connect(DBPath, isolation_level=None)
    sqlite3.Connection
    cursor = connect.cursor()

    print("RER 순위 추가 : ", end='')
    cursor.execute("SELECT Code, Name, Price / EPS FROM StockList WHERE Date = '%s';" % (selectDateTime))
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
    cursor.execute("SELECT Code, Price / BPS FROM StockList WHERE Date = '%s';" % (selectDateTime))
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
    cursor.execute("SELECT Code, Price / CFPS FROM StockList WHERE Date = '%s';" % (selectDateTime))
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
    cursor.execute("SELECT Code, Price / SPS FROM StockList WHERE Date = '%s';" % (selectDateTime))
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
    result_df['Date'] = selectDateTime
    result_df.to_sql('StockRank', connect, if_exists='replace')
    connect.close()

    print("완료")
    print("정보입력 완료")

nowTime = int(datetime.datetime.now().strftime('%H%M'))
print("현재시간 : %s" % datetime.datetime.now().strftime('%H:%M'))

sendText = "리밸런싱을 시작 합니다."
print(sendText)
bot.sendMessage(myId, sendText)

# 키움 로그인
print("키움 로그인 : ", end='')
kiwoom = Kiwoom()  # 키움 API사용
kiwoom.CommConnect(block=True)  # 블록킹 처리
print("완료")
bot.sendMessage(myId, "키움 로그인 완료")

# 키움 계좌 선택
tmpUserInfoAccno = kiwoom.GetLoginInfo("ACCNO")  # 전체 계좌 리스트
userInfoAccno = []

for r in tmpUserInfoAccno:
    if r[8:] == '10':
        userInfoAccno.append(r)

print(userInfoAccno)

acountnum = selectNum(userInfoAccno)
myAccount = userInfoAccno[acountnum]
print('선택 계좌번호 : %s' % myAccount)
bot.sendMessage(myId, '선택 계좌번호 : %s' % myAccount)

# 남길 금액 입력
cache = selectNum(0)

# 주식 정보 선택
connect = sqlite3.connect(DBPath, isolation_level=None)
sqlite3.Connection
cursor = connect.cursor()
sql = "SELECT DISTINCT Date FROM StockList;"
cursor.execute(sql)
dateList = cursor.fetchall()

selectDateTime = dateList[0][0]
print('정보 데이터 : %s' % selectDateTime)
bot.sendMessage(myId, '정보 데이터 : %s' % selectDateTime)

print('현재가를 업데이트 합니다.')
bot.sendMessage(myId, '현재가를 업데이트 합니다.')
updateNowPrice(selectDateTime)  # 현재가 업데이트

print('퀀트 종목을 선정합니다.')
bot.sendMessage(myId, '퀀트 종목을 선정합니다.')
getStockInfo(selectDateTime)  # 지표별 값 및 종합 순위 결정

# 포트폴리오 작성
getQuantList(selectDateTime, myAccount, cache)

# 주식 매매
runTrading(selectDateTime, myAccount)

sendText = "리밸런싱을 완료했습니다."
print(sendText)
bot.sendMessage(myId, sendText)
