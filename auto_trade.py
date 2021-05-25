import win32com.client
import requests
from datetime import datetime
import time, ctypes
import pandas as pd
from TOCKEN import myToken

buy_amount = 0
# =======================================
# 크레온 플러스 공통 OBJECT
# =======================================
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

# =======================================
# Stock Bot
# =======================================



def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
                             headers={"Authorization": "Bearer " + token},
                             data={"channel": channel, "text": text}
                             )
    print(response)


def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    post_message(myToken, "#stock", strbuf)


def printlog(message, *args):
    """인자로 받은 문자열을 파이썬 셸에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)


# =======================================
# Check Creon Plus
# =======================================
def check_creon_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('check_creon_system() : admin user -> FAILED')
        return False

    # 연결 여부 체크
    if cpStatus.IsConnect == 0:
        printlog('check_creon_system() : connect to server -> FAILED')
        return False

    # # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    # if cpTradeUtil.TradeInit(0) != 0:
    #     printlog('check_creon_system() : init trade -> FAILED')
    #     return False
    # return True


# =======================================
# 종목 관련 코드
# =======================================
def get_current_price(code):
    """인자로 받은 종목의 현재가, 매수호가, 매도호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)  # 현재가
    item['ask'] = cpStock.GetHeaderValue(16)  # 매수호가
    item['bid'] = cpStock.GetHeaderValue(17)  # 매도호가
    return item['cur_price'], item['ask'], item['bid']


def current_price(code):
    # 현재가 객체 구하기
    objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    objStockMst.SetInputValue(0, code)  # 종목 코드 - 삼성전자
    objStockMst.BlockRequest()

    # 현재가 통신 및 통신 에러 처리
    rqStatus = objStockMst.GetDibStatus()
    rqRet = objStockMst.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()

    # 현재가 정보 조회
    code = objStockMst.GetHeaderValue(0)  # 종목코드
    name = objStockMst.GetHeaderValue(1)  # 종목명
    time = objStockMst.GetHeaderValue(4)  # 시간
    cprice = objStockMst.GetHeaderValue(11)  # 종가
    diff = objStockMst.GetHeaderValue(12)  # 대비
    open = objStockMst.GetHeaderValue(13)  # 시가
    high = objStockMst.GetHeaderValue(14)  # 고가
    low = objStockMst.GetHeaderValue(15)  # 저가
    offer = objStockMst.GetHeaderValue(16)  # 매도호가
    bid = objStockMst.GetHeaderValue(17)  # 매수호가
    vol = objStockMst.GetHeaderValue(18)  # 거래량
    vol_value = objStockMst.GetHeaderValue(19)  # 거래대금

    # 예상 체결관련 정보
    exFlag = objStockMst.GetHeaderValue(58)  # 예상체결가 구분 플래그
    exPrice = objStockMst.GetHeaderValue(55)  # 예상체결가
    exDiff = objStockMst.GetHeaderValue(56)  # 예상체결가 전일대비
    exVol = objStockMst.GetHeaderValue(57)  # 예상체결수량

    print("코드", code)
    print("이름", name)
    print("시간", time)
    print("종가", cprice)
    print("대비", diff)
    print("시가", open)
    print("고가", high)
    print("저가", low)
    print("매도호가", offer)
    print("매수호가", bid)
    print("거래량", vol)
    print("거래대금", vol_value)

    if exFlag == ord('0'):
        print("장 구분값: 동시호가와 장중 이외의 시간")
    elif exFlag == ord('1'):
        print("장 구분값: 동시호가 시간")
    elif exFlag == ord('2'):
        print("장 구분값: 장중 또는 장종료")

    print("예상체결가 대비 수량")
    print("예상체결가", exPrice)
    print("예상체결가 대비", exDiff)
    print("예상체결수량", exVol)


def get_ohlc(code, qty):
    """인자로 받은 종목의 OHLC 가격 정보를 qty 개수만큼 반환한다."""
    cpOhlc.SetInputValue(0, code)  # 종목코드
    cpOhlc.SetInputValue(1, ord('2'))  # 1:기간, 2:개수
    cpOhlc.SetInputValue(4, qty)  # 요청개수
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5])  # 0:날짜, 2~5:OHLC
    cpOhlc.SetInputValue(6, ord('D'))  # D:일단위
    cpOhlc.SetInputValue(9, ord('1'))  # 0:무수정주가, 1:수정주가
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)  # 3:수신개수
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count):
        index.append(cpOhlc.GetDataValue(0, i))
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
                     cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)])
    df = pd.DataFrame(rows, columns=columns, index=index)
    return df


def get_target_price(code):
    """매수 목표가를 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 10)
        if str_today == str(ohlc.iloc[0].name):
            today_open = ohlc.iloc[0].open
            lastday = ohlc.iloc[1]
        else:
            lastday = ohlc.iloc[0]
            today_open = lastday[3]
        lastday_high = lastday[1]
        lastday_low = lastday[2]
        target_price = today_open + (lastday_high - lastday_low) * 0.4
        return target_price
    except Exception as ex:
        dbgout("`get_target_price() -> exception! " + str(ex) + "`")
        return None


def get_moving_avg(code, window):
    """인자로 받은 종목에 대한 이동평균가격을 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 20)
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        closes = ohlc['close'].sort_index()
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as ex:
        dbgout('get_moving_avg(' + str(window) + ') -> exception! ' + str(ex))
        return None


def chart_data(code):
    # 차트 객체 구하기
    objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

    objStockChart.SetInputValue(0, code)  # 종목 코드 - 삼성전자
    objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
    objStockChart.SetInputValue(4, 100)  # 최근 100일 치
    objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
    objStockChart.SetInputValue(6, ord('D'))  # '차트 주가 - 일간 차트 요청
    objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
    objStockChart.BlockRequest()

    len = objStockChart.GetHeaderValue(3)

    print("날짜", "시가", "고가", "저가", "종가", "거래량")
    print("빼기빼기==============================================-")

    for i in range(len):
        day = objStockChart.GetDataValue(0, i)
        open = objStockChart.GetDataValue(1, i)
        high = objStockChart.GetDataValue(2, i)
        low = objStockChart.GetDataValue(3, i)
        close = objStockChart.GetDataValue(4, i)
        vol = objStockChart.GetDataValue(5, i)
        print(day, open, high, low, close, vol)


# =======================================
# 매수 매도
# =======================================
# 매수 주문
def buy_order(code, quantity, price):
    '''

    :param code: A로 시작하는 종목 코드
    :param quantity: 수량
    :param price: 가격
    :return: True/False
    '''
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # 주식상품 구분
    print(acc, accFlag[0])
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, "2")  # 2: 매수
    objStockOrder.SetInputValue(1, acc)  # 계좌번호
    objStockOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, code)  # 종목코드 - A003540 - 대신증권 종목
    objStockOrder.SetInputValue(4, quantity)  # 매수수량 10주
    objStockOrder.SetInputValue(5, price)  # 주문단가  - 14,100원
    objStockOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
    objStockOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

    # 매수 주문 요청
    objStockOrder.BlockRequest()

    rqStatus = objStockOrder.GetDibStatus()
    rqRet = objStockOrder.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        return False
    return True


def buy_etf(code):
    """인자로 받은 종목을 최유리 지정가 FOK 조건으로 매수한다."""
    try:
        global bought_list  # 함수 내에서 값 변경을 하기 위해 global로 지정
        if code in bought_list:  # 매수 완료 종목이면 더 이상 안 사도록 함수 종료
            # printlog('code:', code, 'in', bought_list)
            return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code)
        target_price = get_target_price(code)  # 매수 목표가
        ma5_price = get_moving_avg(code, 5)  # 5일 이동평균가
        ma10_price = get_moving_avg(code, 10)  # 10일 이동평균가
        buy_qty = 0  # 매수할 수량 초기화
        if ask_price > 0:  # 매수호가가 존재하면
            buy_qty = buy_amount // ask_price
        stock_name, stock_qty = get_stock_balance(code)  # 종목명과 보유수량 조회
        # printlog('bought_list:', bought_list, 'len(bought_list):',
        #    len(bought_list), 'target_buy_count:', target_buy_count)
        if current_price > target_price and current_price > ma5_price \
                and current_price > ma10_price:
            printlog(stock_name + '(' + str(code) + ') ' + str(buy_qty) +
                     'EA : ' + str(current_price) + ' meets the buy condition!`')
            cpTradeUtil.TradeInit()
            acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
            accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체,1:주식,2:선물/옵션
            # 최유리 FOK 매수 주문 설정
            cpOrder.SetInputValue(0, "2")  # 2: 매수
            cpOrder.SetInputValue(1, acc)  # 계좌번호
            cpOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
            cpOrder.SetInputValue(3, code)  # 종목코드
            cpOrder.SetInputValue(4, buy_qty)  # 매수할 수량
            cpOrder.SetInputValue(7, "2")  # 주문조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "12")  # 주문호가 1:보통, 3:시장가
            # 5:조건부, 12:최유리, 13:최우선
            # 매수 주문 요청
            ret = cpOrder.BlockRequest()
            printlog('최유리 FoK 매수 ->', stock_name, code, buy_qty, '->', ret)
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time / 1000)
                time.sleep(remain_time / 1000)
                return False
            time.sleep(2)
            printlog('현금주문 가능금액 :', buy_amount)
            stock_name, bought_qty = get_stock_balance(code)
            printlog('get_stock_balance :', stock_name, stock_qty)
            if bought_qty > 0:
                bought_list.append(code)
                dbgout("`buy_etf(" + str(stock_name) + ' : ' + str(code) +
                       ") -> " + str(bought_qty) + "EA bought!" + "`")
    except Exception as ex:
        dbgout("`buy_etf(" + str(code) + ") -> exception! " + str(ex) + "`")


def sell_all():
    """보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다."""
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
        while True:
            stocks = get_stock_balance('ALL')
            total_qty = 0
            for s in stocks:
                total_qty += s['qty']
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:
                    cpOrder.SetInputValue(0, "1")  # 1:매도, 2:매수
                    cpOrder.SetInputValue(1, acc)  # 계좌번호
                    cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
                    cpOrder.SetInputValue(3, s['code'])  # 종목코드
                    cpOrder.SetInputValue(4, s['qty'])  # 매도수량
                    cpOrder.SetInputValue(7, "1")  # 조건 0:기본, 1:IOC, 2:FOK
                    cpOrder.SetInputValue(8, "12")  # 호가 12:최유리, 13:최우선
                    # 최유리 IOC 매도 주문 요청
                    ret = cpOrder.BlockRequest()
                    printlog('최유리 IOC 매도', s['code'], s['name'], s['qty'],
                             '-> cpOrder.BlockRequest() -> returned', ret)
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        printlog('주의: 연속 주문 제한, 대기시간:', remain_time / 1000)
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        dbgout("sell_all() -> exception! " + str(ex))


# 매도 주문
def sell_order(code, quantity, price):
    '''

    :param code: A로 시작하는 종목 코드
    :param quantity: 수량
    :param price: 가격
    :return: True/False
    '''
    # 주식 매도 주문
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # 주식상품 구분
    print(acc, accFlag[0])
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, "1")  # 1: 매도
    objStockOrder.SetInputValue(1, acc)  # 계좌번호
    objStockOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, code)  # 종목코드 - A003540 - 대신증권 종목
    objStockOrder.SetInputValue(4, quantity)  # 매도수량 10주
    objStockOrder.SetInputValue(5, price)  # 주문단가  - 14,100원
    objStockOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
    objStockOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

    # 매도 주문 요청
    objStockOrder.BlockRequest()

    rqStatus = objStockOrder.GetDibStatus()
    rqRet = objStockOrder.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        return False
    return True


# 주문 초기화
def init_order():
    objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
    initCheck = objTrade.TradeInit(0)
    if initCheck != 0:
        print("주문 초기화 실패")
        exit()


# =======================================
# 계좌 관련
# =======================================

def get_current_cash():
    """증거금 100% 주문 가능 금액을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
    cpCash.SetInputValue(0, acc)  # 계좌번호
    cpCash.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpCash.BlockRequest()
    return cpCash.GetHeaderValue(9)  # 증거금 100% 주문 가능 금액


def get_stock_balance(code):
    """인자로 받은 종목의 종목명과 수량을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)  # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)  # 요청 건수(최대 50)
    cpBalance.BlockRequest()
    if code == 'ALL':
        dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
        dbgout('결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)))
        dbgout('평가금액: ' + str(cpBalance.GetHeaderValue(3)))
        dbgout('평가손익: ' + str(cpBalance.GetHeaderValue(4)))
        dbgout('종목수: ' + str(cpBalance.GetHeaderValue(7)))
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)  # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)  # 수량
        if code == 'ALL':
            dbgout(str(i + 1) + ' ' + stock_code + '(' + stock_name + ')'
                   + ':' + str(stock_qty))
            stocks.append({'code': stock_code, 'name': stock_name,
                           'qty': stock_qty})
        if stock_code == code:
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0


# =======================================
# Main
# =======================================
'''
    1) 거래량 : 1000만 이상
    2) 차트 : 
      2-1) 거래량 폭등 후 거래량 급감(25%) 그리고 음봉 출현했을 때
      2-2) 전날 거래량이 감소하였는가? 그리고 음봉이었는가 ?
      2-3) 현재 주가는 바닥인가?
      2-4) 바닥권이면서 단기고점은 아닌가 ?
      2-5) 단기 이평선 위에 자리잡고 있지 않은가? ( 역추세 ? )
      2-6) 지지를 잘 받고 있는가 ? 저항은 어디인가 ?
      2-7) 눌림목이라면 3, 8일선인가 5, 20일 선인가
    3) 재료
    
    기타 Tips)
     - 주식을 사야하는 지점 : 지지선이 깨지지 않을 때
     - 주식을 팔아야 하는 지점 : 저항선을 못 뚫고 내려앉았을 때
     -   ==> 반면에 저항선을 뚫었다면 매수
     - '거래량 급감 + 음봉' 나온 후 주가가 오를 확률 높다.
     - 거래량 감소폭 클 수록, 음봉 크기가 클 수록 담날 주가상승
    
'''


'''
    1. 
'''
if __name__ == '__main__':
    # check_creon_system()
    # A243890 : TIGER 200 에너지화학레버리지
    # print(get_current_cash())  # 400000 <- 매수 가능 현금
    # print(get_stock_balance('A243890'))  # ('TIGER 200에너지화학레버리지', 0)
    # print(get_current_price('A243890'))  # (22570, 22580, 22570)
    # print(get_ohlc('A243890', 2))
    # '''            open   high    low  close
    #     20210525  22560  22725  22475  22570
    #     20210524  22895  22895  22205  22255
    # '''
    # print(get_target_price('A243890'))  # 22836.0
    # print(get_moving_avg('A243890', 5))  # 22667.0
    # # print(buy_etf('A243890')) # 종목에 대해 최유리 지정가 FOK 매수
    # # print(sell_all())     # 보유 종목 최유리 지정가로



'''
        # 변동성 돌파 전략 + 이동평균선 5일 + 이동평균선 10일 3가지 전략 활용
    try:
        symbol_list = ['A243890', 'A243880', 'A122630', 'A305720',
                       'A138540']  # 거래하고자 하는 종목코드(다음금융 등에서 확인가능->링크에 종목코드 나와있음)
        bought_list = []  # 매수 완료된 종목 리스트
        target_buy_count = 5  # 매수할 종목 수
        buy_percent = 0.2  # 각각 매수할 종목을 몇퍼센트씩 구매할 것인지
        printlog('check_creon_system() :', check_creon_system())  # 크레온 접속 점검
        stocks = get_stock_balance('ALL')  # 보유한 모든 종목 조회
        total_cash = int(get_current_cash())  # 100% 증거금 주문 가능 금액 조회
        buy_amount = total_cash * buy_percent  # 종목별 주문 금액 계산
        printlog('100% 증거금 주문 가능 금액 :', total_cash)
        printlog('종목별 주문 비율 :', buy_percent)
        printlog('종목별 주문 금액 :', buy_amount)
        printlog('시작 시간 :', datetime.now().strftime('%m/%d %H:%M:%S'))
        soldout = False

        while True:
            t_now = datetime.now()
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_start = t_now.replace(hour=9, minute=5, second=0, microsecond=0)
            t_sell = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
            t_exit = t_now.replace(hour=15, minute=20, second=0, microsecond=0)
            today = datetime.today().weekday()
            if today == 5 or today == 6:  # 토요일이나 일요일이면 자동 종료
                printlog('Today is', 'Saturday.' if today == 5 else 'Sunday.')
                sys.exit(0)
            if t_9 < t_now < t_start and soldout == False:
                soldout = True
                # sell_all()
            if t_start < t_now < t_sell:  # AM 09:05 ~ PM 03:15 : 매수
                for sym in symbol_list:
                    if len(bought_list) < target_buy_count:
                        buy_etf(sym)
                        time.sleep(1)
                if t_now.minute == 30 and 0 <= t_now.second <= 5:
                    get_stock_balance('ALL')
                    time.sleep(5)
            if t_sell < t_now < t_exit:  # PM 03:15 ~ PM 03:20 : 일괄 매도
                sell_all()
                if sell_all() == True:
                    dbgout('`sell_all() returned True -> self-destructed!`')
                    sys.exit(0)
            if t_exit < t_now:  # PM 03:20 ~ :프로그램 종료
                dbgout('`self-destructed!`')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        dbgout('`main -> exception! ' + str(ex) + '`')
'''
