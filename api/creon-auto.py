import win32com.client

objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

# 연결 여부 체크
def check_connect():
    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    bConnect = objCpCybos.IsConnect
    if bConnect == 0:
        print("PLUS가 정상적으로 연결되지 않음. ")
        exit()


# 주문 초기화
def init_order():
    objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
    initCheck = objTrade.TradeInit(0)
    if initCheck != 0:
        print("주문 초기화 실패")
        exit()


# 매수 주문
def buy_order(code, quantity, price):
    '''

    :param code: A로 시작하는 종목 코드
    :param quantity: 수량
    :param price: 가격
    :return: True/False
    '''
    acc = objTrade.AccountNumber[0]         # 계좌번호
    accFlag = objTrade.GoodsList(acc, 1)    # 주식상품 구분
    print(acc, accFlag[0])
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, "2")         # 2: 매수
    objStockOrder.SetInputValue(1, acc)         # 계좌번호
    objStockOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, code)   # 종목코드 - A003540 - 대신증권 종목
    objStockOrder.SetInputValue(4, quantity)          # 매수수량 10주
    objStockOrder.SetInputValue(5, price)       # 주문단가  - 14,100원
    objStockOrder.SetInputValue(7, "0")         # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
    objStockOrder.SetInputValue(8, "01")        # 주문호가 구분코드 - 01: 보통

    # 매수 주문 요청
    objStockOrder.BlockRequest()

    rqStatus = objStockOrder.GetDibStatus()
    rqRet = objStockOrder.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        return False
    return True

# 매도 주문
def sell_order(code, quantity, price):
    '''

    :param code: A로 시작하는 종목 코드
    :param quantity: 수량
    :param price: 가격
    :return: True/False
    '''
    # 주식 매도 주문
    acc = objTrade.AccountNumber[0]  # 계좌번호
    accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분
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


#=========================================================================
# =========================================================================
# =========================================================================

#  class CpPublish:
#     def __init__(self, name, serviceID):
#         self.name = name
#         self.obj = win32com.client.Dispatch(serviceID)
#         self.bIsSB = False
#
# def CP_PB_NAV():
#
#
#
#
# # 실시간 nav/iiv 수신
# class CP_PB_NAV(CpPublish):
#     def __init__(self):
#         super().__init__('nav', 'CpSysDib.CpSvrNew7244S')
#
#
# def Request(self, code):
#     self.code = code
#     self.navlist = []
#
#     #######################################
#     # NAV/IIV 시간대별 리스트 요청
#     if (self.rqObj):
#         self.rqObj = None
#
#     self.rqObj = CP_ETF_NAV()
#     self.rqObj.Request(code, self)
#
# # 실시간 현재가 수신
# class CP_PB_CUR(CpPublish):
#     def __init__(self):
#         super().__init__('stockcur', 'DsCbo1.StockCur')
#
# class CP_ETF_NAV:
#     def __init__(self):
#         self.objRq = None
#         self.objReply = None
#         self.objname1 = 'Dscbo1.Cpsvr7244'
#         self.objname2 = 'Dscbo1.Cpsvr7718'
#         self.navlist = []
#         self.code = ''
#         self.caller = None
#
# def OnReply(self, objRq):
#     cnt = objRq.GetHeaderValue(0)
#     print('조회 개수', cnt)
#
#     for i in range(cnt):
#         item = {}
#
#         item['시간'] = objRq.GetDataValue(0, i)
#         item['현재가'] = objRq.GetDataValue(1, i)
#         item['대비'] = objRq.GetDataValue(3, i)
#         item['거래량'] = objRq.GetDataValue(5, i)
#         if (self.code[0] == 'A'):
#             item['NAV대비'] = objRq.GetDataValue(4, i)
#             item['NAV'] = objRq.GetDataValue(6, i)
#         else:
#             item['IIV대비'] = objRq.GetDataValue(4, i)
#             item['IIV'] = objRq.GetDataValue(6, i)
#         item['추적오차'] = objRq.GetDataValue(8, i)
#         item['괴리율'] = objRq.GetDataValue(9, i)
#         item['해당ETF지수'] = objRq.GetDataValue(10, i)
#         item['지수대비'] = objRq.GetDataValue(11, i)
#         print(item)
#         # self.navlist.append(item)
#         # self.caller.OnReply(self.navlist)
#

#=========================================================================
# =========================================================================
# =========================================================================

