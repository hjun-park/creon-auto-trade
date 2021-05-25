import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes
import pandas as pd
import os

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

gExcelFile = '8092.xlsx'


def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False

    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    '''
    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False
    '''
    return True


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관
        self.diccode = {
            41: '주가 5MA 상향돌파',
            42: '주가 5MA 하향돌파',
            43: '거래량 5MA 상향돌파',
            44: '주가데드크로스(5MA < 20MA)',
            45: '주가골든크로스(5MA > 20MA)',
            81: '단기급락후 5MA 상향돌파',
            83: '눌림목재상승-20MA 지지'
        }

    def OnReceived(self):
        print(self.name)
        # 실시간 처리 - marketwatch : 특이 신호(차트, 외국인 순매수 등)
        if self.name == 'marketwatch':
            code = self.client.GetHeaderValue(0)
            name = g_objCodeMgr.CodeToName(code)
            cnt = self.client.GetHeaderValue(2)

            for i in range(cnt):
                item = {}
                newcancel = ''
                time = self.client.GetDataValue(0, i)
                h, m = divmod(time, 100)
                item['시간'] = '%02d:%02d' % (h, m)
                update = self.client.GetDataValue(1, i)
                item['코드'] = code
                item['종목명'] = name
                cate = self.client.GetDataValue(2, i)
                if update == ord('c'):
                    newcancel = '[취소]'
                if cate in self.diccode:
                    item['특이사항'] = newcancel + self.diccode[cate]
                else:
                    item['특이사항'] = newcancel + ''

                self.caller.listWatchData.insert(0, item)
                print(f'item : {item} ')

        # 실시간 처리 - marketnews : 뉴스 및 공시 정보
        elif self.name == 'marketnews2':
            item = {}
            update = self.client.GetHeaderValue(0)
            cont = ''
            if update == ord('D'):
                cont = '[삭제]'
            code = item['코드'] = self.client.GetHeaderValue(1)
            time = self.client.GetHeaderValue(2)
            h, m = divmod(time, 100)
            item['시간'] = '%02d:%02d' % (h, m)
            item['종목명'] = name = g_objCodeMgr.CodeToName(code)
            cate = self.client.GetHeaderValue(4)
            item['특이사항'] = cont + self.client.GetHeaderValue(5)
            print(item)
            self.caller.listWatchData.insert(0, item)


class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def Subscribe(self, var, caller):
        if self.bIsSB:
            self.Unsubscribe()

        if (len(var) > 0):
            self.obj.SetInputValue(0, var)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, caller)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False


# CpPBMarkeWatch:
class CpPBMarkeWatch(CpPublish):
    def __init__(self):
        super().__init__('marketwatch', 'CpSysDib.CpMarketWatchS')


#
# # CpPBMarkeWatch:
# class CpPB8092news(CpPublish):
#     def __init__(self):
#         super().__init__('marketnews', 'Dscbo1.CpSvr8092S')


# CpRpMarketWatch : 특징주 포착 통신
class CpRpMarketWatch:
    def __init__(self):
        self.objStockMst = win32com.client.Dispatch('CpSysDib.CpMarketWatch')
        self.objpbMarket = CpPBMarkeWatch()
        # self.objpbNews = CpPB8092news()
        return

    def Request(self, code):
        code_list = []
        self.objpbMarket.Unsubscribe()
        # self.objpbNews.Unsubscribe()

        self.objStockMst.SetInputValue(0, code)
        # 1: 종목 뉴스 2: 공시정보 10: 외국계 창구첫매수, 11:첫매도 12 외국인 순매수 13 순매도
        rqField = '43'
        self.objStockMst.SetInputValue(1, rqField)
        self.objStockMst.SetInputValue(2, 0)  # 시작 시간: 0 처음부터

        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print('통신상태', self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False

        cnt = self.objStockMst.GetHeaderValue(2)  # 수신 개수
        print(cnt)
        cnt = 2
        for i in range(cnt):
            item = {}

            time = self.objStockMst.GetDataValue(0, i)
            h, m = divmod(time, 100)
            item['시간'] = '%02d:%02d' % (h, m)
            item['코드'] = self.objStockMst.GetDataValue(1, i)
            item['종목명'] = g_objCodeMgr.CodeToName(item['코드'])
            cate = self.objStockMst.GetDataValue(3, i)
            item['특이사항'] = self.objStockMst.GetDataValue(4, i)
            print(item)
            code_list.append(item['코드'])

        return code_list


if __name__ == "__main__":
    cprq = CpRpMarketWatch()
    code_list = cprq.Request('*')
    print(code_list)
    # app = QApplication(sys.argv)
    # myWindow = MyWindow()
    # myWindow.show()
    # app.exec_()
