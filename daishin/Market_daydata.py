import datetime
import sys
from PyQt5.QtWidgets import *
import pandas as pd
import win32com.client
import ctypes

################################################
# PLUS 공통 OBJECT
from daishin import setting

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


################################################


# PLUS 실행 기본 체크 함수
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

    # # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False

    return True

FORMAT_DATETIME = "%Y-%m-%d"
today = setting.get_today_str()
today_1 = setting.get_today_str_1()

class CpMarketEye:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.RpFiledIndex = 0

    def Request(self, codes, dataInfo, flag):
        # 0: 종목코드 4: 현재가 20: 상장주식수
        rqField = [0,2,3,4, 5, 6, 7, 17, 23,]  # 요청 필드

        self.objRq.SetInputValue(0, rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        print("통신상태", rqStatus, self.objRq.GetDibMsg1())
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(2)

        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            contrast_sigh = self.objRq.GetDataValue(1, i) # 대비부호
            the_day_before = self.objRq.GetDataValue(2,i) # 전일대비
            day = self.objRq.GetDataValue(3, i)  # 현재가
            open = self.objRq.GetDataValue(4, i)  # 시가
            high = self.objRq.GetDataValue(5, i) # 고가
            low = self.objRq.GetDataValue(6, i)  # 저가
            name = self.objRq.GetDataValue(7, i)  # 저가
            close = self.objRq.GetDataValue(8,i) #전일 종가

            if flag == 1:
                # key(종목코드) = tuple(상장주식수, 시가총액)
                dataInfo[code] = {'종목': code, '종목명':name ,'날짜': today_1, '현재가':day, '시가':open, '고가':high, '저가':low, '전일종가':close,
                                  '대비부호':chr(contrast_sigh), '전일대비':the_day_before, '목록':'코스피'}

            if flag == 2:
                # key(종목코드) = tuple(상장주식수, 시가총액)
                dataInfo[code] = {'종목': code, '종목명':name ,'날짜': today_1, '현재가':day, '시가':open, '고가':high, '저가':low, '전일종가':close,
                                  '대비부호':chr(contrast_sigh), '전일대비':the_day_before, '목록':'코스닥'}


        return True


class CMarketTotal():
    def __init__(self):
        self.dataInfo = {}

    def GetAllMarketTotal(self):



        codeList = g_objCodeMgr.GetStockListByMarket(1)  # 거래소
        print('전 종목 코드, 거래소 %d' % len(codeList))

        objMarket = CpMarketEye()
        rqCodeList = []
        rqCodeList.append('U001')
        rqCodeList.append('U201')
        for i, code in enumerate(codeList):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                objMarket.Request(rqCodeList, self.dataInfo, 1)
                rqCodeList = []
                continue
        # end of for
        if len(rqCodeList) > 0:
            objMarket.Request(rqCodeList, self.dataInfo,1 )

        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # 코스닥
        for i, code in enumerate(codeList2):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                objMarket.Request(rqCodeList, self.dataInfo, 2)
                rqCodeList = []
                continue

        if len(rqCodeList) > 0:
            objMarket.Request(rqCodeList, self.dataInfo,2)



    def PrintMarketTotal(self):
        print(self.dataInfo)
        data_df = pd.DataFrame(self.dataInfo).T

        print(data_df)
        data_df.to_excel(setting.DATAPATH.format(today, 'new_low', today))

        print('finish')




if __name__ == "__main__":
    objMarketTotal = CMarketTotal()
    objMarketTotal.GetAllMarketTotal()
    objMarketTotal.PrintMarketTotal()