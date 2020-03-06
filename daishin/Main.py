import win32com.client
import pandas as pd
import numpy as np
class _Network:
    def __init__(self):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            exit()

class Code_Load:
    stock_list = []

    # 0 /A000020/ 1/ 8960 / 동화약품
    def __init__(self):
        # 종목코드 리스트 구하기
        objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        codeList = objCpCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = objCpCodeMgr.GetStockListByMarket(2)  # 코스닥



        for i, code in enumerate(codeList):
            secondCode = objCpCodeMgr.GetStockSectionKind(code)
            name = objCpCodeMgr.CodeToName(code)
            stdPrice = objCpCodeMgr.GetStockStdPrice(code)
            status = objCpCodeMgr.GetStockStatusKind(code)
            control = objCpCodeMgr.GetStockControlKind(code)
            if((secondCode == 1) and (status < 1) and (control < 3)):
                self.stock_list.append([i, code, secondCode, stdPrice, name])
        for i, code in enumerate(codeList2):
            secondCode = objCpCodeMgr.GetStockSectionKind(code)
            name = objCpCodeMgr.CodeToName(code)
            stdPrice = objCpCodeMgr.GetStockStdPrice(code)
            status = objCpCodeMgr.GetStockStatusKind(code)
            control = objCpCodeMgr.GetStockControlKind(code)
            if ((secondCode == 1) and (status < 1) and (control < 3)):
                self.stock_list.append([i, code, secondCode, stdPrice, name])

class CurrentPriceEx:
    price_value = []

    def __init__(self, code):
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

        self.price_value = [code, name, time, cprice, diff, open,
                            high, low, offer, bid, vol, vol_value,
                            exFlag, exPrice, exDiff, exVol]



class StockWeek:
    def __init__(self, code):
        # 최초 데이터 요청
        objStockWeek = self._setup(code)
        ret = self.ReqeustData(objStockWeek)
        if ret == False:
            exit()

        # 연속 데이터 요청
        # 예제는 5번만 연속 통신 하도록 함.
        NextCount = 1
        while objStockWeek.Continue:  # 연속 조회처리
            NextCount += 1;
            if (NextCount > 5):
                break
            ret = self.ReqeustData(objStockWeek)
            if ret == False:
                exit()

    def _setup(self, code):
        objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
        objStockWeek.SetInputValue(0, code)  # 종목 코드 - 삼성전자
        return objStockWeek

    def ReqeustData(obj):
        # 데이터 요청
        obj.BlockRequest()

        # 통신 결과 확인
        rqStatus = obj.GetDibStatus()
        rqRet = obj.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        # 일자별 정보 데이터 처리
        count = obj.GetHeaderValue(1)  # 데이터 개수
        for i in range(count):
            date = obj.GetDataValue(0, i)  # 일자
            open = obj.GetDataValue(1, i)  # 시가
            high = obj.GetDataValue(2, i)  # 고가
            low = obj.GetDataValue(3, i)  # 저가
            close = obj.GetDataValue(4, i)  # 종가
            diff = obj.GetDataValue(5, i)  # 종가
            vol = obj.GetDataValue(6, i)  # 종가
            print(date, open, high, low, close, diff, vol)

        return True



if __name__ == '__main__':
    codeList = Code_Load()
    code = np.array(codeList.stock_list)

    code = code[:,1:]

    df_code = pd.DataFrame(code,columns=['종목코드', 'secondCode', '가격', '회사명'])
    df_code.to_excel(r'.\종목코드.xlsx')



