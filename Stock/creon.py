import time
import win32com.client
import pandas as pd

# 크레온 클래스는 win32com 패키지의 모튤을 이용합니다. Dispatch 클래스의 인스턴스로 대신증권 크레온 모듈을 사용할 수 있습니다.
# 여기서는 CpUtil.CpCodeMgr, obj_CpCybos, obj_StockChart 모듈을 사용합니다.

#각 모듈의 상세정보는 웹페이스를 통해 알 수 있습니다.
class Creon:
    def __init__(self):
        self.obj_CpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        self.obj_CpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
        self.obj_StockChart = win32com.client.Dispatch('CpSysDib.StockChart')
    #이 함수의 파라미터로 가져올 주식 종목의 코드(code), 가져올 데이터의 시작일(date_from), 가져올 데이터의 종료일(date_to)를저장
    def creon_7400_주식차트조회(self,code,date_from,date_to):
        #대신증권 크레온 프로그램에 연결되어 있는지 확인합니다.
        b_connected = self.obj_CpCybos.IsConnect
        if b_connected == 0:
            print("연결실패")
            return None
        #연결되어 있으면 가져올 필드 키 (field key)를 정해 줍니다.
        list_field_key = [0,1,2,3,4,5,8]
        #이 키에 대응하는 값입니다.
        list_field_name = ['date','time','open','high','low','close','volume']
        dict_chart = {name:[] for name in list_field_name}

        #API를 호출하기 위해 입력값들을 설정해 주는 부분입니다.

        self.obj_StockChart.SetInputValue(0,'A'+code)
        #기간 or 개수 설정
        self.obj_StockChart.SetInputValue(1, ord('1')) # 0:개수, 1:기간
        self.obj_StockChart.SetInputValue(2, date_to) # 종료일
        self.obj_StockChart.SetInputValue(3, date_from) # 시작일
        self.obj_StockChart.SetInputValue(5, list_field_key) # 필드
        self.obj_StockChart.SetInputValue(6, ord('D')) # 'D', 'W' , 'M' , 'm' , 'T'
        self.obj_StockChart.BlockRequest() #입력한 설정에 따라 데이터를 요청합니다.

        # 요청 결과 상태를 받아옵니다. 이 값이 0이 아니라면 이상이 있음을 의미합니다.
        status = self.obj_StockChart.GetDibStatus()
        msg = self.obj_StockChart.GetDibMsg1()
        print("통신상태:{} {}".format(status,msg))
        if status != 0:
            return None
        # 결과 출력물의 개수를 확인합니다.
        cnt = self.obj_StockChart.GetHeaderValue(3) # 수신개수
        for i in range(cnt):
            dict_item = (
                {name: self.obj_StockChart.getDataValue(pos,i)
                 for pos, name in zip(range(len(list_field_name)),list_field_name)}
            )
            for k, v in dict_item.items():
                dict_chart[k].append(v) # 받아온 값을 dict_chart 딕셔너리에 추가해 줍니다.


        print("차트 : {} {}".format(cnt,dict_chart))
        # 이렇게 구성된 dict_chart 딕셔너리를 pandas DataFrame 객체로 만들어서 반환합니다.
        return pd.DataFrame(dict_chart,columns =list_field_name);

if __name__ == '__main__':
    creon = Creon()
    print(creon.creon_7400_주식차트조회('035420',20150101,20171201))
