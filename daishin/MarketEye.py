import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes
import pandas as pd
import numpy as np
import time
import datetime

################################################
# PLUS 공통 OBJECT
from daishin import setting

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

################################################

FORMAT_DATETIME = "%Y-%m-%d"
today = setting.get_today_str()

DATAPATH_1 = setting.DATAPATH.format(today, 'data', today)
DATAPATH_2 = setting.DATAPATH.format(today, 'data_small', today)

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


class CpMarketEye:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.RpFiledIndex = 0

    def Request(self, codes, dataInfo):
        # 0: 종목코드 4: 현재가 20: 상장주식수
        # 갯수: 58
        rqField = [0, 2, 3, 4, 5, 10, 17, 20, 67, 68 ,69 ,70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106
                   ,107, 108, 109, 110,111 ,118, 120, 123, 124, 125, 141]  # 요청 필드

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
            # 갯수:58
            temp0 = self.objRq.GetDataValue(0, i)   # 종목코드
            temp17 = self.objRq.GetDataValue(6, i)  # 종목명
            temp67 = self.objRq.GetDataValue(8, i)  # PER
            temp4 = self.objRq.GetDataValue(3, i)   # 현재가
            temp5 = self.objRq.GetDataValue(4, i)   # 시가
            temp10 = self.objRq.GetDataValue(5, i)  # 거래량
            temp20 = self.objRq.GetDataValue(7, i)  # 총상장주식수
            temp68 = self.objRq.GetDataValue(9, i)  # 시간외매수잔량
            temp69 = self.objRq.GetDataValue(10, i)  # 시간외매도잔량
            temp70 = self.objRq.GetDataValue(11, i)  # EPS
            temp71 = self.objRq.GetDataValue(12, i)  # 자본금
            temp72 = self.objRq.GetDataValue(13, i)  # 액면가
            temp73 = self.objRq.GetDataValue(14, i)  # 배당률
            temp74 = self.objRq.GetDataValue(15, i)  # 배당수익률
            temp75 = self.objRq.GetDataValue(16, i)  # 부채비율
            temp76 = self.objRq.GetDataValue(17, i)  # 유보율
            temp77 = self.objRq.GetDataValue(18, i)  # 자기자본이익률
            temp78 = self.objRq.GetDataValue(19, i)  # 매출액증가율
            temp79 = self.objRq.GetDataValue(20, i)  # 경상이익증가율
            temp80 = self.objRq.GetDataValue(21, i)  # 순이익증가율
            temp81 = self.objRq.GetDataValue(22, i)  # 투자심리
            temp82 = self.objRq.GetDataValue(23, i)  # VR
            temp83 = self.objRq.GetDataValue(24, i)  # 5일 회전율
            temp84 = self.objRq.GetDataValue(25, i)  # 4일 종가합
            temp85 = self.objRq.GetDataValue(26, i)  # 9일 종가합
            temp86 = self.objRq.GetDataValue(27, i)  # 매출액
            temp87 = self.objRq.GetDataValue(28, i)  # 경상이익
            temp88 = self.objRq.GetDataValue(29, i)  # 당기순이익
            temp89 = self.objRq.GetDataValue(30, i)  # BPS
            temp90 = self.objRq.GetDataValue(31, i)  # 영업이익증가율
            temp91 = self.objRq.GetDataValue(32, i)  # 영업이익
            temp92 = self.objRq.GetDataValue(33, i)  # 매출액영업이익률
            temp93 = self.objRq.GetDataValue(34, i)  # 매출액경상이익률
            temp94 = self.objRq.GetDataValue(35, i)  # 이자보상비율
            temp95 = self.objRq.GetDataValue(36, i)  # 결산년월
            temp96 = self.objRq.GetDataValue(37, i)  # 분기BPS
            temp97 = self.objRq.GetDataValue(38, i)  # 분기매출액증가율
            temp98 = self.objRq.GetDataValue(39, i)  # 분기영업이액증가율
            temp99 = self.objRq.GetDataValue(40, i)  # 분기경상이익증가율
            temp100 = self.objRq.GetDataValue(41, i)  # 분기순이익증가율
            temp101 = self.objRq.GetDataValue(42, i)  # 분기매출액
            temp102 = self.objRq.GetDataValue(43, i)  # 분기영업이익
            temp103 = self.objRq.GetDataValue(44, i)  # 분기경상이익
            temp104 = self.objRq.GetDataValue(45, i)  # 분기당기순이익
            temp105 = self.objRq.GetDataValue(46, i)  # 분개매출액영업이익률
            temp106 = self.objRq.GetDataValue(47, i)  # 분기매출액경상이익률
            temp107 = self.objRq.GetDataValue(48, i)  # 분기ROE
            temp108 = self.objRq.GetDataValue(49, i)  # 분기이자보상비율
            temp109 = self.objRq.GetDataValue(50, i)  # 분기유보율
            temp110 = self.objRq.GetDataValue(51, i)  # 분기부채비율
            temp111 = self.objRq.GetDataValue(52, i)  # 분기결산년월
            temp118 = self.objRq.GetDataValue(53, i)  # 당일외국인순매수
            temp120 = self.objRq.GetDataValue(54, i)  # 당일기관순매수
            temp123 = self.objRq.GetDataValue(55, i)  # SPS
            temp124 = self.objRq.GetDataValue(56, i)  # CFPS
            temp125 = self.objRq.GetDataValue(57, i)  # EBITDA
            temp141 = self.objRq.GetDataValue(58, i)  # ELW 손익분기율



            if temp124 == 0:
                PCR = np.nan
            else:
                PCR = temp4 / temp124
                if PCR < 0:
                    PCR = np.nan

            if temp123 == 0:
                PSR = np.nan
            else:
                PSR = temp4 / temp123
                if PSR < 0:
                    PSR = np.nan

            if temp96 == 0:
                PBR = np.nan
            else:
                PBR = temp4 / temp96
                if PBR < 0:
                    PBR = np.nan


            if temp67 == 0:
                temp67 = np.nan
            maketAmt = temp20 * temp4
            if g_objCodeMgr.IsBigListingStock(temp0):
                maketAmt *= 1000
            #            print(code, maketAmt)

            # asset
            # debt =

            # key(종목코드) = tuple(상장주식수, 시가총액)
            dataInfo.append([temp0,temp17, maketAmt, PCR, PSR, PBR,temp67,
                               temp4, temp5,
                               temp10,
                               temp20,
                                temp68, temp69, temp70,
                               temp71, temp72, temp73, temp74,temp75, temp76, temp77, temp78, temp79,
                               temp80, temp81, temp82, temp83, temp84, temp85, temp86, temp87, temp88, temp89,
                               temp90,temp91, temp92, temp93, temp94, temp95, temp96, temp97, temp98, temp99,
                               temp100, temp101, temp102, temp103, temp104, temp105, temp106, temp107, temp108, temp109,
                               temp110, temp111, temp118,
                               temp120, temp123, temp124, temp125,
                               temp141])

        return True

# class CpStockMsg:
#     def __init__(self):
#         self.objRq = win32com.client.Dispatch("Dscbo1.StockMst")
#
#     def Request(self, code, dataInfo):
#         self.objRq.SetInputValue(0, code)
#         self.objRq.BlockRequest()
#
#         # 현재가 통신 및 통신 에러 처리
#         rqStatus = self.objRq.GetDibStatus()
#         print("통신상태", rqStatus, self.objRq.GetDibMsg1())
#         if rqStatus != 0:
#             return False
#
#         self.objRq.GetHeaderValue(3)
#         self.objRq.GetHeaderValue(4)
#         self.objRq.GetHeaderValue(44)










class CMarketTotal():

    FORMAT_DATETIME = "%Y-%m-%d"
    def __init__(self):
        self.dataInfo = []
        self.codelist = []
    def GetAllMarketTotal(self):
        codeList = g_objCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # 코스닥
        allcodelist = codeList + codeList2
        print('전 종목 코드 %d, 거래소 %d, 코스닥 %d' % (len(allcodelist), len(codeList), len(codeList2)))

        objMarket = CpMarketEye()
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            secondCode = g_objCodeMgr.GetStockSectionKind(code)
            status = g_objCodeMgr.GetStockStatusKind (code)
            control = g_objCodeMgr.GetStockControlKind  (code)
            if (secondCode == 1) and (status < 1) and (control < 3):
                rqCodeList.append(code)
            if len(rqCodeList) == 200:
                objMarket.Request(rqCodeList, self.dataInfo)
                rqCodeList = []
                continue
        # end of for

        if len(rqCodeList) > 0:
            objMarket.Request(rqCodeList, self.dataInfo)


    def PrintMarketTotal(self):

        # 시가총액 순으로 소팅
        # data2 = sorted(self.dataInfo.items(), key=lambda x: x[1][2], reverse=True)
        print(self.dataInfo)

        # print('전종목 시가총액 순 조회 (%d 종목)' % (len(data2)))
        # for item in data2:
        #     pass
            # print(item)
            # self.stock_list.append([i, code, secondCode, stdPrice, name])
        #
        # data3 = {'종목코드':self.dataInfo[0],
        #          '종목명':self.dataInfo[1][0]}
        # df_code = pd.DataFrame(data3)
        # data = {self.dataInfo}
        # data = {self.dataInfo}

        data_df = pd.DataFrame(self.dataInfo, columns=['종목코드', '종목명', '시가총액', 'PCR', 'PSR', 'PBR', 'PER', '현재가', '시가', '거래량',
                                                       '총상장주식수', '시간외매수잔량', '시간외매도잔량', 'EPS', '자본금', '액면가', '배당률', '배당수익률', '부채비율', '유보율',
                                                       '자기자본이익률', '매출액증가율', '경상이익증가율', '순이익증가율', '투자심리', 'VR', '5일 회전율', '4일 종가합', '9일 종가합', '매출액',
                                                       '경상이익', '당기순이익', 'BPS', '영업이익증가율', '영업이익', '매출액영업이익률', '매출액경상이익률', '이자보상비율', '결산년월', '분기BPS',
                                                       '분기매출액증가율', '분기영업이액증가율', '분기경상이익증가율', '분기순이익증가율', '분기매출액', '분기영업이익', '분기경상이익', '분기당기순이익', '분개매출액영업이익률', '분기매출액경상이익률',
                                                       '분기ROE', '분기이자보상비율', '분기유보율', '분기부채비율', '분기결산년월','당일외국인순매수', '당일기관순매수', 'SPS', 'CFPS', 'EBITDA', 'ELW 손익분기율'])
        data_df.index = data_df['종목코드']


        # sorted_df = self.make_low_cap(data_df)
        # sorted_df.to_excel(setting.DATAPATH.format(today, 'data_small', today))
        data_df.to_excel(setting.DATAPATH.format(today, 'data', today))


        print('finish')



    def make_low_cap(self, data_df):
        data_df = data_df.sort_values('시가총액', ascending=False)
        sorted_df = data_df[int(len(data_df) / 4) * 3: int(len(data_df) / 4) * 4]
        return sorted_df


class quant:
    FORMAT_DATETIME = "%Y-%m-%d"
    def __init__(self):
        pass

    # 저평가 지표 조합 함수
    def make_value_combo(self, value_list, invest_df, num):

        for i, value in enumerate(value_list):
            temp_df = self.get_value_rank_asc(invest_df, value, None)
            if i == 0:
                value_combo_df = temp_df
                rank_combo = temp_df[value + '순위']
            else:
                value_combo_df = pd.merge(value_combo_df, temp_df, how='outer', left_index=True, right_index=True)
                rank_combo = rank_combo + temp_df[value + '순위']

        value_combo_df['종합순위'] = rank_combo.rank()
        value_combo_df = value_combo_df.sort_values(by='종합순위')

        return value_combo_df[:num]

    #high GPA 추출
    def make_high_combo(self, value_list , invest_df, num):
        for i, value in enumerate(value_list):
            temp_df = self.get_value_rank_des(invest_df, value, None)
            if i == 0:
                value_combo_df = temp_df
                rank_combo = temp_df[value + '순위']
            else:
                value_combo_df = pd.merge(value_combo_df, temp_df, how='outer', left_index=True, right_index=True)
                rank_combo = rank_combo + temp_df[value + '순위']

        value_combo_df['종합순위'] = rank_combo.rank()
        value_combo_df = value_combo_df.sort_values(by='종합순위')

        return value_combo_df[:num]



    # F-score
    def get_fscore(self, fscore_df, num):
        fscore_df['당기순이익점수'] = fscore_df['분기당기순이익'] > 0
        fscore_df['분기영업이익점수'] = fscore_df['분기영업이익'] > 0
        fscore_df['분기매출액증가율점수'] = fscore_df['분기매출액증가율'] > 1.0
        fscore_df['분기영업이액증가율점수'] = fscore_df['분기영업이액증가율'] > 1.0
        fscore_df['분기경상이익증가율점수'] = fscore_df['분기경상이익증가율'] > 1.0
        fscore_df['분기ROE점수'] = fscore_df['분기ROE'] > 0.1
        fscore_df['분기유보율점수'] = (fscore_df['분기유보율'] > 150) & (fscore_df['분기유보율'] < 2000)
        fscore_df['분기부채비율점수'] = fscore_df['분기부채비율'] < 180

        if fscore_df['거래량'].any():
            fscore_df['당일외국인순매수점수'] = (fscore_df['당일외국인순매수'] / fscore_df['거래량']) > 0.05
        else:
            fscore_df['당일외국인순매수점수'] = False
        if fscore_df['거래량'].any():
            fscore_df['당일기관순매수점수'] = (fscore_df['당일기관순매수'] / fscore_df['거래량']) > 0.05
        else:
            fscore_df['당일기관순매수점수'] = False

        fscore_df['상승추세점수'] = fscore_df['9일 종가합'] > fscore_df['현재가']

        fscore_df['종합점수'] = fscore_df[['당기순이익점수', '분기영업이익점수', '분기매출액증가율점수', '분기영업이액증가율점수', '분기경상이익증가율점수', '분기ROE점수',  '분기유보율점수', '분기부채비율점수',
                                       '당일외국인순매수점수', '당일기관순매수점수', '상승추세점수']].sum(axis=1)

        fscore_df = fscore_df[fscore_df['종합점수'] > 7]

        return fscore_df[:num]

    def get_fscore_1(self, fscore_df, num):
        fscore_df['영업활동으로인한현금흐름점수'] = fscore_df['영업활동으로인한현금흐름'] > 0
        fscore_df['전년대비ROA증가율점수'] = fscore_df['전년대비ROA증가율'] > 0
        fscore_df['영업활동으로인한현금흐름점수_1'] = (fscore_df['영업활동으로인한현금흐름점수']/100000000) > fscore_df['분기당기순이익']
        fscore_df['전년대비총자산회전율점수'] = fscore_df['전년대비총자산회전율'] >= 0
        fscore_df['상장주식수변화량점수'] = fscore_df['상장주식수변화량'] <= 0
        fscore_df['종합점수'] = fscore_df[['영업활동으로인한현금흐름점수', '전년대비ROA증가율점수', '영업활동으로인한현금흐름점수_1', '전년대비총자산회전율점수', '상장주식수변화량점수']].sum(axis=1)
        print(len(fscore_df[fscore_df['영업활동으로인한현금흐름점수'] == 1]))
        print(len(fscore_df[fscore_df['전년대비ROA증가율점수'] == 1]))
        print(len(fscore_df[fscore_df['영업활동으로인한현금흐름점수_1'] == 1]))
        print(len(fscore_df[fscore_df['전년대비총자산회전율점수'] == 1]))
        print(len(fscore_df[fscore_df['상장주식수변화량점수'] == 1]))

        fscore_df = fscore_df[fscore_df['종합점수'] == 5]


        return fscore_df[:num]



    # 저평가 + F-score
    def get_value_quality(self, fs_df, num):
        # low_market_cap = self.make_market_cap(fs_df, 200)
        value = self.make_value_combo(['PER', 'PBR', 'PSR', 'PCR'], fs_df, None)
        quality = self.get_fscore_1(fs_df, None)
        value_quality = pd.merge(value, quality, how='outer', left_index=True, right_index=True)
        value_quality_filtered = value_quality[value_quality['종합점수'] == 5]
        vq_df = value_quality_filtered.sort_values(by='종합순위')
        return vq_df[:num]

    def get_value_quality_2(self, fs_df,num):
        value = self.make_high_combo(['GP/A'], fs_df, None)
        quality = self.get_fscore_1(fs_df, None)
        value_quality = pd.merge(value, quality, how='outer', left_index=True, right_index=True)
        value_quality_filtered = value_quality[value_quality['종합점수'] == 5]
        vq_df = value_quality_filtered.sort_values(by='종합순위')
        return vq_df[:num]


    # 저평가 지수를 기준으로 정렬하여 순위 만들어 주는 함수
    def get_value_rank_asc(self,invest_df, value_type, num):
        invest_df[value_type] = pd.to_numeric(invest_df[value_type])
        value_sorted = invest_df.sort_values(by=value_type)
        value_sorted[value_type + '순위'] = value_sorted[value_type].rank()
        return value_sorted[[value_type, value_type + '순위']][:num]

    # 저평가 지수를 기준으로 정렬하여 순위 만들어 주는 함수
    def get_value_rank_des(self,invest_df, value_type, num):
        invest_df[value_type] = pd.to_numeric(invest_df[value_type])
        value_sorted = invest_df.sort_values(by=value_type, ascending=False)
        value_sorted[value_type + '순위'] = value_sorted[value_type].rank(ascending=False)
        return value_sorted[[value_type, value_type + '순위']][:num]

    # 재무 관련 데이터 전처리하는 함수
    def get_load(self, path):
        data_path = path
        raw_data = pd.read_excel(data_path , index_col=0)
        return raw_data

    def save_finance_data(self, data_df, path):
        _path  = path.split('\\')
        dotpath =_path[3].split('.')
        path = str(_path[0]) + '\\' +str(_path[1]) +'\\'+ str(_path[2])  + '\\' +str(dotpath[0]) +'_sort.' + str(dotpath[1])
        data_df.to_excel(path)


    def get_finance_data(self, path):
        data_path = path
        raw_data = pd.read_excel(data_path, index_col=0)
        big_col = list(raw_data.columns)
        small_col = list(raw_data.iloc[0])

        new_big_col = []
        for num, col in enumerate(big_col):
            if 'Unnamed' in col:
                if num == 0:
                    new_big_col.append('종목코드')
                else:
                    new_big_col.append(new_big_col[num - 1])
            else:
                new_big_col.append(big_col[num])

        raw_data.columns = [new_big_col, small_col]
        clean_df = raw_data.loc[raw_data.index.dropna()]

        return clean_df

    # def get_data(self, main_df, sub_df):
    #     df = df_data_sub[(df_data_main['분기결산년월'],'총자산회전율')]
    #     print(df)





def float_to_str(df):
    df = str(df)
    df = df[:len(df)]


def data_merge_fr(value_list):
    quant_test = quant()
    df_data_main = quant_test.get_load(DATAPATH_1)

    path = setting.DATAPATH.format(today, '재무비율_년', today)
    df_data_sub = quant_test.get_finance_data(path)

    data = {}
    df_data = df_data_main['분기결산년월'].dropna()
    for i, contents in enumerate(value_list):
        for num, value in enumerate(df_data):
            if num >= len(df_data_sub):
                break
            value = str(value)
            value = value[:4] + "/" + value[4:len(value) - 2]
            data[df_data_sub.index[num]] = df_data_sub[(format(value), contents)].loc[df_data_sub.index[num]]
        add_srs = pd.Series(data)
        df_data_main[contents] = add_srs
    data_2 = {}
    for num, value in enumerate(df_data):
        if num >= len(df_data_sub):
            break
        value = str(value)
        year = int(value[:4]) - 1
        value_1 = str(year) + "/" + "12"
        value_2 = value[:4] + "/" + value[4:len(value) - 2]
        data[df_data_sub.index[num]] = float(df_data_sub[(format(value_2), '총자산회전율')].apply(check_IFRS).loc[df_data_sub.index[num]]) - float(df_data_sub[(format(value_1), '총자산회전율')].apply(check_IFRS).loc[df_data_sub.index[num]])
        data_2[df_data_sub.index[num]] = float(df_data_sub[(format(value_2), 'ROA')].apply(check_IFRS).loc[df_data_sub.index[num]]) - float(df_data_sub[(format(value_1), 'ROA')].apply(check_IFRS).loc[df_data_sub.index[num]])


    add_srs = pd.Series(data)
    add_srs_2 = pd.Series(data_2)
    df_data_main['전년대비총자산회전율'] = add_srs
    df_data_main['전년대비ROA증가율'] = add_srs_2

    df_data_main.to_excel(DATAPATH_1)

    return df_data_main

def data_merge_fs(value_list):
    quant_test = quant()
    df_data_main = quant_test.get_load(DATAPATH_1)

    path = setting.DATAPATH.format(today, '재무제표_년', today)
    df_data_sub = quant_test.get_finance_data(path)

    data = {}
    print(len(df_data_main))
    df_data = df_data_main['분기결산년월'].dropna()
    print(len(df_data))
    for i, contents in enumerate(value_list):

        for num, value in enumerate(df_data):
            print(num)
            if num >= len(df_data_sub):
                break
            value = str(value)
            value = value[:4] + "/" + value[4:len(value) - 2]
            data[df_data_sub.index[num]] = df_data_sub[(format(value), contents)].loc[df_data_sub.index[num]]
        add_srs = pd.Series(data)
        df_data_main[contents] = add_srs

    for num, value in enumerate(df_data):
        if num >= len(df_data_sub):
            break
        value = str(value)
        value = value[:4] + "/" + value[4:len(value) - 2]
        data[df_data_sub.index[num]] = df_data_sub[(format(value), '매출총이익')].loc[df_data_sub.index[num]] / \
                                       df_data_sub[(format(value), '자산')].loc[df_data_sub.index[num]]
    add_srs = pd.Series(data)
    df_data_main['GP/A'] = add_srs

    df_data_main.to_excel(DATAPATH_1)


    return df_data_main

def data_merge_iv(value_list):
    quant_test = quant()
    df_data_main = quant_test.get_load(DATAPATH_1)

    path = setting.DATAPATH.format(today, '투자지표_년', today)
    df_data_sub = quant_test.get_finance_data(path)

    data = {}
    df_data = df_data_main['분기결산년월'].dropna()
    for i, contents in enumerate(value_list):
        for num, value in enumerate(df_data):
            if num >= len(df_data_sub):
                break
            value = str(value)
            value = value[:4] + "/" + value[4:len(value) - 2]
            data[df_data_sub.index[num]] = df_data_sub[(format(value), contents)].loc[df_data_sub.index[num]]
        add_srs = pd.Series(data)
        df_data_main[contents] = add_srs

    df_data_main.to_excel(DATAPATH_1)

    return df_data_main

def data_merge_st(value_list):
    quant_test = quant()
    df_data_main = quant_test.get_load(DATAPATH_1)

    path = setting.DATAPATH.format(today, '상장주식수', today)
    df_data_sub = quant_test.get_finance_data(path)

    data = {}
    df_data = df_data_main['결산년월'].dropna()

    for i, contents in enumerate(value_list):
        for num, value in enumerate(df_data):
            if num >= len(df_data_sub):
                break
            value = str(value)
            year = int(value[:4]) - 1
            month = int(value[4:len(value) - 2])
            value = str(value)
            if month < 10:
                value_2 = value[:4] + "/" + "0" + str(month)
                value_1 = str(year) + "/" + "0" + str(month)
            else:
                value_2 = value[:4] + "/" + str(month)
                value_1 = str(year) + "/" + str(month)
            try:
                data[df_data_sub.index[num]] = float(df_data_sub[(format(value_2), '발행주식수')].loc[df_data_sub.index[num]]) - float(df_data_sub[(format(value_1), '발행주식수')].loc[df_data_sub.index[num]])
            except KeyError:
                data[df_data_sub.index[num]] = 0
        add_srs = pd.Series(data)
        df_data_main['상장주식수변화량'] = add_srs

    df_data_main.to_excel(DATAPATH_1)




#     순위 구하기
def data_rank():
    quant_test =  quant()
    df_data_1 = quant_test.get_load(DATAPATH_1)
    df_data_2 = quant_test.get_load(DATAPATH_2)
    df_sort_1 = quant_test.get_value_quality_2(df_data_1, None)
    df_sort_2 = quant_test.get_value_quality_2(df_data_2, None)
    quant_test.save_finance_data(df_sort_1, DATAPATH_1)
    quant_test.save_finance_data(df_sort_2, DATAPATH_2)

def data_load():
    # 매일 데이터 불러오기.
    objMarketTotal = CMarketTotal()
    objMarketTotal.GetAllMarketTotal()
    objMarketTotal.PrintMarketTotal()

def check_IFRS(x):
    if x == 'N/A(IFRS)':
        return np.NaN
    else:
        return x


def make_low_cap(data_df):
    data_df = data_df.sort_values('시가총액', ascending=False)
    sorted_df = data_df[int(len(data_df) / 4) * 3: int(len(data_df) / 4) * 4]
    return sorted_df

def make_small_data():
    quant_test = quant()
    df_data_main = quant_test.get_load(DATAPATH_1)
    sorted_df = make_low_cap(df_data_main)
    sorted_df.to_excel(DATAPATH_2)



if __name__ == "__main__":
    # 매일 데이터 불러오기.
    data_load()
    data_merge_fs(['자산', '부채', '자본', '영업활동으로인한현금흐름','매출총이익'])
    data_merge_fr(['총자산회전율', 'ROA'])
    data_merge_iv(['총현금흐름'])
    data_merge_st(['발행주식수'])
    make_small_data()
    data_rank()
    print('끝!!!!')

