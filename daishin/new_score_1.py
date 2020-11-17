import ctypes
import datetime

import pandas as pd
import numpy as np

# Settings on Logging
import win32com.client

from daishin import setting

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')

FORMAT_DATETIME = "%Y-%m-%d"
today = setting.get_today_str()
FORMAT_DATETIME = "%Y%m%d"
today_1 = setting.get_today_str_1()

DATAPATH_1 = setting.DATAPATH.format(today, 'new_low', today)

def get_today_str():
    global today
    today = datetime.datetime.combine(datetime.date.today(), datetime.datetime.min.time())
    today_str = today.strftime('%Y%m%d')
    return today_str


def load_data(df):
    today = get_today_str()
    new_df = setting.read_mongo('new', {'$and': [{'날짜': {'$eq': today_1}, 'Update': {'$eq': False},'상태':{'$eq':1}}]})
    # setting.update_mongo('new', {'Update': False}, {"$set": {'Update': True}})
    f1 = lambda x : datetime.datetime.strptime(x,"%H%M")

    new_df.drop(['_id'], axis=1, inplace=True)
    convert_time = new_df['시간'].apply(f1)


    start = datetime.datetime.strptime('0900',"%H%M")
    finish = datetime.datetime.strptime('1530',"%H%M")

    new_df['점수'] = .0
    # print(new_df)
    for i in range(len(convert_time)):
        # if convert_time[i] >= start and convert_time[i] =< finish:
        if True:
            try:
                new_df.iloc[i]['특이사항'] = new_df.iloc[i]['특이사항'].replace(new_df.iloc[i]['종목명'],"")
                # print((((int(df['현재가'][new_df.iloc[i]['코드']]) - int(df['시가'][new_df.iloc[i]['코드']]) )/ int(df['시가'][new_df.iloc[i]['코드']]) * 100) / 3))
                # value = ((float(df['전일종가'][new_df.iloc[i]['코드']]) - float(df['시가'][new_df.iloc[i]['코드']]))/ float(df['전일종가'][new_df.iloc[i]['코드']]) * 100)
                value = ((df['전일대비'][new_df.iloc[i]['코드']] /df['전일종가'][new_df.iloc[i]['코드']] ) * 100)
                # new_df['Score_Update'][i] = True
                if df['목록'][new_df.iloc[i]['코드']] == '코스피':
                    market = ((df['전일대비']['U001'] / df['전일종가']['U001']) * 100)
                else :
                    market = ((df['전일대비']['U201'] / df['전일종가']['U201']) * 100)
                # print(market , value)
                if market < value:
                    score = (10 / (30 -  market)) * (value - market)
                else:
                    score = (10 / (-30 - market)) * (market - value)
                # print(score)
                new_df['점수'][i] = round(score)

                # new_df['점수'][i] = (int(df['현재가'][new_df['코드']])- int(['시가'][new_df['코드']]) / int(df['시가'][new_df['코드']]) * 100) / 3
            except KeyError:
                pass
                # new_df['점수'][i] = (int(df['현재가'][new_df['코드']])- int(['시가'][new_df['코드']]) / int(df['시가'][new_df['코드']]) * 100) / 3
            except ZeroDivisionError:
                pass


    print('DB SAVE')
    # setting.write_mongo('new_replace', new_df)
    print('EXCEL SAVE')
    new_df.to_excel(setting.NEWPATH.format(today), index = False)


    # df.to_excel()


    print('finish')

# 재무 관련 데이터 전처리하는 함수
def strTodate(str_value):
    return datetime.datetime.strptime(str_value,"%H%M").date()

def get_load(path):
    data_path = path
    raw_data = pd.read_excel(data_path, index_col=0)
    return raw_data


if __name__ == "__main__":
    df = get_load(DATAPATH_1)
    load_data(df);