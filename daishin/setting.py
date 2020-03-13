'''
MongoDB 연동
'''
import pymongo
from pymongo import MongoClient  # from이 있을경우 import뒤에있는건 함수 인데 db연동 모듈 추가역할을 합니다
import datetime
import pandas as pd
import time


DATAPATH = r'.\data\{}\{}{}.xlsx'
NEWPATH = r'.\data\new{}.xlsx'


collect = None
conn = None


def get_connetion(host = '127.0.0.1', port = 27017, db = 'DeepLearning'):
    """ A util for making a connection to mongo """

    # if username and password:
    #     mongo_uri = 'mongodb://%s:%s@%s:%s/%s' % (username, password, host, port, db)
    #     conn = MongoClient(mongo_uri)
    # else:
    global conn
    conn = MongoClient(host, port)

    return conn[db]

def get_collect( col, host = '127.0.0.1', port = 27017, db = 'DeepLearning'):
    global conn
    global collect

    if conn == None:
        get_connetion()
    collect = conn[db][col]
    return collect

def read_mongo(col, query, host = '127.0.0.1', port = 27017 , db = 'DeepLearning'):
    global collect

    if collect == None:
        get_collect(col)
    if query == None:
        cursor = collect.find()
    else:
        cursor = collect.find(query)
    return pd.DataFrame(list(cursor))

def write_mongo(col, df):
    global collect
    if collect == None:
        get_collect(col)
    dict = df.to_dict('records')
    if(len(dict) == 0):
        pass
    elif(len(dict) == 1):
        collect.insert_one(df.to_dict('records'))
    else:
        collect.insert_many(df.to_dict('records'))


def update_mongo(col, query , document):
    global collect
    if collect == None:
        get_collect(col)
    collect.update_many(query, document)

# df = pd.DataFrame({'high':[1,2,6],'low':[4,5,5],'Update':[False,False,True]})
# print(df)
# # write_mongo('new',df)
# update_mongo('new',{'Update':False}, {"$set":{'Update':True}})




# Date Time Format
timestr = None
today = None
FORMAT_DATE = "%Y%m%d"
FORMAT_DATETIME = "%Y%m%d%H%M%S"

# Settings on Logging
def get_today_str():
    global today
    today = datetime.datetime.combine(datetime.date.today(), datetime.datetime.min.time())
    today_str = today.strftime('%Y-%m-%d')
    return today_str

def get_time_str():
    global timestr
    timestr = datetime.datetime.fromtimestamp(
        int(time.time())).strftime(FORMAT_DATETIME)
    return timestr