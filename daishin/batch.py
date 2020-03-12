#  [코드 3.33] 필요한 모듈들 임포트 (CH3. 데이터 수집하기 2.ipynb)
import errno

import requests
import bs4
import pandas as pd
import time

# [코드 3.25] 다운받은 종목 엑셀파일 읽어오기 (CH3. 데이터 수집하기.ipynb)
import os

from daishin import setting

path = r'.\종목코드.xlsx'
code_data = pd.read_excel(path)
code_data = code_data[['종목코드']]

# [코드 3.26] 종목코드 컬럼 개선 (CH3. 데이터 수집하기.ipynb)
today = setting.get_today_str()

def make_code(x):
    x = str(x)
    return 'A' + '0' * (6-len(x)) + x

def delete_code(x):
    return x[1:]


# [코드 3.15] 재무제표 데이터를 가져와 데이터프레임으로 만드는 함수 (CH3. 데이터 수집하기.ipynb)

def make_fs_dataframe_year(firm_code):
    fs_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Finance.asp?pGB=1&gicode={}&cID=&MenuYn=Y&ReportGB=&NewMenuID=103&stkGb=701'.format(
        firm_code)
    fs_page = requests.get(fs_url)
    fs_tables = pd.read_html(fs_page.text)

    temp_df = fs_tables[0]
    temp_df = temp_df.set_index(temp_df.columns[0])
    temp_df = temp_df[temp_df.columns[:4]]
    temp_df = temp_df.loc[['매출액', '매출총이익', '영업이익', '당기순이익']]

    temp_df2 = fs_tables[2]
    temp_df2 = temp_df2.set_index(temp_df2.columns[0])
    temp_df2 = temp_df2.loc[['자산', '부채', '자본']]

    temp_df3 = fs_tables[4]
    temp_df3 = temp_df3.set_index(temp_df3.columns[0])
    temp_df3 = temp_df3.loc[['영업활동으로인한현금흐름']]

    fs_df = pd.concat([temp_df, temp_df2, temp_df3])

    return fs_df


def make_fs_dataframe_quarter(firm_code):
    fs_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Finance.asp?pGB=1&gicode={}&cID=&MenuYn=Y&ReportGB=&NewMenuID=103&stkGb=701'.format(
        firm_code)
    fs_page = requests.get(fs_url)
    fs_tables = pd.read_html(fs_page.text)

    temp_df = fs_tables[1]
    temp_df = temp_df.set_index(temp_df.columns[0])
    temp_df = temp_df[temp_df.columns[:4]]
    temp_df = temp_df.loc[['매출액', '매출총이익', '영업이익', '당기순이익']]

    temp_df2 = fs_tables[3]
    temp_df2 = temp_df2.set_index(temp_df2.columns[0])
    temp_df2 = temp_df2.loc[['자산', '부채', '자본']]

    temp_df3 = fs_tables[5]
    temp_df3 = temp_df3.set_index(temp_df3.columns[0])
    temp_df3 = temp_df3.loc[['영업활동으로인한현금흐름']]

    fs_df = pd.concat([temp_df, temp_df2, temp_df3])

    return fs_df

def make_is_dataframe(firm_code):

    fs_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Main.asp?pGB=1&gicode={}&cID=&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701'.format(firm_code)
    fs_page = requests.get(fs_url)
    fs_tables = pd.read_html(fs_page.text)

    temp_df = fs_tables[11]
    temp_df = temp_df.set_index(temp_df.columns[0])
    temp_df = temp_df.loc[['발행주식수']]
    if temp_df.empty == True:
        raise KeyError
    temp_df.index = ['발행주식수']

    new_big_col = temp_df.columns
    change_col = []
    for num, col in enumerate(new_big_col):
        change_col.append(col[1])
    temp_df.columns = change_col
    return temp_df


# [코드 3.21] 재무 비율 데이터프레임을 만드는 함수 (CH3. 데이터 수집하기.ipynb)

def make_fr_dataframe_year(firm_code):
    fr_url = 'https://comp.fnguide.com/SVO2/asp/SVD_FinanceRatio.asp?pGB=1&gicode={}&cID=&MenuYn=Y&ReportGB=&NewMenuID=104&stkGb=701'.format(
        firm_code)
    fr_page = requests.get(fr_url)
    fr_tables = pd.read_html(fr_page.text)

    temp_df = fr_tables[0]
    temp_df = temp_df.set_index(temp_df.columns[0])
    temp_df = temp_df.loc[['유동비율계산에 참여한 계정 펼치기',
                           '부채비율계산에 참여한 계정 펼치기',
                           '유보율계산에 참여한 계정 펼치기',
                           '매출액증가율계산에 참여한 계정 펼치기',
                           '매출총이익율계산에 참여한 계정 펼치기',
                           '영업이익증가율계산에 참여한 계정 펼치기',
                           '영업이익률계산에 참여한 계정 펼치기',
                           'ROA계산에 참여한 계정 펼치기',
                           'ROE계산에 참여한 계정 펼치기',
                           'ROIC계산에 참여한 계정 펼치기',
                           '총자산회전율계산에 참여한 계정 펼치기']]

    temp_df.index = ['유동비율', '부채비율', '유보율', '매출액증가율', '매출총이익율', '영업이익증가율', '영업이익률', 'ROA', 'ROE', 'ROIC', '총자산회전율']
    return temp_df


def make_fr_dataframe_quarter(firm_code):
    fr_url = 'https://comp.fnguide.com/SVO2/asp/SVD_FinanceRatio.asp?pGB=1&gicode={}&cID=&MenuYn=Y&ReportGB=&NewMenuID=104&stkGb=701'.format(
        firm_code)
    fr_page = requests.get(fr_url)
    fr_tables = pd.read_html(fr_page.text)

    temp_df2 = fr_tables[1]
    temp_df2 = temp_df2.set_index(temp_df2.columns[0])
    temp_df2 = temp_df2.loc[['매출액증가율계산에 참여한 계정 펼치기',
                             '영업이익증가율계산에 참여한 계정 펼치기',
                             '영업이익율계산에 참여한 계정 펼치기',
                             'EPS증가율계산에 참여한 계정 펼치기']]
    temp_df2.index = ['매출액증가율', '영업이익증가율', '영업이익율계산', 'EPS증가율계산']
    return temp_df2

# [코드 3.23] 투자지표 데이터프레임을 만드는 함수 (CH3. 데이터 수집하기.ipynb)

def make_invest_dataframe(firm_code):
    invest_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Invest.asp?pGB=1&gicode={}&cID=&MenuYn=Y&ReportGB=&NewMenuID=105&stkGb=701'.format(firm_code)
    invest_page = requests.get(invest_url)
    invest_tables = pd.read_html(invest_page.text)
    temp_df = invest_tables[1]
    temp_df = temp_df.set_index(temp_df.columns[0])
    temp_df = temp_df.loc[['PER계산에 참여한 계정 펼치기',
                           'PCR계산에 참여한 계정 펼치기',
                           'PSR계산에 참여한 계정 펼치기',
                           'PBR계산에 참여한 계정 펼치기',
                          '총현금흐름']]
    temp_df.index = ['PER', 'PCR', 'PSR', 'PBR', '총현금흐름']

    return temp_df


#  [코드 3.40] 가격을 가져와 데이터프레임 만드는 함수 (CH3. 데이터 수집하기 2.ipynb)

def make_fr_dataframe_sales(firm_code):
    fr_url = 'https://comp.fnguide.com/SVO2/ASP/SVD_Corp.asp?pGB=1&gicode={}&cID=&MenuYn=Y&ReportGB=&NewMenuID=102&stkGb=701'.format(
        firm_code)
    fr_page = requests.get(fr_url)
    fr_tables = pd.read_html(fr_page.text)

    #     print(fr_tables[2])
    temp_df = fr_tables[2][:10]
    product_name = temp_df['제품명']

    temp_df = fr_tables[3][:10]
    main_product = temp_df['주요제품']

    temp_df = fr_tables[10][:10]
    product = temp_df['제품명']

    sales = pd.concat([product_name, main_product, product], axis=1)

    sales.columns = ['판매제품군', '주요제품', '제품명']

    return sales


# [코드 3.19] 데이터프레임 형태 바꾸기 코드 함수화 (CH3. 데이터 수집하기.ipynb)
def change_df(firm_code, dataframe):

    for num, col in enumerate(dataframe.columns):
        temp_df = pd.DataFrame({firm_code: dataframe[col]})
        temp_df = temp_df.T
        temp_df.columns = [[col] * len(dataframe), temp_df.columns]
        if num == 0:
            total_df = temp_df
        else:
            total_df = pd.merge(total_df, temp_df, how='outer', left_index=True, right_index=True)

    return total_df

def save_fs_year():

    for num, code in enumerate(code_data['종목코드']):
        try:
            print(num, code)
            time.sleep(0.1)
            try:
                fs_ds_y = make_fs_dataframe_year(code)
            except requests.exceptions.Timeout:
                time.sleep(60)
                fs_ds_y = make_fs_dataframe_year(code)
            except requests.exceptions.ConnectionError:
                time.sleep(30)
                fs_ds_y = make_fs_dataframe_year(code)
            fs_ds_changed = change_df(code, fs_ds_y)
            if num == 0:
                total_fs_y = fs_ds_changed
            else:
                total_fs_y = pd.concat([total_fs_y, fs_ds_changed])
        except ValueError:
            continue
        except KeyError:
            continue

    total_fs_y.to_excel(setting.DATAPATH.format(today, '제무제표_년', today))


def save_fr_year():
    today = setting.get_today_str()
    for num, code in enumerate(code_data['종목코드']):
        try:
            print(num, code)
            time.sleep(0.1)
            try:
                fs_ds_y = make_fr_dataframe_year(code)
            except requests.exceptions.Timeout:
                time.sleep(60)
                fs_ds_y = make_fs_dataframe_year(code)
            except requests.exceptions.ConnectionError:
                time.sleep(30)
                fs_ds_y = make_fs_dataframe_year(code)
            fs_df_changed = change_df(code, fs_ds_y)
            if num == 0:
                total_fs = fs_df_changed
            else:
                total_fs = pd.concat([total_fs, fs_df_changed])
        except ValueError:
            continue
        except KeyError:
            continue
    total_fs.to_excel(setting.DATAPATH.format(today, '재무비율_년', today))

def save_iv_year():
    today = setting.get_today_str()
    for num, code in enumerate(code_data['종목코드']):
        try:
            print(num, code)
            time.sleep(0.1)
            try:
                fs_df = make_invest_dataframe(code)
            except requests.exceptions.Timeout:
                time.sleep(60)
                fs_df = make_invest_dataframe(code)
            except requests.exceptions.ConnectionError:
                time.sleep(30)
                fs_df = make_invest_dataframe(code)
                print('ConnetError')
            fs_df_changed = change_df(code, fs_df)
            if num == 0:
                total_iv = fs_df_changed
            else:
                total_iv = pd.concat([total_iv, fs_df_changed])
        except ValueError:
            continue
        except KeyError:
            continue
    total_iv.to_excel(setting.DATAPATH.format(today, '투자지표_년', today))

def save_is():
    today = setting.get_today_str()
    for num, code in enumerate(code_data['종목코드']):
        try:
            print(num, code)
            time.sleep(0.1)
            try:
                fs_df = make_is_dataframe(code)
            except requests.exceptions.Timeout:
                time.sleep(60)
                fs_df = make_is_dataframe(code)
                print('TimeoutError')
            except requests.exceptions.ConnectionError:
                time.sleep(30)
                fs_df = make_is_dataframe(code)
                print('ConnetError')
            fs_df_changed = change_df(code, fs_df)
            if num == 0:
                total_iv = fs_df_changed
            else:
                total_iv = pd.concat([total_iv, fs_df_changed])
        except ValueError:
            continue
        except KeyError:
            continue

    total_iv.to_excel(setting.DATAPATH.format(today, '상장주식수', today))

def make_dir():
    today = setting.get_today_str()
    try:
        if not (os.path.isdir(r'.\data\{}'.format(today))):
            os.makedirs(os.path.join(r'.\data\{}'.format(today)))
    except OSError as e:
        if e.errno != errno.EEXIST:
            print("Failed to create directory!!!!!")
            raise


if __name__ == "__main__":
    make_dir()
    save_fs_year()
    # save_fr_year()
    # save_iv_year()
    # save_is()


