import pandas as pd
import requests
import os

gExcelFile = '1.xlsx'
fs_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Finance.asp?pGB=1&gicode=A005930&cID=&MenuYn=Y&ReportGB=&NewMenuID=103&stkGb=701'
fs_page = requests.get(fs_url)
fs_tables = pd.read_html(fs_page.text)
writer = pd.ExcelWriter(gExcelFile, engine='xlsxwriter')

temp_df = fs_tables[0]
temp_df = temp_df.set_index(temp_df.columns[0])
temp_df = temp_df[temp_df.columns[:4]]
temp_df = temp_df.loc[['매출액','영업이익','당기순이익']]

temp_df2 = fs_tables[2]
temp_df2 = temp_df2.set_index(temp_df2.columns[0])
temp_df2 = temp_df2.loc[['자산','부채','자본']]

temp_df3 = fs_tables[4]
temp_df3 = temp_df3.set_index(temp_df3.columns[0])
temp_df3 = temp_df3.loc[['영업활동으로인한현금흐름']]


fs_df = pd.concat([temp_df,temp_df2,temp_df3])

fs_df.to_excel(writer, sheet_name='Sheet1')
# Close the Pandas Excel writer and output the Excel file.
writer.save()
os.startfile(gExcelFile)

print(fs_df)

def change_df(firm_code, dataframe):
    for num, col in enumerate(dataframe.columns):
        temp_df = pd.DataFrame({firm_code: dataframe[col]})
        temp_df = temp_df.T
        temp_df.columns = [[col]*len(dataframe),temp_df.columns]
        if num == 0:
            total_df = temp_df
        else:
            total_df = pd.merge(total_df, temp_df, how = 'outer',left_index = True,right_index =True)
    return total_df

def make_fr_dataframe(firm_code):
    fr_url = 'https://comp.fnguide.com/SVO2/asp/SVD_FinanceRatio.asp?pGB=1&cID=&MenuYn=Y&ReportGB=D&NewMenuID=104&stkGb=701&gicode=' + firm_code
    fr_page = requests.get(fr_url)
    fr_tables = pd.read_html(fr_page.text)

    temp_df = fr_tables[0]
    temp_df = temp_df.set_index(temp_df.columns[0])
    temp_df = temp_df.loc[['유동비율(유동자산 / 유동부채) * 100 유동비율계산에 참여한 계정 펼치기',
                           '부채비율(총부채 / 총자본) * 100 부채비율계산에 참여한 계정 펼치기',
                           '영업이익률(영업이익 / 영업수익) * 100 영업이익률계산에 참여한 계정 펼치기',
                           'ROA(당기순이익(연율화) / 총자산(평균)) * 100 ROA계산에 참여한 계정 펼치기',
                           'ROIC(세후영업이익(연율화)/영업투하자본(평균))*100 ROIC계산에 참여한 계정 펼치기']]
    temp_df.index = ['유동비율', '부채비율', '영업이익률', 'ROA', 'ROIC']
    return temp_df


def make_invest_dataframe(firm_code):
    invest_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Invest.asp?pGB=1&cID=&MenuYn=Y&ReportGB=D&NewMenuID=105&stkGb=701&gicode=' + firm_code
    invest_page = requests.get(invest_url)
    invest_tables = pd.read_html(invest_page.text)
    temp_df = invest_tables[1]
    
    temp_df = temp_df.set_index(temp_df.columns[0])
    temp_df = temp_df.loc[['PER수정주가(보통주) / 수정EPS PER계산에 참여한 계정 펼치기',
                           'PCR수정주가(보통주) / 수정CFPS PCR계산에 참여한 계정 펼치기',
                           'PSR수정주가(보통주) / 수정SPS PSR계산에 참여한 계정 펼치기',
                           'PBR수정주가(보통주) / 수정BPS PBR계산에 참여한 계정 펼치기',
                          '총현금흐름세후영업이익 + 유무형자산상각비 총현금흐름']]
    temp_df.index = ['PER', 'PCR', 'PSR', 'PBR', '총현금흐름']
    return temp_df


    

