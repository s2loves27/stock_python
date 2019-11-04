## utf-8 codec can't decode byte 0xc5 in position 0: invalid continuation byte
# 위 에러는 윈도우 기본인 CP949 또는 EUC-KR로 CSV를 저장했는데 UTF-8로 CSV를 읽을 때 발생합니다.
# 이 경우 Pandas read_csv() 함수의 파아미터에 encoding='CP949'를 추가해서 read_csv를 호출 하면 해결 됩니다.


## "pandas/_libs/parsers.pyx" ....
## "pandas/_libs/parsers.pyx" ....
# 이 경우 파일명을 영문과 숫자로만 구성하거나 read_csv() 함수에 engine='python'를 넣어서 read_csv를 호출 하면됩니다.



import pandas as pd
import numpy as np
# CSV 파일 경로를 입력으로 받습니다. Pandas의 read_csv()함수의 첫 번째 인자에 이 파일 경로를 입력합니다.

def load_chart_data(fpath):
    # thousands 파라미터로 ','를 넣어주면 1,234,567과 같이 천단위로 콤마가 붙은 값을 숫자로 인식합니다.
    # header=None는 헤더가 없다는 것을 알려주기 위해서 사용 하였습니다.
    chart_data = pd.read_csv(fpath, thousands=',', header=None)
    chart_data.columns = ['date', 'open', 'high', 'low', 'close', 'volume']
    return chart_data

#Pandas의 rolling(window) 함수는 window 크기만큼 데이터를 묶어서 합, 평균, 표준편차등을 계산할 수 있께 준비합니다.
# 이를 이동합, 이동평균, 이동표준편차라고 합니다.
# 이동합 <Pandas 객체>.rolling().sum()
# 이동평균 <Pandas 객체>.rolling().mean()
# 이동표준편차 <Pandas 객체>.rolling().std()
def preprocess(chart_data):
    prep_data = chart_data
    windows =[5,10,20,60,120]
    for window in windows:
        prep_data['close_ma{}'.format(window)] = prep_data['close'].rolling(window).mean()
        prep_data['volume_ma{}'.format(window)] = (
            prep_data['volume'].rolling(window).mean()
        )
    return prep_data


def build_training_data(prep_data):
    training_data = prep_data

    training_data['open_lastclose_ratio'] = np.zeros(len(training_data))
    training_data['open_lastclose_ratio'].iloc[1:] = \
        (training_data['open'][1:].values - training_data['close'][:-1].values) / \
        training_data['close'][:-1].values
