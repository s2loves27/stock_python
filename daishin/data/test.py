from daishin import setting

DATAPATH = r'.\data\{}\{}{}.xlsm'

print(DATAPATH.format(setting.get_today_str(), 'lowcap_sorted',setting.get_today_str() ))

