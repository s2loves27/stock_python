from daishin import setting
import pandas as pd
df = pd.DataFrame({'a':[1,2], 'b':[2,3]})
print(df)
setting.write_mongo('new_replace', df)
