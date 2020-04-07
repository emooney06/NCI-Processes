import pandas as pd
from pathlib import Path
from my_functions import max_pd_display
from datetime import datetime

max_pd_display()

micu_drsng_path = Path('C:/Users/ejmooney/Desktop/testData/8528 MICU Dressing Changes.csv')
out_path = Path('C:/Users/ejmooney/Desktop/testData/8528 MICU Dressing Changes Output.xlsx')
df = pd.read_csv(micu_drsng_path)


df['result_date'] = df['RESULT_DT_TM'].str[:10]
df['result_date'] = pd.to_datetime(df['result_date'])
df = df.sort_values(['CE_DYNAMIC_LABEL_ID', 'RESULT_DT_TM'])
df['time_delta'] = df.groupby('CE_DYNAMIC_LABEL_ID')['result_date'].diff()

df['result_date'].dtype

df.to_excel(out_path)