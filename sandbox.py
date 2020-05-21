import pandas as pd
from pathlib import Path
from functools import reduce
from my_functions import max_pd_display

max_pd_display()

data_path = Path('C:/Users/ejmooney/Desktop/testData/tap_rule_data3.xlsx')

braden_df = pd.read_excel(data_path, 'Report 1')
assist_df = pd.read_excel(data_path, 'Report 2')
repos_df = pd.read_excel(data_path, 'Report 3')

comb_data = reduce(lambda x,y: pd.merge(x,y, on='fin', how='left'), [braden_df, assist_df, repos_df])

comb_data = comb_data.dropna()

summary = comb_data.groupby('location').fin.nunique()
summary