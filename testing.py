import pandas as pd
from pathlib import Path
from functools import reduce

pd.option_context('display.max_rows', None, 'display.max_columns', None)  # more options can be specified also
pd.options.display.max_colwidth = 199
pd.options.display.max_columns = 1000
pd.options.display.width = 1000
pd.options.display.precision = 2  # set as needed

tc_path = Path('C:/Users/ejmooney/Desktop/testData/tc_users.xlsx')
hr_path = Path('C:/Users/ejmooney/Desktop/testData/hr_report.xlsx')
alias_path = Path('C:/Users/ejmooney/Desktop/testData/alias_report.xlsx')

hr = pd.read_excel(hr_path)
tc = pd.read_excel(tc_path)
alias = pd.read_excel(alias_path)

alias['USERNAME'] = alias['USERNAME'].str.lower()

hr = hr.rename(columns={'EE ID': 'ALIAS'})

combinedData = reduce(lambda x,y: pd.merge(x,y, on='ALIAS', how='left'), [hr, alias])

tc = tc.rename(columns={'username':'USERNAME'})

combinedData = reduce(lambda x,y: pd.merge(x,y, on='USERNAME', how='left'), [combinedData, tc])

write_path = Path('C:/Users/ejmooney/Desktop/testData/combined_tc_hr_alias.xlsx')

combinedData.to_excel(write_path)