from pathlib import Path
import pandas as pd
from my_functions import max_pd_display
import numpy as np
import datetime

max_pd_display()

desktop_path = Path('C:/Users/ejmooney/Desktop')
hr_path = Path('//uh-nas/Groupshare3/NDNQI/SourceData/Nurse Turnover/HR Turnover Reports/all')
#create an array to iterate through file names in folder
file_dates = ['2017-01', '2017-02', '2017-03', '2017-04', '2017-05', '2017-06', '2017-07', '2017-08', '2017-09', '2017-10', '2017-11', '2017-12', '2018-01', '2018-02', 
              '2018-03', '2018-04', '2018-05', '2018-06', '2018-07', '2018-08', '2018-09', '2018-10', '2018-11', '2018-12', '2019-01', '2019-02', '2019-03', '2019-04',
             '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11', '2019-12', '2020-01', '2020-02', '2020-03', '2020-04', '2020-05']

# read the export from the resident data smartsheet
res_df = pd.read_excel(desktop_path / 'res_data.xlsx', usecols='A:C, D')
#replace the blank employee IDs with NAN so they can easily be dropped
res_df['EE ID'].replace('', np.nan, inplace=True)
#drop the EE ID NANs
res_df.dropna(subset=['EE ID'], inplace=True)
#preserve a clean copy of the residency data for use in later merge operations
combined = res_df

for date in file_dates:
    hr_file = 'HR Report ' + date + '.xls'
    hr_df = pd.read_excel(hr_path / hr_file, usecols='b:c, i:j')
    xy_df = res_df
    xy_df = xy_df.join(hr_df, lsuffix='EE ID')
    merge_df = res_df.merge(hr_df, left_on='EE ID', right_on='EE ID')
    merge_df['unit date'] = date

    combined = combined.append(merge_df, sort=False)


termed_df = combined
termed_df = termed_df.replace(r'^\s*$', np.nan, regex=True)
termed_df = termed_df.dropna(subset=['Term date'])

combined2 = combined.sort_values('name', ascending=False)
combined2 = combined2.drop_duplicates(subset=['EE ID', 'Department'])
combined2 = combined2.dropna()

combined2 = combined2.append(termed_df, sort=False)
combined2 = combined2.sort_values(['name','unit date'], ascending=True)

total.to_csv(desktop_path / 'total_set.csv')
total = combined2




total['Hire date'] = total['Hire date'].values.astype('datetime64[M]')
total['unit date'] = total['unit date'].values.astype('datetime64[M]')
total['diff_in_months'] = ( total['unit date'] - total['Hire date']).astype('timedelta64[M]')

conditions = [
    (total['diff_in_months'] <=6),
    (total['diff_in_months'] > 6) & (total['diff_in_months'] <= 18),
    (total['diff_in_months'] > 18) & (total['diff_in_months'] <= 24)]
choices = ['6_mo', '12_mo', '24_mo']
total['tenure'] = np.select(conditions, choices, default='>24_mo')

grouped_df = total.reset_index().groupby(['EE ID', 'tenure'])['Department'].aggregate('first').unstack()

grouped_df.to_csv(desktop_path / 'test_res.csv')

new_df = pd.read_csv(desktop_path / 'test_res.csv')

new_df = new_df.merge(res_df, left_on='EE ID', right_on='EE ID')

termed_df = termed_df[['EE ID', 'Term date']]
new_df = new_df.join(termed_df.set_index('EE ID'), on='EE ID')

res_df = res_df[['EE ID', 'hire_date']]

new_df = new_df.join(res_df.set_index('EE ID'), on='EE ID', lsuffix='_left')

new_df = new_df[['EE ID', 'name', 'hire_date', 'hire_location', '6_mo','12_mo', '24_mo', '>24_mo', 'Term date']]

new_df.to_csv(desktop_path / 'final_set.csv', index=False)