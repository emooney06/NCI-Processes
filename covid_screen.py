import pandas as pd
from my_functions import max_pd_display
from pathlib import Path
from functools import reduce

max_pd_display()

data_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940 - covid-screen')
file_name = 'covid_screen.xlsx'

df = pd.read_excel(data_path / file_name)

filter_by = ['P ICN-3 (IN3P)', 'P ICN-4 (IN4P)', 'P NBICU (NBIP)', 'P NB Nursery (NBNP)', 
             'P Admit Prep (APIP)']


df = df[~df.location.isin(filter_by)]

pos_scrn_df = df[(df['exposure_result'] != 'No high exposure risk') &
                  (df['symptoms_result'] != 'No high risk symptoms')]

pos_scrn_not_neg_test_df = pos_scrn_df[(pos_scrn_df['testing_result'] != 'Not detected') &
                                       (pos_scrn_df['testing_result'] != 'Detected')]

ma_df = pd.read_excel(data_path / 'ma_copy.xlsx')
ma_df = ma_df[['cerner_unit_name', 'UD_Email']]
ma_df = ma_df.dropna()
ma_df = ma_df.rename(columns={'cerner_unit_name': 'location'})

concern_df = reduce(lambda x,y: pd.merge(x,y, on='location', how='left'), [pos_scrn_not_neg_test_df, ma_df])

pos_scrn_not_neg_test_df.to_csv(data_path / 'test_out.csv')