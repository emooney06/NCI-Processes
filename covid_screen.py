import pandas as pd
import os
from functools import reduce
import string
import numpy as np
import collections
import re
from datetime import datetime
from pathlib import Path
from my_functions import max_pd_display, check_answer, make_string_cost_center, add_columns_for_reporting, double_check
from my_classes import FileDateVars
from my_variables import master_alias, mmm_dict
import win32com.client
import time 


max_pd_display()
data_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen')
archive_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen/8940_archive')
file_name = 'covid_screen.xlsx'
timestr = time.strftime("_%Y%m%d-%H%M")
archive_file =  timestr + file_name 

display_email = True

df = pd.read_excel(data_path / file_name)
df.to_excel(archive_path / archive_file)

filter_by = ['P ICN-3 (IN3P)', 'P ICN-4 (IN4P)', 'P NBICU (NBIP)', 'P NB Nursery (NBNP)', 
                'P Admit Prep (APIP)', 'P MCICU (MCIP)', 'P OB Spec Care (HRMP)', 'P TSICU (TSIP)']


df = df[~df.location.isin(filter_by)]

pos_scrn_df = df[(df['exposure_result'] != 'No high exposure risk') &
                    (df['symptoms_result'] != 'No high risk symptoms')]

pos_scrn_not_neg_test_df = pos_scrn_df[(pos_scrn_df['testing_result'] != 'Not detected') &
                                        (pos_scrn_df['testing_result'] != 'Detected')]

pos_scrn_not_neg_test_df['careset_order'] = pos_scrn_not_neg_test_df.careset_order.replace(np.nan, 'no results found', regex=True)
pos_scrn_not_neg_test_df['testing_result'] = pos_scrn_not_neg_test_df.testing_result.replace(np.nan, 'no results found', regex=True)
pos_scrn_not_neg_test_df['exposure_result'] = pos_scrn_not_neg_test_df.exposure_result.replace(np.nan, 'no results found', regex=True)
pos_scrn_not_neg_test_df['symptoms_result'] = pos_scrn_not_neg_test_df.symptoms_result.replace(np.nan, 'no results found', regex=True)



pos_scrn_not_neg_test_df['careset_order_dt_tm'] = pos_scrn_not_neg_test_df.careset_order_dt_tm.astype(object).where(pos_scrn_not_neg_test_df.careset_order_dt_tm.notnull(), 'no results found')
pos_scrn_not_neg_test_df['testing_dt_tm'] = pos_scrn_not_neg_test_df.testing_dt_tm.astype(object).where(pos_scrn_not_neg_test_df.testing_dt_tm.notnull(), 'no results found')
pos_scrn_not_neg_test_df['exposure_dt_tm'] = pos_scrn_not_neg_test_df.exposure_dt_tm.astype(object).where(pos_scrn_not_neg_test_df.exposure_dt_tm.notnull(), 'no results found')
pos_scrn_not_neg_test_df['symptoms_dt_tm'] = pos_scrn_not_neg_test_df.symptoms_dt_tm.astype(object).where(pos_scrn_not_neg_test_df.symptoms_dt_tm.notnull(), 'no results found')


pos_scrn_not_neg_test_df.to_excel(data_path / 'test_out.xlsx')

ma_df = pd.read_excel(data_path / 'ma_copy.xlsx')
ma_df = ma_df[['cerner_unit_name', 'UD_Email']]
ma_df = ma_df.dropna()
ma_df = ma_df.rename(columns={'cerner_unit_name': 'location'})

email_dict = ma_df.set_index('location')['UD_Email'].to_dict()

location_list = pos_scrn_not_neg_test_df.location.unique()


#concern_df = reduce(lambda x,y: pd.merge(x,y, on='location', how='left'), [pos_scrn_not_neg_test_df, ma_df])


olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
global unit_table

for location in location_list:
    try:
        #cost_center = str(cost_center)
        # initialize variable that is email address which is a found from the email dictionary object
        email = email_dict.get(location)
        if email == None:
            email = 'no email found for this unit: ' + location
    except KeyError:
        # If there is no UD Email associated with that cost center, just print it to console
        print('There is no email address for: ') 
        print(email)
    unit_df = pos_scrn_not_neg_test_df[(pos_scrn_not_neg_test_df['location']) == location]
    unit_table = unit_df.to_html()
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = 'FYI - Possible COVID-19 Risk *Secure*'
    newMail.To = email
    newMail.Cc = 'ejmooney@salud.unm.edu'
    greeting = '''\
    <html> 
        <head> 
            <font size='4'> 
            Hello Unit Director,<br><br> This is an automated message from your Nursing 
            Clinical Informatics team.  This message is for your information only - no response is needed.  
            <br><br>
            Below you will find a patient identified by our algorithm as a potential COVID-19 exposure risk.  This process is 
            intended to identify patients who have not been screened and/or have not yet been tested for COVID-19. 
            <br><br>
            Please be advised: <br><br>
            While we have employed a new analytic process to minimize the volume of non-actionable notifications, we fully 
            anticipate there will be some rate of error in our process and we do not intend for this to replace a clinician review 
            of the medical record.  Please consider this notice a "heads-up" that you may want to look into the 
            records listed below for appropriate testing and screening.
            <br><br>
            If you find a patient needs COVID-19 Testing, you may inform the provider that testing can be found in the "COVID-19 Test 
            careset.  If you find a patient should be screened, the screening can be found in the ad-hoc form titled "Infectious Disease 
            Travel Screening".  
            <br><br>
            As always, we welcome any questions or feedback. <br><br>
            Ethan Mooney, RN, MSN<br>
            Nursing Clinical Informatics<br><br><br>
            </font>
        </head>
        <body><font size='4'>COVID-19 Risk Summary</font></body>            
    <html>
    '''
    html = greeting + unit_table
    newMail.HTMLBody = html
    newMail.Display()

#time.sleep(20)

