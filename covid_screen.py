import pandas as pd
import os
from functools import reduce
import string
import numpy as np
import collections
import re
from datetime import datetime
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from pathlib import Path
from my_functions import max_pd_display, check_answer, make_string_cost_center, add_columns_for_reporting, double_check
from my_classes import FileDateVars
from my_variables import master_alias, mmm_dict
import sys
import win32com.client



max_pd_display()

display_email = True

data_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen')
file_name = '2020-05_covid_screen.xlsx'

df = pd.read_excel(data_path / file_name)

filter_by = ['P ICN-3 (IN3P)', 'P ICN-4 (IN4P)', 'P NBICU (NBIP)', 'P NB Nursery (NBNP)', 
             'P Admit Prep (APIP)', 'P MCICU (MCIP)', 'P OB Spec Care (HRMP)', 'P TSICU (TSIP)']


df = df[~df.location.isin(filter_by)]

pos_scrn_df = df[(df['exposure_result'] != 'No high exposure risk') &
                  (df['symptoms_result'] != 'No high risk symptoms')]

pos_scrn_not_neg_test_df = pos_scrn_df[(pos_scrn_df['testing_result'] != 'Not detected') &
                                       (pos_scrn_df['testing_result'] != 'Detected')]

ma_df = pd.read_excel(data_path / 'ma_copy.xlsx')
ma_df = ma_df[['cerner_unit_name', 'UD_Email']]
ma_df = ma_df.dropna()
ma_df = ma_df.rename(columns={'cerner_unit_name': 'location'})

email_dict = ma_df.set_index('location')['UD_Email'].to_dict()

location_list = pos_scrn_not_neg_test_df.location.unique()

#concern_df.to_excel(data_path / 'test_out.xlsx')
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
    newMail.Subject = 'FYI - Possible COVID-19 Risk'
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
            of the medical record.  Please consider this notice simply a "heads-up" that you may want to look into the 
            records listed below a little closer.
            <br><br>
            As always, we welcome any questions or feedback. <br><br>
            Thank You,<br>
            Ethan Mooney, RN, MSN, PCCN<br>
            Nursing Clinical Informatics<br><br><br>
            </font>
        </head>
        <body><font size='4'>COVID-19 Risk Summary</font></body>            
    <html>
    '''
    html = greeting + unit_table
    newMail.HTMLBody = html
    newMail.Display()




newMail.HTMLBody = html
newMail.Display()



if send_email == True:
    x = 0
    #alias list is a list of all cost centers from MasterAliasRecord
    for line in cost_center_list:
        # variable cost_center be what the cost center is for this iteration
        cost_center = line
        str_cost_center = str(cost_center)
        # try and except block is used because not every cost center has a Unit Director email, without
        #   the try and except block it will throw exception error when one of these are encountered
        try:
            #cost_center = str(cost_center)
            # initialize variable that is email address which is a found from the email dictionary object
            email = email_dict.get(str_cost_center)
            unit_name = unit_name_dict.get(str_cost_center)
            if unit_name == None:
                unit_name = 'no unit found for cost center: ' + line
        except KeyError:
            # If there is no UD Email associated with that cost center, just print it to console
            print('There is no email address for: ') 
            print(str_cost_center)
        # edu_data is the dataframe of all education data but filtered by the cost center for this iteration
        temp_data = complete_df.loc[complete_df['UNMH_Cost_Center'] == str_cost_center]
        edu_totals = temp_data['Degree'].value_counts()   
        edu_totals = pd.DataFrame(edu_totals)
        edu_totals['degree_pct'] = (edu_totals['Degree']/edu_totals['Degree'].sum()) * 100 
        edu_totals['degree_pct'] = edu_totals['degree_pct'].round(1)
        edu_totals['degree_pct'] = (edu_totals['degree_pct'].astype(str)) + '%'
        detail_table = temp_data[['UNMH_Cost_Center','Department', 'Name', 'Title', 'Degree', 'Certification']]
        detail_table = detail_table.to_html()
        edu_table = edu_totals.to_html()

        cert_totals = temp_data['Certification'].value_counts()
        cert_totals =pd.DataFrame(cert_totals)
        cert_totals['Certification'] = (cert_totals['Certification']/cert_totals['Certification'].sum()) * 100
        cert_totals['Certification'] = cert_totals['Certification'].round(1)
        cert_totals['Certification'] = (cert_totals['Certification'].astype(str)) + '%'
        cert_table = cert_totals.to_html()

        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "FYI - NDNQI RN Education Data " + date_month + ' ' + unit_name
        newMail.To = email
        newMail.Cc = 'NDNQI@salud.unm.edu'

        greeting = '''\
        <html> 
            <head> 
                <font size='4'> 
                Hello Unit Director,<br><br> This is an automated message from your NDNQI team at Nursing 
                Clinical Informatics.  This message is for your information only, and no response is needed.  
                <br><br>
                Below you will find RN Education and Certification data as it appeared in the 
                Nurse Recognition Database on the final day of previous month.  We hope you will find this actionable data 
                that improves the transparency of our data reporting.    
                <br><br>
                Please be advised on just a few points: <br><br>
                1.  The the RN Education and Certification data below does not include RN positions which are excluded from
                NDNQI reporting of Education and Certification.  Current exclusion criteria can be found at
                https://members.nursingquality.org/ndnqiportal. If you find an employee is missing from the Nurse Recognition Database 
                after you have verified they entered their data, please submit an IT help ticket at https://help.health.unm.edu/CherwellPortal/. 
                <br><br>
                2.  RN Education and Certification data contained in the Nurse Recognition Database on the last day of the quarter
                is reported per NDNQI standards.  Any updates that are captured between now and the last day
                of the quarter will be included in the data we report to NDNQI.
                <br><br>
                As always, we welcome any questions or feedback. <br><br>
                Thank You,<br>
                Barry Brooks, RN <br>
                Ethan Mooney, RN, MSN, PCCN<br>
                Hospital NDNQI Team<br>
                Nursing Clinical Informatics<br><br><br>
               </font>
            </head>
            <body><font size='4'>RN Education Summary:</font></body>            
        <html>
        '''

        cert_title = '''\
        <html>
            <head>
                <font size='4'><br><br> RN Certification Summary:</font>
            <head>
        <html>
        '''

        details_title = '''\
        <html>
            <head>
                <font size='4'><br><br>Education and Certification Details</font>
            <head>
        <html>
        '''

        html = greeting + edu_table + cert_title + cert_table + details_title + detail_table

        newMail.HTMLBody = html
        newMail.Send()
        x = x + 1



