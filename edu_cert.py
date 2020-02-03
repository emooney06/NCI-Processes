import pandas as pd
import os
from functools import reduce
import string
import numpy as np
import collections
import re
from datetime import datetime
import sys
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

''' 
Date: 2019-12-09
Author:  Ethan Mooney
Details:  This module takes input 
'''

display_email = False
max_pd_display()
fileVars = FileDateVars.from_inputs()
email_input = str(input('''Are you ready to send all the education emails to each Director?  Please know this is a \
butt-load of emails and is super embarrassing if you make a mistake! \nPlease input \'yes\' or \'no\', or type \'exit()\' to quit.'''))
email_input = check_answer(email_input)
send_email = double_check(email_input)
if email_input == False:
    display_email = str(input('''Would you like to display the first 10 emails as a test set - this will not 
    send emails, only display them in an outlook window.  \n Please input \'yes\' or \'no\', or type \'
    exit()\' to quit.'''))
    display_email = check_answer(display_email)



#def build_file_paths():
# Real path:  master_alias_path = Path('K:/NDNQI/MasterAliasRecord.xlsx')
#testing path
#master_alias_path = Path('C:/Users/ejmooney/Desktop/testData/MasterAlias_test.xlsx')
edu_job_codes_path = Path('K:/NDNQI/jobCodesMasterList.xlsm')
edu_path = Path('K:/NDNQI/SourceData/RN Education')
hr_path = Path('K:/NDNQI/SourceData/Nurse Turnover/HR Turnover Reports')
edu_path = Path('K:/NDNQI/SourceData/RN Education')
duplicates_path = Path('C:/Users/ejmooney/Desktop/testData/duplicates.xlsx')
reporting_path = Path('K:/NDNQI/Data Reporting/RN Education')

#define the folder named by year and quarter (ie 2019Q2)
folder_year_qtr = fileVars.year_str + 'Q' + fileVars.qtr_str 
# define strings for the file names
edu_file_str = 'RN Education ' + fileVars.year_str  + '-' + fileVars.numeric_month_str + '.xls'
hr_file_str = 'HR Report ' + fileVars.year_str + '-' + fileVars.numeric_month_str + '.xls'
# combine the file paths and the file strings
raw_edu_file = edu_path / folder_year_qtr/ edu_file_str
raw_hr_file = hr_path / hr_file_str

# map the available degrees to an integer value that acts as a bit-mask
eduDict = {
'Advanced Practice Nurse': 0,
'Associate Degree -Nursing': 10,
'Associate of Arts Degree - Non-Nursing': 0,
'Associate of Science Degree - Non-Nursing': 0,
'Bachelors of Art':	0,
'Bachelors of Bus Admin': 0,
'Bachelors Science': 0,
'Bachelors Science Nursing': 100,
'Certified Nurse Practitioner':	0,
'Clincal Nurse Specialist':	100,
'Diploma - RN':	1,
'Doctor of Education': 0,
'Doctor of Nursing Practice': 10000,
'Doctor of Nursing Practice - Candidate': 0,
'Doctor of Oriental Medicine': 0,
'Doctor of Oriental Medicine - Candidate': 0,
'Doctor of Philosophy - Nursing': 10000,
'Doctor of Philosophy - Other Discipline':	0,
'Doctor of Philosophy (nursing) - Candidate': 0,
'Doctor of Philosophy (other) - Candidate':	0,
'Licensed Practical Nurse':	0,
'Master Public Health Nursing':	1000,
'Masters Business Administration': 0,
'Masters Degree - Other': 0,
'Masters of Arts': 0,
'Masters Public Health': 0,
'Masters Science': 0,
'Masters Science Nursing': 1000,
'Nurse Practitioner Specialist': 1000,
'Other - Foreign RN Education':	0
}

# convert the summed integer to the highest degree achieved
def intToDegree(row):
    if 10 > row['eduScore'] > 0 :
        val = 'Diploma'
    elif 100 > row['eduScore'] > 9:
        val = 'ADN'
    elif 1000 > row['eduScore'] >99:
        val = 'BSN'
    elif 10000 > row['eduScore'] >999:
        val = 'MSN'
    elif 100000 > row['eduScore'] > 9999:
        val = 'PhD or DNP'
    elif row['eduScore'] > 99999:
        val = 'No Ed Record Found'
    else: 
        val = 'Nursing Foreign or Unknown Degree'
    return val;

# since degrees in progress do not matter, change all in progress to zero (0)
def inProgressCorrection(row):
    if row['Degree Category'] == 'In Progress':
        val = 0
    else: 
        val = 1
    return val;
# build more file paths
edu_data = pd.read_excel(edu_path / folder_year_qtr/ edu_file_str)
hr_data = pd.read_excel(hr_path / hr_file_str)
edu_job_codes = pd.read_excel(edu_job_codes_path, 'RN_Education')
#Need to change EEID to ID in HR report
hr_data = hr_data.rename(columns = {'EE ID':'ID'})
# keep only records that do not have a term date.
hr_data = hr_data[hr_data['Term date'] == ' ']
#reduce the df to columns that are needed
edu_data = edu_data[['ID', 'Degree Category', 'Description', 'Certification']]
# combine the education data and HR data
edu_hr_comb = reduce(lambda x,y: pd.merge(x,y, on='ID', how='left'), [hr_data, edu_data])
# Reduce the set of columns to those columns that are usefule
edu_hr_comb = edu_hr_comb[['Cost Center', 'ID', 'Department', 'Name', 'JOB', 'Title',  'Description', 'Degree Category', 'Certification']]
# convert the excluded job codes to a list to iterate over
exclusion_list = edu_job_codes['excl_job_codes'].tolist()

# loop through the df and remove records that are job codes on the exclusion list
for i in exclusion_list:
    edu_hr_comb = edu_hr_comb[edu_hr_comb.JOB != i]

# correct for "in progress status" by adding a column of 0 for in progress and 1 for all else
edu_hr_comb['multInProgr'] = edu_hr_comb.apply (lambda row: inProgressCorrection(row), axis=1)
# change the title cost center and ID to string to protect them from the sum function later truncated
edu_hr_comb[['Title', 'Cost Center', 'ID']] = edu_hr_comb[['Title', 'Cost Center', 'ID']].astype(str)
# convert the calendar to the integer (bit mask) based on the eduDictionary - this is the eduScore
edu_hr_comb['eduScore']= edu_hr_comb['Description'].map(eduDict)
# make any null value in the eduScore column a zero so when it is multiplied by the degree in progress indicator
#   the product will be zero (indicating no degree was specified.
edu_hr_comb['eduScore'] = edu_hr_comb['eduScore'].fillna(100000)
edu_hr_comb['Certification'] = edu_hr_comb['Certification'].fillna('N')
# multiply the eduScore and multInProg values to get the adjustment for degrees in progress
edu_hr_comb['eduScore'] = edu_hr_comb['eduScore'] * edu_hr_comb['multInProgr']
# group by rows that will be duplicate so the eduScore integer (bit mask) can be summed
edu_hr_comb = edu_hr_comb.groupby(['Name', 'Title', 'Cost Center','ID', 'JOB', 'Certification']).sum().reset_index()
# Reduce HR data to useful columns
hr_data = hr_data[['ID', 'Department']]
# make the ID a string so it can be merged with the other combined data
hr_data['ID'] = hr_data['ID'].astype(str)
#merge the combined data with hr data again to get the department
edu_hr_comb = reduce(lambda x,y: pd.merge(x,y, on='ID', how='left'), [edu_hr_comb, hr_data])
#edu_hr_comb2['eduScore'] = edu_hr_comb2['eduScore'].astype(int)
edu_hr_comb['corrEdu'] = edu_hr_comb.apply (lambda row: intToDegree(row), axis=1)
# Reduce the dataframe to columns that are needed now and re-order to make more readable
edu_hr_comb = edu_hr_comb[['Cost Center', 'Department', 'Name', 'Title', 'corrEdu', 'Certification']]
edu_hr_comb.columns = ['Cost Center', 'Department', 'Name', 'Title', 'Degree', 'Certification']
# make cost center a string so the left 4 digits can be truncated
edu_hr_comb['Cost Center'] = edu_hr_comb['Cost Center'].astype(str)
# truncate the cost center to the common 5-digit cost center
edu_hr_comb['Cost Center'] = edu_hr_comb['Cost Center'].str[4:]
# drop duplicates that are created having both expired and active certifications (keep last because 'y' is alphabetically after 'E'
edu_hr_comb = edu_hr_comb.drop_duplicates(['Cost Center', 'Name'], keep='last')
edu_hr_comb.loc[edu_hr_comb.Certification == 'N', 'Certification'] = 'No'
edu_hr_comb.loc[edu_hr_comb.Certification == 'Y', 'Certification'] = 'Yes'
edu_hr_comb.loc[edu_hr_comb.Certification == 'E', 'Certification'] = 'Expired'
ma_df = master_alias[['UNMH_Cost_Center','UD_Email', 'NCI_Standard_Name', 'NDNQI_Reporting_Unit_Name', \
    'NDNQI_Unit_Type_Monthly', 'Executive_Director']]
# convert the ma_df cost center into a 5-character string
ma_df = make_string_cost_center(ma_df)
edu_hr_comb = edu_hr_comb.rename(columns = {'Cost Center':'UNMH_Cost_Center'})
#########################################################################################################
# Key: Complete Dataframe:
# this is a complete data set containing both education and certificaton merged with masterAlias data
complete_df = pd.merge(edu_hr_comb, ma_df, on='UNMH_Cost_Center')
#########################################################################################################

# this function creates tables with the data needed for data entry to NDNQI
def create_tables(column_str):
    # reduce the dataset columns
    df = complete_df[['NDNQI_Reporting_Unit_Name', 'Name', column_str]] 
    df_pivot = df.pivot_table(df, index='NDNQI_Reporting_Unit_Name', columns=[column_str], aggfunc='count')
    df_pivot = df_pivot.fillna(0)
    #drop the second level so later merges are accurate
    df_pivot = df_pivot.droplevel(0, axis=1)
    return df_pivot;

#call the function to create the tables for degrees and for certifications
edu_pivot = create_tables('Degree')
cert_pivot = create_tables('Certification')
# re-order the columns for easier data entry
edu_pivot = edu_pivot[['Diploma', 'ADN','BSN', 'MSN', 'PhD or DNP', 'Nursing Foreign or Unknown Degree', 'No Ed Record Found']]
cert_pivot = cert_pivot[['Yes', 'No', 'Expired', 'No Expiration Listed']]
# build paths and Write the files to the RN Education folder for the quarter as certSummary and eduSummary + yyyy-mm
cert_summary_file_str = 'certSummary ' + fileVars.year_str + '-' + fileVars.numeric_month_str + '.xlsx'
cert_pivot.to_excel(edu_path / folder_year_qtr / cert_summary_file_str)
edu_summary_file_str = 'eduSummary ' + fileVars.year_str + '-' + fileVars.numeric_month_str + '.xlsx'
edu_pivot.to_excel(edu_path / folder_year_qtr / edu_summary_file_str)

# call the create_tables function to rebuild tables for monthly reporting
edu_pivot = create_tables('Degree')
cert_pivot = create_tables('Certification')

#add a column containing the math to produce % BSN or greater ((BSN + MSN + PHD/DNP) / (sum of all categories(ie total roster))
edu_pivot.insert(1, 'Unit_Score', ((edu_pivot['BSN'] + edu_pivot['MSN'] + edu_pivot['PhD or DNP']) / (edu_pivot['Diploma'] \
   + edu_pivot['ADN'] + edu_pivot['BSN'] + edu_pivot['MSN'] + edu_pivot['PhD or DNP'] +  \
   edu_pivot['Nursing Foreign or Unknown Degree'] + edu_pivot['No Ed Record Found'])) * 100)

#add a column with the math to determine a certification rate
cert_pivot.insert(1, 'Unit_Score', (cert_pivot['Yes'] / (cert_pivot['Expired'] + cert_pivot['No'] +  \
   cert_pivot['No Expiration Listed'] + cert_pivot['Yes'])) * 100)

# call the add_columns_for_reporting function from my_functions
edu_report = add_columns_for_reporting(edu_pivot, '% Direct Care RNs with BSN, MSN, or PhD', fileVars.year_str, fileVars.numeric_month_str)
cert_report = add_columns_for_reporting(cert_pivot, '% Direct Care RNs with Specialty Certification', fileVars.year_str, fileVars.numeric_month_str)

edu_report_file = 'Monthly Edu Reporting ' + fileVars.year_str  + '-' + fileVars.numeric_month_str + '.xlsx'
cert_report_file = 'Monthly Cert Reporting ' + fileVars.year_str  + '-' + fileVars.numeric_month_str + '.xlsx'

edu_report.to_excel(reporting_path / edu_report_file)
cert_report.to_excel(reporting_path / cert_report_file)

if send_email == False and display_email == False:  
    sys.exit("Your Monthly reporting file is in the data reporting folder and your data entry file is in the source data folder.  Well thank you, I think you\'re a pretty good tool too")

# initialize a time variable for current month/year for email text
global date_month
date_month = datetime.today().strftime('%Y-%m')
# main alias sheet from masterAliasRecord to a dataframe
# limit alias dataframe to cost center and UD Email
alias_df = master_alias[['UNMH_Cost_Center', 'UD_Email', 'NCI_Standard_Name']]
alias_df = alias_df.dropna()
alias_df['UNMH_Cost_Center'] = alias_df['UNMH_Cost_Center'].astype(str)
alias_df['UNMH_Cost_Center'] = alias_df['UNMH_Cost_Center'].str[:5]

# get the cost centers to iterate through from the alias dataframe; make it a list
cost_center_list = alias_df['UNMH_Cost_Center'].values.tolist()

# convert the alias record to a dictionary
email_dict = alias_df.set_index('UNMH_Cost_Center')['UD_Email'].to_dict()
#alias_df = master_alias[['UNMH_Cost_Center', 'NDNQI_Reporting_Unit_Name']]
unit_name_dict = alias_df.set_index('UNMH_Cost_Center')['NCI_Standard_Name'].to_dict()

global unit_name
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")

if send_email == True:
    print('Fire in the hole! \nThere goes your kazillion emails!  \n\nWell, the good news is that I\'ve included a copy of \
certSummary YYYY-MM and eduSummary YYYY-MM in the RN Education and Data Reporting folders.  Since we\'ve saved so \
much time together, now we have time for a fun fact:  A blue whale\'s heart is the size of a VW Beetle, and beats 5 \
times a miniute - even slower when they are on a really deep dive!  Have a great day and remember - You Nailed It!')
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

if display_email == True:
    x = 0
    #alias list is a list of all cost centers from MasterAliasRecord
    for line in cost_center_list[:5]:
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
        newMail.Display()
        x = x + 1
    print('Well here you go, just a little taste of what I\'m capable of... oh, and by the way I stashed your \
eduSummary and certSummary data in the RN Education and Data reporting folders.  Now have fun checking those \
beautiful emails I drafted for you.  And can I just say, You Really Nailed it Today Buddy!')

   