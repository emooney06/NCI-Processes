############################################################################################################################
# Title: COVID-19 Risk Alert
# Date:  2020-05-18
# Author:  Ethan Mooney
# Description:  This function looks for a file that should be safed by the outlook rule "8940 Save and Move" running on the
# "server box".  If no file is found, it will repeat in 30 min.  If the file is found it processes it using a number of 
# filters, then generates unit-specific emails to notify unit directors of patients who are identified as potential 
# COVID-19 risks.  These are generallyt patients who have fallen throught the cracks for testing and/or screening.  
############################################################################################################################

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
import time 
import stdiomask
from exchangelib import DELEGATE, Account, Credentials, Configuration, FileAttachment, ItemAttachment, Message, CalendarItem, HTMLBody, Mailbox, FaultTolerance, HTMLBody

 
try: 
    ad_password = stdiomask.getpass(prompt= 'Enter Active Directory Password: ', mask='*') 
except Exception as error: 
    print('ERROR', error) 


#credentials are the domain name with username and password
creds = Credentials(username='health\\ejmooney', password=ad_password)
#account configuration
config = Configuration(server='HSCLink.health.unm.edu', credentials=creds, retry_policy=FaultTolerance(max_wait=3600))



#set the thresholds for testing and screening hours difference
hrs_screen_threshold = 36
hrs_test_threshold = 72

max_pd_display()

#define the file paths and file names0
data_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen')
support_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/support_files')
archive_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen/8940_archive')
file_name = '#8940 Covid Screen.xlsx'
while True:
    try:
        #create a timestamp for the archive file name
        timestr = time.strftime("%Y%m%d-%H%M_")
        #add the time stamp to the file name to create the archive file name string
        archive_file =  timestr + file_name 

        #read the file from the PI report
        df = pd.read_excel(data_path / file_name)
        #save the dataframe as an archive
        df.to_excel(archive_path / archive_file, index=False)
        # drop duplicates
        df = df.drop_duplicates()
    
        #create the list of units that should be filetered out of the dataframe (MBU will be added back later, it's filtered here so it can be manipulated separately
        filter_by = ['P ICN-3 (IN3P)', 'P ICN-4 (IN4P)', 'P NBICU (NBIP)', 'P NB Nursery (NBNP)', 'P Mother-Baby (MBUP)', 
                        'P Admit Prep (APIP)', 'P MCICU (MCIP)', 'P OB Spec Care (HRMP)', 'P E-D Inpt (EDIP)', 'P Ped ICU (PICP)']
        #create a timedelta column for screening to admit time to capture those screened outside the threshold for screening recency
        df['admt_scrn_diff'] = (df['admit_dt_tm'] - df['screen_dt_tm']).astype('timedelta64[h]')
        df['admt_test_diff'] = (df['admit_dt_tm'] - df['testing_dt_tm']).astype('timedelta64[h]')

        #pull mbu into a separte df 
        mbu_df = df[(df.location == 'P Mother-Baby (MBUP)')]
        #eliminate the rooms that babies will be in    
        mbu_df = mbu_df[(mbu_df.location_bed != '1B')]

        #filter out the list of units from the dataframe
        df = df[~df.location.isin(filter_by)]
        #add mbu back to the dataframe (without beds that babies will be in)
        df = df.append(mbu_df)
        #filter out patients who have screened negative
        pos_scrn_df = df[(df['exposure_result'] != 'No exposure risk') &
                            (df['symptoms_result'] != 'None of the above') | ((df['admt_scrn_diff'] > hrs_screen_threshold) | (df['admt_scrn_diff'] == None))]
        #filter using another set of criteria to accomodate different variations of the screening form
        pos_scrn_df = pos_scrn_df[(pos_scrn_df['exposure_result'] != 'The patient has had no close contact') &
                            (pos_scrn_df['symptoms_result'] != 'None of the above') | ((pos_scrn_df['admt_scrn_diff'] > hrs_screen_threshold) | (pos_scrn_df['admt_scrn_diff'] == None))]

        #filter out patients who have results for COVID-19 test
        pos_scrn_not_neg_test_df = pos_scrn_df[(pos_scrn_df['testing_result'] != 'Not detected') &
                                                (pos_scrn_df['testing_result'] != 'Detected')  | ((pos_scrn_df['admt_test_diff'] > hrs_test_threshold) | (pos_scrn_df['admt_test_diff'] == None))] 

        pos_scrn_not_neg_test_df = pos_scrn_not_neg_test_df[(pos_scrn_not_neg_test_df['outside_result'] != 'Yes')] 

        ##################################################################################################################################################
        #This section can easily be modified to generate a file of fins that have been reported already.  This could feasibly be modified to enact limits 
        #on reporting such as report each patient only 2 times in each unit.  This would be necessary if we moved closer to a real-time reporting model.
        ###################################################################################################################################################
        #pos_scrn_not_neg_test_df['loc_fin'] = pos_scrn_not_neg_test_df['Financial Number'].astype(str) + '_' + pos_scrn_not_neg_test_df['location']
        #prev_reported_df = pd.read_excel(data_path / 'reported_fins.xlsx')
        #prev_reported_list = prev_reported_df['loc_fin'].to_list()
        #pos_scrn_not_neg_test_df = pos_scrn_not_neg_test_df[~pos_scrn_not_neg_test_df.loc_fin.isin(prev_reported_list)]
        #reported_df = pos_scrn_not_neg_test_df[['Financial Number', 'location']]
        #reported_df['loc_fin'] = reported_df['Financial Number'].astype(str) + '_' + reported_df['location']
        #prev_reported_df = prev_reported_df.append(reported_df, sort=True)
        #prev_reported_df.to_excel(data_path / 'reported_fins.xlsx', index=False)
        ##################################################################################################################################################

        #get rid of the NaN values (null values) and replace with "no results found string"
        pos_scrn_not_neg_test_df['careset_order'] = pos_scrn_not_neg_test_df['careset_order'].replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['testing_result'] = pos_scrn_not_neg_test_df['testing_result'].replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['exposure_result'] = pos_scrn_not_neg_test_df['exposure_result'].replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['symptoms_result'] = pos_scrn_not_neg_test_df['symptoms_result'].replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['outside_result'] = pos_scrn_not_neg_test_df['outside_result'].replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['outside_result_dt_tm'] = pos_scrn_not_neg_test_df['outside_result_dt_tm'].replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['admt_scrn_diff'] = pos_scrn_not_neg_test_df['admt_scrn_diff'].replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['admt_test_diff'] = pos_scrn_not_neg_test_df['admt_test_diff'].replace(np.nan, 'no results found', regex=True)

        #get rid of the "naT" values, which are blank (null) date/time values - replace these values with the string "no results found"
        pos_scrn_not_neg_test_df['careset_order_dt_tm'] = pos_scrn_not_neg_test_df['careset_order_dt_tm'].astype(object).where(pos_scrn_not_neg_test_df.careset_order_dt_tm.notnull(), 'no results found')
        pos_scrn_not_neg_test_df['testing_dt_tm'] = pos_scrn_not_neg_test_df['testing_dt_tm'].astype(object).where(pos_scrn_not_neg_test_df.testing_dt_tm.notnull(), 'no results found')
        pos_scrn_not_neg_test_df['screen_dt_tm'] = pos_scrn_not_neg_test_df['screen_dt_tm'].astype(object).where(pos_scrn_not_neg_test_df.screen_dt_tm.notnull(), 'no results found')
        #remove the columns that are not needed
        pos_scrn_not_neg_test_df = pos_scrn_not_neg_test_df[['MRN- Organization', 'location', 'admit_dt_tm', 'location_room', 'location_bed', 'careset_order', 'careset_order_dt_tm',
        'testing_result', 'testing_dt_tm', 'exposure_result', 'symptoms_result', 'screen_dt_tm', 'outside_result', 'outside_result_dt_tm', 'admt_scrn_diff', 'admt_test_diff', 'report_time']]
        #change the column names to be more human-readable
        pos_scrn_not_neg_test_df = pos_scrn_not_neg_test_df.rename(columns={"MRN- Organization": "MRN", "admit_dt_tm": "Admit Date", "location_room": "Room", "location_bed": "Bed", "careset_order": "Order", 
                                                    "careset_order_dt_tm": "Order Date", "testing_result": "Test Result","testing_dt_tm": "Test Date", "exposure_result": "Exposure", "symptoms_result": "Symptoms", 
                                                    "screen_dt_tm": "Screen Date", "outside_result": "OSH Result", "outside_result_dt_tm": "OSH Result Date", "admt_scrn_diff": "Admit-Screen Hrs",
                                                    "admt_test_diff": "Admit-Test Hrs", "report_time": "Report Time"})
        #Note*** this is a copy of the Master Alias because normally reported units like CTH inpatient and Peds PACU are under ICU directors
        ma_df = pd.read_excel(support_path / 'ma_copy.xlsx')
        #limit the data set to only the columns needed
        ma_df = ma_df[['cerner_unit_name', 'UD_Email']]
        #drop na values from ma
        ma_df = ma_df.dropna()
        #rename the columns to match the cerner location column
        ma_df = ma_df.rename(columns={'cerner_unit_name': 'location'})
        #create a dictionary to map the UD emails to their unit locations
        email_dict = ma_df.set_index('location')['UD_Email'].to_dict()
        #create a list of locations to iterate through
        location_list = pos_scrn_not_neg_test_df.location.unique()
        #generate a global object to be accessible across functions  - I don't think this is necessary any longer... not sure why I defined it global
        global unit_table
        #loop through each location in the list of locations
        for location in location_list:
            #try to get the email address            
            try:
                # initialize variable that is email address which is a found from the email dictionary object
                email = email_dict.get(location)
                #if there is no email associated with the unit, print an error - this should not happen
                if email == None:
                    email = 'no email found for this unit: ' + location
            except KeyError:
                # If there is no UD Email associated with that cost center, just print it to console
                print('There is no email address for: ') 
                print(email)
            #filter the dataframe of non-negative screens with non-negative test results by unit location
            unit_df = pos_scrn_not_neg_test_df[(pos_scrn_not_neg_test_df['location']) == location]
            #create the instance of account class object
            a = Account('nci@salud.unm.edu', credentials=creds, autodiscover=True)
            #convert it to a dataframe to embed in email
            unit_table = unit_df.to_html(index=False)

            greeting = '''\
            <html> 
                <head> 
                    <font size='3'> 
                    Hello Unit Director,<br><br> This is an automated message from your Nursing 
                    Clinical Informatics team.  This message is for your information only - no response is needed.  
                    <br><br>
                    Below you will find patients identified by our algorithm as a potential COVID-19 exposure risk because they have not screened negative
                    within 36 hours of admission and/or do not have COVID-19 test results within 72 hours of admission. 
                    <br><br>
                    If you find a patient who needs COVID-19 testing, you may inform the provider that testing can be ordered via the "COVID-19 Test 
                    careset".  If you find a patient who should be screened, the screening can be found in the ad-hoc form titled "Infectious Disease 
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

            disclaimer = '''\
            <html> 
                <head> 
                    <font size='2'> 
                    Produced by:  UNMH Nursing Clinical Informatics<br>
                    This material is produced in connection with, and for the purpose of the Patient Safety Evaluation System
                    and-or Review Organization established at the University of New Mexico Hospital, and is therefore confidential 
                    Patient Safety Work Product (“PSWP”) and/or confidential peer review material of the University of New Mexico Hospital 
                    as defined in 42 C.F.R. subsection 3.20 and-or the Review Organizations Immunity Act, Section 41-9-1 et seq., NMSA 1978 
                    as amended (ROIA).  As such, it is confidential and is protected under federal law 42 C.F.R. subsection3.206 and/or 
                    ROIA.  Unauthorized disclosure of this document, enclosures thereto, and information therefrom is strictly prohibited.
                    </font>
                </head>            
            <html>
            '''
            #add the parts of the mail message
            html = greeting + unit_table + disclaimer
            m = Message(
            account=a,
            subject='FYI - Possible COVID-19 Risk *Secure*',
            body= HTMLBody(html),
            to_recipients=[
                Mailbox(email_address=email), # will be email variable
            ],
            cc_recipients=[],  # Simple strings work, too
            bcc_recipients=[],  # Or a mix of both
            )
            #################################################################
            # to implement this in production, change .Display to .Send
            #################################################################
            m.send()
        ##################################################################################################################################################
        ##delete the original file to prevent re-running the script on an outdated file if the process to drop the file in a folder errors
        os.remove(data_path / file_name)
        ##################################################################################################################################################
        print(str(timestr) + ' COVID-19 Risk Emails Sent')

    ####################################################################################################################################################
    except Exception as e:
        print(e)
        print(str(timestr) + ' file not found; 30 min sleep')
    time.sleep(1800)
    ####################################################################################################################################################


