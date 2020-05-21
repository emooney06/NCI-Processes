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
import win32com.client
import time 
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from google import auth
from oauth2client.client import GoogleCredentials


#authorize the google drive access
gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)


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
        pos_scrn_df = df[(df['exposure_result'] != 'No high exposure risk') &
                            (df['symptoms_result'] != 'No high risk symptoms') | ((df['admt_scrn_diff'] > 36) | (df['admt_scrn_diff'] == None))]
        
        pos_scrn_df2 = pos_scrn_df[(pos_scrn_df['exposure_result'] != 'The patient has had no close contact') &
                            (pos_scrn_df['symptoms_result'] != 'No high risk symptoms') | ((pos_scrn_df['admt_scrn_diff'] > 36) | (pos_scrn_df['admt_scrn_diff'] == None))]

        #filter out patients who have results for COVID-19 test
        pos_scrn_not_neg_test_df = pos_scrn_df2[(pos_scrn_df2['testing_result'] != 'Not detected') &
                                                (pos_scrn_df2['testing_result'] != 'Detected')  | ((pos_scrn_df2['admt_test_diff'] > 72) | (pos_scrn_df2['admt_test_diff'] == None))] 

        pos_scrn_not_neg_test_df3 = pos_scrn_not_neg_test_df[(pos_scrn_not_neg_test_df['outside_result'] != 'Yes')] 

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
        pos_scrn_not_neg_test_df['careset_order'] = pos_scrn_not_neg_test_df.careset_order.replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['testing_result'] = pos_scrn_not_neg_test_df.testing_result.replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['exposure_result'] = pos_scrn_not_neg_test_df.exposure_result.replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['symptoms_result'] = pos_scrn_not_neg_test_df.symptoms_result.replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['outside_result'] = pos_scrn_not_neg_test_df.outside_result.replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['outside_result_dt_tm'] = pos_scrn_not_neg_test_df.outside_result_dt_tm.replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['admt_scrn_diff'] = pos_scrn_not_neg_test_df.admt_scrn_diff.replace(np.nan, 'no results found', regex=True)
        pos_scrn_not_neg_test_df['admt_test_diff'] = pos_scrn_not_neg_test_df.admt_test_diff.replace(np.nan, 'no results found', regex=True)

        #get rid of the "naT" values, which are blank (null) date/time values - replace these values with the string "no results found"
        pos_scrn_not_neg_test_df['careset_order_dt_tm'] = pos_scrn_not_neg_test_df.careset_order_dt_tm.astype(object).where(pos_scrn_not_neg_test_df.careset_order_dt_tm.notnull(), 'no results found')
        pos_scrn_not_neg_test_df['testing_dt_tm'] = pos_scrn_not_neg_test_df.testing_dt_tm.astype(object).where(pos_scrn_not_neg_test_df.testing_dt_tm.notnull(), 'no results found')
        pos_scrn_not_neg_test_df['screen_dt_tm'] = pos_scrn_not_neg_test_df.screen_dt_tm.astype(object).where(pos_scrn_not_neg_test_df.screen_dt_tm.notnull(), 'no results found')
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

        #delete the original file to prevent re-running the script on an outdated file if the process to drop the file in a folder errors
        os.remove(data_path / file_name)
        #reate the mail object
        olMailItem = 0x0
        #create an object to interact with the outlook application
        obj = win32com.client.Dispatch("Outlook.Application")
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
            #convert it to a dataframe to embed in email
            unit_table = unit_df.to_html(index=False)
            #generate the email with all the parts
            newMail = obj.CreateItem(olMailItem)
            newMail.Subject = 'FYI - Possible COVID-19 Risk *Secure*'
            newMail.To = email
            newMail.Cc = 'ejmooney@salud.unm.edu'
            greeting = '''\
            <html> 
                <head> 
                    <font size='3'> 
                    Hello Unit Director,<br><br> This is an automated message from your Nursing 
                    Clinical Informatics team.  This message is for your information only - no response is needed.  
                    <br><br>
                    Below you will find patients identified by our algorithm as a potential COVID-19 exposure risk because they have not screened negative
                    and do not yet have current COVID-19 test results. 
                    <br><br>
                    If you find a patient who needs COVID-19 Testing, you may inform their provider that testing can be ordered via the "COVID-19 Test 
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
            newMail.HTMLBody = html
            #################################################################
            # to implement this in production, change .Display to .Send
            #################################################################
            newMail.Display()
            #generate a timestamp to write to a file in my google drive.  - the file is checked by my raspbery pi to ensure this function is still online
            timestr = time.strftime("%Y%m%d-%H%M")
            #access the file by the id
            file1 = drive.CreateFile({'id': '1U362h3YgTplBN6uNIWXQV9dq4i0Z7VrY'})
            #load the string as content to write
            file1.SetContentString(timestr)
            #write to the specified file
            file1.Upload() # Files.insert()
        print(str(timestr) + ' file found; emails sent and entering 30 min sleep')
        time.sleep(1800)
    except:
        #generate a timestamp to write to a file in my google drive.  - the file is checked by my raspbery pi to ensure this function is still online
        timestr = time.strftime("%Y%m%d-%H%M")
        #access the file by the id
        file1 = drive.CreateFile({'id': '1U362h3YgTplBN6uNIWXQV9dq4i0Z7VrY'})
        #load the string as content to write
        file1.SetContentString(timestr)
        #write to the specified file
        file1.Upload() # Files.insert()
        print(str(timestr) + ' file not found; 30 min sleep')
        time.sleep(1800)
    


