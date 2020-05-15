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

#define the file paths and file names0
data_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen')
archive_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen/8940_archive')
file_name = 'covid_screen.xlsx'

while True:
    try:
    #create a timestamp for the archive file name
        timestr = time.strftime("%Y%m%d-%H%M_")
        #add the time stamp to the file name to create the archive file name string
        archive_file =  timestr + file_name 
        print('attempting to read the file')
        #read the file from the PI report
        df = pd.read_excel(data_path / file_name)
        #save the dataframe as an archive
        df.to_excel(archive_path / archive_file, index=False)
        # drop duplicates
        df = df.drop_duplicates()
    except:
        print('executing the except statement')
        #sleep for 24 hours
        time.sleep(20)
    print("now i'm doing the stuff")
    time.sleep(30)
