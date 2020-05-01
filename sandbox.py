import sqlalchemy
from sqlalchemy import Table, Column, Integer, String, MetaData
import os
import pandas as pd
from pathlib import Path
from my_functions import max_pd_display
import datetime as dt
from sqlalchemy.types import BigInteger
from sqlalchemy.types import VARCHAR
from sqlalchemy.types import DateTime

meta = MetaData()
max_pd_display()

sql_all = " \
    SELECT * \
    FROM scd_stats" 


# define a timestamp variable
timestamp = dt.datetime.now()
#define new variable of the left 10 timestamp digits to modify file names showing date they were uploaded to database
timestamp_left10 = str(timestamp)[:10]

engine = sqlalchemy.create_engine('mssql+pymssql://unmmg-epide/PULSE')


engine.execute(sql_all)

df = pd.read_sql_table('ndnqi_raw_export', engine)

raw_folder = Path('K:/ClinicalAdvisoryTeam/data_folders/8370 - anticoag')
file_string = '2020-04.xlsx'

df_to_upload = raw_folder / file_string
df_to_upload = pd.read_excel(df_to_upload)
def wrangle_upload_columns(df):
    # replace any " - " with " " (spaces) so they can be removed    
    df.columns = df.columns.str.replace(' - ', ' ')
    # replace " " with "_" to work better with sql statements
    df.columns = df.columns.str.replace(' ', '_')

    # add a timestamp column to record when the upload transaction takes place
    df['upload_timestamp'] = timestamp

wrangle_upload_columns(df_to_upload)

scd_documentation = Table(
    'scd_stats', meta,
    Column('Clinical_Event_Id', BigInteger),
    Column('Action_Personnel', VARCHAR),
    Column('Person_Location-_Nurse_Unit_(Curr)', VARCHAR),
    Column('Clinical_Event_End_Date_&_Time', DateTime),
    Column('Clinical_Event_Result', VARCHAR),
    Column('Financial_Number', BigInteger),
    Column('upload_timestamp', DateTime)
    )

meta.create_all(engine)

df_to_upload.to_sql('scd_stats', con=engine, chunksize=10, schema='NI', if_exists='append')


,
                    schema='NI', dtype={'Clinical_Event_Id': BigInteger,
                                        'Action_Personnel': VARCHAR,
                                        'Person_Location-_Nurse_Unit_(Curr)': VARCHAR,
                                        'Clinical_Event_End_Date_&_Time': DateTime,
                                        'Clinical_Event_Result': VARCHAR,
                                        'Financial_Number': BigInteger,
                                        'upload_timestamp': DateTime})


engine.execute(sql_all)





"""
Date:  2019-10-18
Author:  Ethan Mooney
Description:  This function takes NDNQI raw output files (specified from data provided in user input) from 
the raw_file variable and uploads them to the unmmg-epide/PULSE database.  Before it uploads data to 
ndnqi_raw_export table, it removes the previous quarter data to allow the most accurate (retrod) data to be 
included in the table.  It then takes the current quarter and the previous quarter and adds it to the table.
"""

#set the pandas disply options to view more complete data set - this is useful for debugging
def max_pd_display_options():
    pd.option_context('display.max_rows', None, 'display.max_columns', None)  # more options can be specified also
    pd.options.display.max_colwidth = 199
    pd.options.display.max_columns = 1000
    pd.options.display.width = 1000
    pd.options.display.precision = 2  # set as needed

#print the entire contents of the updated table to excel for exploratory or debugging purposes
def new_table_to_excel():
    print('reading table for excel output')
    #read the sql query return to a dataframe
    test_table = pd.read_sql_query(sql_all, con=engine)
    print('writing to excel')
    # write the table to an excel file
    test_table.to_excel(raw_file + 'test_out.xlsx')

#make column headers in a format that works well with pandas
def wrangle_upload_columns():
    # replace any " - " with " " (spaces) so they can be removed    
    df_to_upload.columns = df_to_upload.columns.str.replace(' - ', ' ')
    # replace " " with "_" to work better with sql statements
    df_to_upload.columns = df_to_upload.columns.str.replace(' ', '_')

    # add a timestamp column to record when the upload transaction takes place
    df_to_upload['upload_timestamp'] = timestamp

#define the butt-load of variables I used for this function
def define_variables():
    year = input('What is the year of your data (ex: 2019)?')
    qtr = input('What is the quarter of your data(ex: 2)?')
    # define the path that all the files to upload are in
    raw_file = r'J:\NDNQI\Data Reporting\NDNQI Raw Output Files\\'
    # define dictionary to return previous quarter of qtr input string
    qtr_dict = {'1':4, '2':1, '3':2, '4':3}
    # if statement to correct year of previous quarter in the case the input qtr is 1
    if qtr == '1':
        last_qtr_year = int(year) - 1
    else:
        last_qtr_year = int(year)
    # define the last quarter (qtr)
    last_qtr = str(qtr_dict[qtr])
    # concat the last year and last quarter as an int type variable
    last_qtr_year_and_qtr_int = int(str(last_qtr_year) + last_qtr)
    #concat the input year and qtr as a string variable used in file names
    year_and_qtr_str = year + '_Q' + qtr
    # concat the year and qtr then make it an int for use in sql queries
    year_and_qtr_int = int(year + qtr)
    # define a timestamp variable
    timestamp = dt.datetime.now()
    #define new variable of the left 10 timestamp digits to modify file names showing date they were uploaded to database
    timestamp_left10 = str(timestamp)[:10]
    #sql statement to remove the data from the last quarter
    delete_last_quarter = str(" \
        DELETE \
        FROM ndnqi_raw_export" \
       " WHERE quarter = " + str(last_qtr_year_and_qtr_int))
    #sql query to select all records from the table (mostly useful in debugging but not bad to
    # have an excel copy if the table is not too big
    sql_all = " \
        SELECT * \
        FROM ndnqi_raw_export" 

   
max_pd_display_options()

#define the butt-load of variables I used for this function
define_variables()

# connect to database
print('connecting to database')
engine = sqlalchemy.create_engine('mssql+pymssql://unmmg-epide/PULSE')

#execute the sql to delete last quarter; the sql statement is defined in the variables function
print('removing previous quarter data (' + str(last_qtr_year_and_qtr_int) +')')
engine.execute(delete_last_quarter)


# append the table with the trimmed dataframe
print("loading file: " 'NDNQI Raw Output ' + year_and_qtr_str)
df_to_upload.to_sql('ndnqi_raw_export', con=engine, chunksize=10, if_exists='append')
print('upload complete')
    



