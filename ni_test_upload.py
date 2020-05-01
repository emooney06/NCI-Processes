import sqlalchemy
import os
import pandas as pd
import datetime as dt
from my_functions import max_pd_display

"""
Date:  2019-10-18
Author:  Ethan Mooney
Description:  This function takes NDNQI raw output files (specified from data provided in user input) from 
the raw_file variable and uploads them to the unmmg-epide/PULSE database.  Before it uploads data to 
ndnqi_raw_export table, it removes the previous quarter data to allow the most accurate (retrod) data to be 
included in the table.  It then takes the current quarter and the previous quarter and adds it to the table.
"""

##set the pandas disply options to view more complete data set - this is useful for debugging
#def max_pd_display_options():
#    pd.option_context('display.max_rows', None, 'display.max_columns', None)  # more options can be specified also
#    pd.options.display.max_colwidth = 199
#    pd.options.display.max_columns = 1000
#    pd.options.display.width = 1000
#    pd.options.display.precision = 2  # set as needed

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

#########################################################################################
# defining a butt-load of variables here:
#########################################################################################
year = input('What is the year of your data (ex: 2019)?')
qtr = input('What is the quarter of your data(ex: 2)?')
# define the path that all the files to upload are in
raw_file = r'K:\\NDNQI\Data Reporting\NDNQI Raw Output Files\\'
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

# read the excel file from the raw file path with the year and quarter as named in folder
df_to_upload = pd.read_excel(raw_file + 'NDNQI Raw Output ' + year_and_qtr_str + '.xlsx')
#limit the dataframe to upload to only the current quarter and the last quarter (to capture retro data)
df_to_upload = df_to_upload[df_to_upload['Quarter'].isin([year_and_qtr_int, last_qtr_year_and_qtr_int])]


max_pd_display()

# connect to database
print('connecting to database')
engine = sqlalchemy.create_engine('mssql+pymssql://unmmg-epide/PULSE')

#execute the sql to delete last quarter; the sql statement is defined in the variables function
print('removing previous quarter data (' + str(last_qtr_year_and_qtr_int) +')')
engine.execute(delete_last_quarter)

#make column headers to a format that works well with pandas
wrangle_upload_columns()

# append the table with the trimmed dataframe
print("loading file: " 'NDNQI Raw Output ' + year_and_qtr_str)
df_to_upload.to_sql('ndnqi_raw_export', con=engine, chunksize=10, if_exists='append')
print('upload complete')
    
#rename the files so you can tell which files have been uploaded to the database and when
os.rename(raw_file + 'NDNQI Raw Output ' + year_and_qtr_str + '.xlsx', raw_file + 'NDNQI Raw Output ' + year_and_qtr_str + '_upload_' + timestamp_left10 + '.xlsx')

#print the entire contents of the updated table to excel for exploratory or debugging purposes
#new_table_to_excel_()

