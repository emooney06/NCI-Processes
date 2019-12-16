import pandas as pd
from my_variables import mmm_dict
from my_variables import master_alias

# this function maximized display of pandas dataframes in the immediate window
def max_pd_display():
    pd.option_context('display.max_rows', None, 'display.max_columns', None)  # more options can be specified also
    pd.options.display.max_colwidth = 199
    pd.options.display.max_columns = 1000
    pd.options.display.width = 1000
    pd.options.display.precision = 2  # set as needed

# this function takes a 9-digit integer cost center and converts it to a 5-character string
def make_string_cost_center(df):
    df = df.dropna()
    df = df.applymap(str)
    df['UNMH_Cost_Center'] = df['UNMH_Cost_Center'].str[0:5]
    return df

# this function returns a boolean value based on user input 'answer'
def check_answer(myAnswer):
    if (
        myAnswer == 'y' 
        or myAnswer == 'Y'
        or myAnswer == 'yes'
        or myAnswer == 'Yes'
        or myAnswer == 'YES'
        or myAnswer == 'YE'
        or myAnswer == 'Ye'
        or myAnswer == 'ye'
        or myAnswer == 'yES'
        ):
        myBool = True
    elif (
        myAnswer == 'N'
        or myAnswer == 'n'
        or myAnswer == 'no'
        or myAnswer == 'No'
        or myAnswer == 'NO'
        or myAnswer == 'nO'):
        myBool = False
    else:
        print('I don\'t understand your answer; please close this window, try again, and input a normal response next time newbie!')
        quit()
    return myBool;

# this function returns the number of days in a given month passed in mm format
def getNumDaysInMonth(month_numStr):
    if (month_numStr == '02'):
        numDays = 28
    elif (
        month_numStr == '04'
        or month_numStr == '06'
        or month_numStr == '09'
        or month_numStr == '11'
        ):
        numDays = 30
    else:
        numDays = 31

    return numDays;

# this function double checks if an answer that was given is really the intended answer; use this for things like: \
#  "are you sure you want to send a kazillion emails?"
def double_check(my_answer):
    if my_answer == True:
        double_check = str(input('Are you sure you want to send 100+ emails? \n Please input \'yes\' or \'no\', or type \'exit()\' to quit.'))
        if double_check == 'no':
            my_answer = False
        else:
            my_answer = True
    return my_answer

# Pass a dataframe and a "measure long description" to this function and it takes the df with 
# NDNQI_Reporting_Unit_Names and a value, then joins (on NDNQI_Reporting_Unit_Name) the other required columns from
# MasterAlias to fit the shape of the monthly report.  Of note: this uses the return from FileDateVars Class and the 
# mmm_dict (from my_variables) to insert the time(date) in the yyyy_mmm format.
def add_columns_for_reporting(df, measure_long_desc_str, year_str, numeric_month_str):
    df.insert(0, 'Measure_Long_Description', measure_long_desc_str)
    df.insert(0, 'Time', year_str + '_' + mmm_dict[numeric_month_str])
    df = pd.merge(df, master_alias[['NDNQI_Reporting_Unit_Name', 'NDNQI_Unit_Type_Monthly', 'Executive_Director']], on='NDNQI_Reporting_Unit_Name')
    #reorder the columns
    df = df[['NDNQI_Reporting_Unit_Name', 'Executive_Director', 'NDNQI_Unit_Type_Monthly', 'Measure_Long_Description', 'Time', 'Unit_Score']]
    # round to one decimal
    df['Unit_Score'] = df['Unit_Score'].round(1)
    df = df.drop_duplicates()
    return df;

## this function returns the int representing the quarter as a string (so it is ready to use in file paths and names
#def getQtrStr(month_string):
#    if (
#        month_string == 'Jan'
#        or month_string == 'JAN'
#        or month_string == 'jan'
#        or month_string == 'jAN'
#        or month_string == 'January'
#        or month_string == 'january'
#        or month_string == 'JANUARY'
#        or month_string == 'jANUARY'
#        or month_string == '01'
#        or month_string == '1'
#        ): 
#        qtrStr = '1'
#    elif (
#        month_string == 'Feb'
#        or month_string == 'FEB'
#        or month_string == 'feb'
#        or month_string == 'fEB'
#        or month_string == 'February'
#        or month_string == 'february'
#        or month_string == 'FEBRUARY'
#        or month_string == 'fEBRUARY'
#        or month_string == '02'
#        or month_string == '2'
#        ):
#        qtrStr = '1'
#    elif (
#        month_string == 'Mar'
#        or month_string == 'MAR'
#        or month_string == 'mar'
#        or month_string == 'mAR'
#        or month_string == 'March'
#        or month_string == 'march'
#        or month_string == 'MARCH'
#        or month_string == 'mARCH'
#        or month_string == '03'
#        or month_string == '3'
#        ):
#        qtrStr = '1'
#    elif (
#        month_string == 'Apr'
#        or month_string == 'APR'
#        or month_string == 'apr'
#        or month_string == 'aPR'
#        or month_string == 'April'
#        or month_string == 'april'
#        or month_string == 'APRIL'
#        or month_string == 'aPRIL'
#        or month_string == '04'
#        or month_string == '4'
#        ):
#        qtrStr = '2'
#    elif (
#        month_string == 'May'
#        or month_string == 'MAY'
#        or month_string == 'may'
#        or month_string == 'mAY'
#        or month_string == '05'
#        or month_string == '5'
#        ):
#        qtrStr = '2'
#    elif (
#        month_string == 'Jun'
#        or month_string == 'JUN'
#        or month_string == 'jun'
#        or month_string == 'jUN'
#        or month_string == 'June'
#        or month_string == 'june'
#        or month_string == 'JUNE'
#        or month_string == 'jUNE'
#        or month_string == '06'
#        or month_string == '6'
#        ):
#        qtrStr = '2'
#    elif (
#        month_string == 'Jul'
#        or month_string == 'JUL'
#        or month_string == 'jul'
#        or month_string == 'jUL'
#        or month_string == 'July'
#        or month_string == 'july'
#        or month_string == 'JULY'
#        or month_string == 'jULY'
#        or month_string == '07'
#        or month_string == '7'
#        ):
#        qtrStr = '3'
#    elif (
#        month_string == 'Aug'
#        or month_string == 'AUG'
#        or month_string == 'aug'
#        or month_string == 'aUG'
#        or month_string == 'August'
#        or month_string == 'august'
#        or month_string == 'AUGUST'
#        or month_string == 'aUGUST'
#        or month_string == '08'
#        or month_string == '8'
#        ):
#        qtrStr = '3'
#    elif (
#        month_string == 'Sep'
#        or month_string == 'SEP'
#        or month_string == 'sep'
#        or month_string == 'sEP'
#        or month_string == 'September'
#        or month_string == 'september'
#        or month_string == 'SEPTEMBER'
#        or month_string == 'sEPTEMBER'
#        or month_string == '09'
#        or month_string == '9'
#        ):
#        qtrStr = '3'
#    elif (
#        month_string == 'Oct'
#        or month_string == 'OCT'
#        or month_string == 'oct'
#        or month_string == 'oCT'
#        or month_string == 'October'
#        or month_string == 'october'
#        or month_string == 'OCTOBER'
#        or month_string == 'oCTOBER'
#        or month_string == '10'
#        ):
#        qtrStr = '4'
#    elif (
#        month_string == 'Nov'
#        or month_string == 'NOV'
#        or month_string == 'nov'
#        or month_string == 'nOV'
#        or month_string == 'November'
#        or month_string == 'november'
#        or month_string == 'NOVEMBER'
#        or month_string == 'nOVEMBER'
#        or month_string == '11'
#        ):
#        qtrStr = '4'
#    elif (
#        month_string == 'Dec'
#        or month_string == 'DEC'
#        or month_string == 'dec'
#        or month_string == 'dEC'
#        or month_string == 'December'
#        or month_string == 'december'
#        or month_string == 'DECEMBER'
#        or month_string == 'dECEMBER'
#        or month_string == '12'
#        ):
#        qtrStr = '4'
    
#    return qtrStr;

## this function returns the int representing the month as a string (so it is ready to use in file paths and names
#def getMonthNumStr(month_string):
#     if (
#        month_string == 'Jan'
#        or month_string == 'JAN'
#        or month_string == 'jan'
#        or month_string == 'jAN'
#        or month_string == 'January'
#        or month_string == 'january'
#        or month_string == 'JANUARY'
#        or month_string == 'jANUARY'
#        or month_string == '01'
#        or month_string == '1'
#        ): 
#        monthNumStr = '01'
#     elif (
#        month_string == 'Feb'
#        or month_string == 'FEB'
#        or month_string == 'feb'
#        or month_string == 'fEB'
#        or month_string == 'February'
#        or month_string == 'february'
#        or month_string == 'FEBRUARY'
#        or month_string == 'fEBRUARY'
#        or month_string == '02'
#        or month_string == '2'
#        ):
#        monthNumStr = '02'
#     elif (
#        month_string == 'Mar'
#        or month_string == 'MAR'
#        or month_string == 'mar'
#        or month_string == 'mAR'
#        or month_string == 'March'
#        or month_string == 'march'
#        or month_string == 'MARCH'
#        or month_string == 'mARCH'
#        or month_string == '03'
#        or month_string == '3'
#        ):
#        monthNumStr = '03'
#     elif (
#        month_string == 'Apr'
#        or month_string == 'APR'
#        or month_string == 'apr'
#        or month_string == 'aPR'
#        or month_string == 'April'
#        or month_string == 'april'
#        or month_string == 'APRIL'
#        or month_string == 'aPRIL'
#        or month_string == '04'
#        or month_string == '4'
#        ):
#        monthNumStr = '04'
#     elif (
#        month_string == 'May'
#        or month_string == 'MAY'
#        or month_string == 'may'
#        or month_string == 'mAY'
#        or month_string == '05'
#        or month_string == '5'
#        ):
#        monthNumStr = '05'
#     elif (
#        month_string == 'Jun'
#        or month_string == 'JUN'
#        or month_string == 'jun'
#        or month_string == 'jUN'
#        or month_string == 'June'
#        or month_string == 'june'
#        or month_string == 'JUNE'
#        or month_string == 'jUNE'
#        or month_string == '06'
#        or month_string == '6'
#        ):
#        monthNumStr = '06'
#     elif (
#        month_string == 'Jul'
#        or month_string == 'JUL'
#        or month_string == 'jul'
#        or month_string == 'jUL'
#        or month_string == 'July'
#        or month_string == 'july'
#        or month_string == 'JULY'
#        or month_string == 'jULY'
#        or month_string == '07'
#        or month_string == '7'
#        ):
#        monthNumStr = '07'
#     elif (
#        month_string == 'Aug'
#        or month_string == 'AUG'
#        or month_string == 'aug'
#        or month_string == 'aUG'
#        or month_string == 'August'
#        or month_string == 'august'
#        or month_string == 'AUGUST'
#        or month_string == 'aUGUST'
#        or month_string == '08'
#        or month_string == '8'
#        ):
#        monthNumStr = '08'
#     elif (
#        month_string == 'Sep'
#        or month_string == 'SEP'
#        or month_string == 'sep'
#        or month_string == 'sEP'
#        or month_string == 'September'
#        or month_string == 'september'
#        or month_string == 'SEPTEMBER'
#        or month_string == 'sEPTEMBER'
#        or month_string == '09'
#        or month_string == '9'
#        ):
#        monthNumStr = '09'
#     elif (
#        month_string == 'Oct'
#        or month_string == 'OCT'
#        or month_string == 'oct'
#        or month_string == 'oCT'
#        or month_string == 'October'
#        or month_string == 'october'
#        or month_string == 'OCTOBER'
#        or month_string == 'oCTOBER'
#        or month_string == '10'
#        ):
#        monthNumStr = '10'
#     elif (
#        month_string == 'Nov'
#        or month_string == 'NOV'
#        or month_string == 'nov'
#        or month_string == 'nOV'
#        or month_string == 'November'
#        or month_string == 'november'
#        or month_string == 'NOVEMBER'
#        or month_string == 'nOVEMBER'
#        or month_string == '11'
#        ):
#        monthNumStr = '11'
#     elif (
#        month_string == 'Dec'
#        or month_string == 'DEC'
#        or month_string == 'dec'
#        or month_string == 'dEC'
#        or month_string == 'December'
#        or month_string == 'december'
#        or month_string == 'DECEMBER'
#        or month_string == 'dECEMBER'
#        or month_string == '12'
#        ):
#        monthNumStr = '12'
    
#     return monthNumStr;
