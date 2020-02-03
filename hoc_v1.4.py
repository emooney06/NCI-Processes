
from tika import parser
import os
import time
import csv
import pandas as pd
from time import sleep
from functools import reduce
import numpy as np
import sys
from my_functions import max_pd_display, getQtrStr, getMonthNumStr, check_answer, getNumDaysInMonth

max_pd_display()

yearStr = str(input(' Enter the year of the data you want:'))
monthStr = str(input(' Enter the month of the data you want:'))

answerInput = str(input(' Have you already updated the job description list and contractor names in the appropriate lists? \n Please input \'yes\' or \'no\', or type \'exit()\' to quit.'))

#print(answerInput)
checkedMatrix = check_answer(answerInput)

#print(checkedMatrix)

qtrStr = getQtrStr(monthStr)
monNumStr = getMonthNumStr(monthStr)
daysInMonth = getNumDaysInMonth(monNumStr)

# initialize the file paths
laborLevelsPDFPath = os.path.join('K:\\', 'NDNQI', 'SourceData', 'Hours Of Care', yearStr + 'Q' + qtrStr, 'Labor Levels ' + yearStr + '-' + monNumStr + '.pdf')
supportFilePath = os.path.join('K:\\', 'NDNQI', 'SourceData', 'Hours Of Care', 'hocSupportFile.xlsx')
checkLaborLevelsPath = os.path.join('K:\\', 'NDNQI', 'SourceData', 'Hours Of Care', yearStr + 'Q' + qtrStr, 'checkLaborLevels ' + yearStr + '-' + monNumStr + '.xlsx')
rawDataPath = os.path.join('K:\\', 'NDNQI', 'SourceData', 'Hours Of Care', yearStr + 'Q' + qtrStr, 'StaffRaw ' + yearStr + '-' + monNumStr + '.xls')
tmpTxtPath = os.path.join('K:\\', 'NDNQI', 'ndnqi_python', 'TmpFiles_DoNotDelete', 'tmpTxt.txt')
tmpCSVPath = os.path.join('K:\\', 'NDNQI', 'ndnqi_python', 'TmpFiles_DoNotDelete', 'pdfToCSV.csv')
print('\n Please be patient, sometimes the Tika server takes a nap and says she \'can\'t see the lame startup log message...\'\n I might have to wake her up... \n')
# use tika parser to read the pdf file
laborLevels = parser.from_file(laborLevelsPDFPath)
# Save the content of the pdf as raw
rawtext = laborLevels['content']
# open or create a temporary file in the current working directory to store the raw as a txt file
f = open(tmpTxtPath, 'w+')
# write the raw to the text file
f.write(str(rawtext))
# close the txt file
f.close
# open the temporary text file and create a new csv file to write the text to 
with open(tmpTxtPath, 'r') as fin, open(tmpCSVPath, 'w+') as fout:
    # loop through each line in the temp text file
    for line in fin:
        # Replace the first space on each line in the text file so it can be saved as a csv file with 2 columns
        fout.write(line.replace(' ', ',', 1))
print('\n I am almost done, please \"bear\" with me... that\'s funny because I\'m using pandas! :)')
# Put the csv contents into a 2 column dataframe; it will encounter errors for lines with additional commas, so just skip those lines
# this is essentially the pdf contents converted to csv format
laborLevels_df = pd.read_csv(tmpCSVPath, error_bad_lines=False, warn_bad_lines=False)
# put the position title to labor type matrix into a pandas dataframe
posMatrix_df = pd.read_excel(supportFilePath, 'PosMatrix')
# name the columns of the content originally from the pdf file
laborLevels_df.columns = ['labor_code', 'description']
# name the columns of the job title matrix file
posMatrix_df.columns = ['description', 'type']
# merge the dataframe from the pdf with the position title matrix on the description so that the labor type is added to the 3rd column
laborLevels_df = reduce(lambda x,y: pd.merge(x,y, on='description', how='inner'), [laborLevels_df, posMatrix_df])
# get rid of any duplicates
laborLevels_df = laborLevels_df.drop_duplicates()
## Write the new labor levels file to excel
#laborLevels_df.to_excel(checkLaborLevelsPath)
# take the raw data output from Kronos into a pandas dataframe
pd_rawData = pd.read_excel(rawDataPath)
# slice the dataframe to only labor account, pay code, and hours columns
pd_rawData = pd_rawData[['Labor Account', 'Pay Code', 'Hours']]
# split the labor account column to get a cost center column from it
pd_rawData['cost_center'] = pd_rawData['Labor Account'].str[12:20]
# split the labor account to get the labor code from it
pd_rawData['labor_code'] = (pd_rawData['Labor Account'].str[15:]).str.split('/').str[1]
# merge the new labor levels dataframe with the modified kronos raw data output so the type of hours are included (ie RN, UAP, etc)
combinedData = reduce(lambda x,y: pd.merge(x,y, on='labor_code', how='left'), [pd_rawData, laborLevels_df])
# Drop the description column; it is no longer needed since the hours type (RN, UAP, etc) is included in the combinedData Dataframe
combinedData = combinedData.drop(columns=['description', 'Labor Account', 'Pay Code'])
# drop the contract hours (labor code 8888) because there is not enough info in this data set to determine the hours type (rn, uap, etc)
combinedData = combinedData[combinedData.labor_code != '8888']
# write file to check that all labor levels have an associated type 
combinedData.to_excel(checkLaborLevelsPath)
# establish the file paths
contractDataPath = os.path.join('K:\\', 'NDNQI', 'SourceData', 'Hours Of Care', yearStr + 'Q' + qtrStr, 'contractRaw ' + yearStr + '-' + monNumStr + '.xls')
contractNonContractCombinedPath =  os.path.join('K:\\', 'NDNQI', 'SourceData', 'Hours Of Care', yearStr + 'Q' + qtrStr, 'ContracNonContractCombined ' + yearStr + '-' + monNumStr + '.xlsx')
# put the contract data into a dataframe
pd_contractData = pd.read_excel(contractDataPath)
# reduce the dataframe to only ne needed columns:  name, account and Hours
pd_contractData = pd_contractData[['Name', 'Account', 'Hours']]
# parse out the cost center from the account column
pd_contractData['cost_center'] = pd_contractData['Account'].str[12:20]
# parse out the labor code from the account
pd_contractData['labor_code'] = (pd_contractData['Account'].str[15:]).str.split('/').str[1]
# reduce the dataframe to only contract hours (identified by labor code 8888)
pd_contractData = pd_contractData.loc[pd_contractData['labor_code'] == '8888'] 
# load the contract matrix; list of contract names and their type (ie. UAP, RN, etc)
pd_contractMatrix = pd.read_excel(supportFilePath, 'ContractNames')
# merge the contract matrix with the contract data on Name so the data contains the work type from the matrix
combinedContractData = reduce(lambda x,y: pd.merge(x,y, on='Name', how='left'), [pd_contractData, pd_contractMatrix])
# create path for file to check that all contract names are associated with a work type in contract matrix
checkContractNamesPath = os.path.join('K:\\', 'NDNQI', 'SourceData', 'Hours Of Care', yearStr + 'Q' + qtrStr, 'checkContractNames ' + yearStr + '-' + monNumStr + '.xlsx')
# write to the file to check that all contractor names are in contract matrix
combinedContractData.to_excel(checkContractNamesPath)
if checkedMatrix == True:
    # drop the account and name columns because it is no longer needed
    combinedContractData = combinedContractData.drop(columns=['Account', 'Name'])
    # append the staff data with the contract data to create the complete hours file
    contractNonContractCombined = combinedData.append(combinedContractData)
    # parse out the UNMH_Cost_Center from the 'cost center'
    contractNonContractCombined['UNMH_Cost_Center'] = contractNonContractCombined['cost_center'].str[-5:]
    # write the contract and non contract complete file to excel
    contractNonContractCombined.to_excel(contractNonContractCombinedPath)
    # establish the path of masterAliasRecord
    masterAliasPath = os.path.join('K:\\', 'NDNQI', 'masterAliasRecord.xlsx')
    # load masterAlias to a dataframe
    masterAlias = pd.read_excel(masterAliasPath, 'MainAlias')
    # reduce masterAlias to only the needed columns
    masterAlias = masterAlias[['UNMH_Cost_Center', 'NDNQIUnitID', 'NDNQI_Reporting_Unit_Name']]
    # convert the UNMH_Cost_Center column in masterAlias to a string (it is int64 in the file)
    masterAlias['UNMH_Cost_Center'] = masterAlias['UNMH_Cost_Center'].astype(str)
    #use only the first five characters in the cost center so any decimals from previous formatting is trunked
    masterAlias['UNMH_Cost_Center'] = masterAlias['UNMH_Cost_Center'].str[:5]
    # Merge the complete data file with masterAlias on the cost center to get the NDNQI unit and the description
    contractNonContractCombined = reduce(lambda x,y: pd.merge(x,y, on='UNMH_Cost_Center', how='left'), [masterAlias, contractNonContractCombined])
    # create a pivot table with columns for each of the work types (ie. RN, UAP, etc)
    pivotData = pd.pivot_table(contractNonContractCombined, values='Hours', index =['UNMH_Cost_Center', 'NDNQIUnitID','NDNQI_Reporting_Unit_Name'], columns=['type'], aggfunc=np.sum, fill_value= float(0))
    # load UAP corrections from the support file into dataframe; this subtracts from UAP hours the expected hours spent as HUC for inpatient units
    corrections = pd.read_excel(supportFilePath, 'Corrections')
    # Reduce completions to only the information that is needed
    corrections = corrections[['UNMH_Cost_Center', 'uapCorrectPerDay', 'supDirectCareIndicator']]
    # convert UNMH_Cost_Center to string so it can be merged with other dataframes
    corrections['UNMH_Cost_Center'] = corrections['UNMH_Cost_Center'].astype(str)
    # merge the pivot table with corrections
    pivotData = reduce(lambda x,y: pd.merge(x,y, on='UNMH_Cost_Center', how='left'), [pivotData, corrections])
    # 'sup direct care indicator' shows which units count their supervisors in Hours of Care; use fillna to fill empty values with 1
    # Units not counting sups have a 0 (zero) indicator; this makes the default that sups will be counted in HOC 
    pivotData['supDirectCareIndicator'] = pivotData['supDirectCareIndicator'].fillna(value=1) 
    # make sup hours = sup hours * 1 if sups are counted, sup hours * 0 if sups not counted
    pivotData['SUP'] = pivotData['SUP'] * pivotData['supDirectCareIndicator']
    # add the corrected sup hours to the RN hours
    pivotData['RN'] = pivotData['RN'] + pivotData['SUP']
    # use fillna to replace any null values in 'uap correct per day' columns; essentially making default = no corrections
    pivotData['uapCorrectPerDay'] = pivotData['uapCorrectPerDay'].fillna(value=0) 
    # subtract the uap correction per day * days in month from uap hours 
    pivotData['UAP'] = pivotData['UAP'] - (pivotData['uapCorrectPerDay'] * daysInMonth)
    # apply a minimum value of 0 to UAP to handle areas who only have HUCs who are coded as UAPs, and may have had call-ins
    # which resulted in a negative number of UAP hours
    pivotData['UAP'] = pivotData['UAP'].apply(lambda x: max(x, 0))
    # establish the path for the final product of the pivot table
    pivotDataPath =  os.path.join('K:\\', 'NDNQI', 'SourceData', 'Hours Of Care', yearStr + 'Q' + qtrStr, 'HOC_Final ' + yearStr + '-' + monNumStr + '.xlsx')
    # merge pivot data with masterAlias to add back the NDNQI Unit ID
    pivotData = reduce(lambda x,y: pd.merge(x,y, on='UNMH_Cost_Center', how='left'), [pivotData, masterAlias])
    # set NDNQI Unit ID to be the index so the dataframe can be combined and summed on that value
    pivotData = pivotData.set_index('NDNQIUnitID')
    # combine and sum the dataframe on NDNQI Unit ID
    pivotData = pivotData.sum(level='NDNQIUnitID')
    # Reduce master alias dataframe to only NDNQI Unit ID and NDNQI_Reporting_Unit_Name since cost center is no longer needed
    masterAlias = masterAlias[['NDNQIUnitID', 'NDNQI_Reporting_Unit_Name']]
    # merge the pivotData and masterAlias on NDNQI Unit ID to add the NDNQI Unit Name
    pivotData = reduce(lambda x,y: pd.merge(x,y, on='NDNQIUnitID', how='left'), [pivotData, masterAlias])
    # re-order the columns to match the upload file
    pivotData = pivotData[['NDNQI_Reporting_Unit_Name', 'NDNQIUnitID', 'RN', 'Contract_RN', 'LPN', 'UAP', 'Contract_UAP', 'MHT',  'PARA', 'EMT', 'LAC']]
    # Drop any duplicate rows
    pivotData = pivotData.drop_duplicates()
    # remove rows that are space holder cost center from master alias (ie. dept nursing excellence, Nursing clinical informatics, etc)
    pivotData = pivotData[pivotData.NDNQIUnitID != 0] 
    # Remove row(s) that is staffing NDNQI unit number
    pivotData = pivotData[pivotData.NDNQIUnitID != 74643]
    #rename the columns to match the NDNQI Standards
    pivotData = pivotData.rename(index=str, columns={'NDNQI_Reporting_Unit_Name': 'Monthly reporting unit', 'NDNQIUnitID': 'UnitID',
                                        'RN': 'RNHospEmplHours', 'Contract_RN': 'RNContractEmplHours', 'LPN': 'LPNHospEmplHours',
                                        'UAP': 'UAPHospEmplHours', 'Contract_UAP': 'UAPContractEmplHours', 'MHT': 'MHTHospEmplHours',
                                        'PARA': 'ParamedicEmplHours', 'EMT': 'EMTEmplHours'})
    # Add missing columns to meet NDNQI upload standards
    pivotData.insert(0, 'Month', monNumStr)
    pivotData.insert(0, 'Quarter', qtrStr)
    pivotData.insert(0, 'Year', yearStr)
    pivotData.insert(8, 'LPNContractEmplHours', 0)
    pivotData.insert(12, 'MHTContractEmplHours', 0)
        # Write the final product file
    pivotData.to_excel(pivotDataPath)
    print('\n \n* The hours of care for ' + yearStr + '-' + monNumStr + ' is now complete.  The final file title is called \'HOC_Final ' + yearStr + '-' + monNumStr + '\'.' )
else:
    print('\n \n* Please review the files checkContractNames and checkLaborLevels files for records missing \n a value in the Type column. Then update the hocSupportFile.  When that is complete please return and run this program again.')

