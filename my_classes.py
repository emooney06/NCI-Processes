###############################################################################################################
# Date:  2019-11-22
# Author:  Ethan Mooney
# Details: This class is intended to create variables needed to access any files on the share drive with
# the standard naming convention "[name] yyyy-mm".  It takes 2 user input values (year and month), then 
# converts the common input values for month (ie. aug, AUG, Aug, August, etc) to a standard numerical value
# (in string datatype).  It also provides the numeric quarter (in a string datatype) which is used to write
# variables for folders using the convention yyyyQ(n)
##############################################################################################################
class FileDateVars:

    def __init__(self, year_str, month_str):
        global numeric_month_str
        self.year_str = year_str
        self.month_str = month_str
        self.numeric_month_str = FileDateVars.getMonthNumStr(self.month_str)
        self.qtr_str = FileDateVars.getQtrStr(self.numeric_month_str)
        
    @classmethod
    def from_inputs(cls):
        return cls(
            input('Enter the year of the data you want: '), 
            input('Enter the month of the data you want (ie. Sep or Sept):  '),
        );

    # this function returns the int representing the month as a string (so it is ready to use in file paths and names
    @staticmethod
    def getMonthNumStr(month_str):
     if (
        month_str == 'Jan'
        or month_str == 'JAN'
        or month_str == 'jan'
        or month_str == 'jAN'
        or month_str == 'January'
        or month_str == 'january'
        or month_str == 'JANUARY'
        or month_str == 'jANUARY'
        or month_str == '01'
        or month_str == '1'
        ): 
        monthNumStr = '01'
     elif (
        month_str == 'Feb'
        or month_str == 'FEB'
        or month_str == 'feb'
        or month_str == 'fEB'
        or month_str == 'February'
        or month_str == 'february'
        or month_str == 'FEBRUARY'
        or month_str == 'fEBRUARY'
        or month_str == '02'
        or month_str == '2'
        ):
        monthNumStr = '02'
     elif (
        month_str == 'Mar'
        or month_str == 'MAR'
        or month_str == 'mar'
        or month_str == 'mAR'
        or month_str == 'March'
        or month_str == 'march'
        or month_str == 'MARCH'
        or month_str == 'mARCH'
        or month_str == '03'
        or month_str == '3'
        ):
        monthNumStr = '03'
     elif (
        month_str == 'Apr'
        or month_str == 'APR'
        or month_str == 'apr'
        or month_str == 'aPR'
        or month_str == 'April'
        or month_str == 'april'
        or month_str == 'APRIL'
        or month_str == 'aPRIL'
        or month_str == '04'
        or month_str == '4'
        ):
        monthNumStr = '04'
     elif (
        month_str == 'May'
        or month_str == 'MAY'
        or month_str == 'may'
        or month_str == 'mAY'
        or month_str == '05'
        or month_str == '5'
        ):
        monthNumStr = '05'
     elif (
        month_str == 'Jun'
        or month_str == 'JUN'
        or month_str == 'jun'
        or month_str == 'jUN'
        or month_str == 'June'
        or month_str == 'june'
        or month_str == 'JUNE'
        or month_str == 'jUNE'
        or month_str == '06'
        or month_str == '6'
        ):
        monthNumStr = '06'
     elif (
        month_str == 'Jul'
        or month_str == 'JUL'
        or month_str == 'jul'
        or month_str == 'jUL'
        or month_str == 'July'
        or month_str == 'july'
        or month_str == 'JULY'
        or month_str == 'jULY'
        or month_str == '07'
        or month_str == '7'
        ):
        monthNumStr = '07'
     elif (
        month_str == 'Aug'
        or month_str == 'AUG'
        or month_str == 'aug'
        or month_str == 'aUG'
        or month_str == 'August'
        or month_str == 'august'
        or month_str == 'AUGUST'
        or month_str == 'aUGUST'
        or month_str == '08'
        or month_str == '8'
        ):
        monthNumStr = '08'
     elif (
        month_str == 'Sep'
        or month_str == 'SEP'
        or month_str == 'sep'
        or month_str == 'sEP'
        or month_str == 'September'
        or month_str == 'september'
        or month_str == 'SEPTEMBER'
        or month_str == 'sEPTEMBER'
        or month_str == '09'
        or month_str == '9'
        ):
        monthNumStr = '09'
     elif (
        month_str == 'Oct'
        or month_str == 'OCT'
        or month_str == 'oct'
        or month_str == 'oCT'
        or month_str == 'October'
        or month_str == 'october'
        or month_str == 'OCTOBER'
        or month_str == 'oCTOBER'
        or month_str == '10'
        ):
        monthNumStr = '10'
     elif (
        month_str == 'Nov'
        or month_str == 'NOV'
        or month_str == 'nov'
        or month_str == 'nOV'
        or month_str == 'November'
        or month_str == 'november'
        or month_str == 'NOVEMBER'
        or month_str == 'nOVEMBER'
        or month_str == '11'
        ):
        monthNumStr = '11'
     elif (
        month_str == 'Dec'
        or month_str == 'DEC'
        or month_str == 'dec'
        or month_str == 'dEC'
        or month_str == 'December'
        or month_str == 'december'
        or month_str == 'DECEMBER'
        or month_str == 'dECEMBER'
        or month_str == '12'
        ):
        monthNumStr = '12'
    
     return monthNumStr;

    # this function returns the int representing the quarter as a string (so it is ready to use in file paths and names
    @staticmethod
    def getQtrStr(month_str):
        if (
            month_str == 'Jan'
            or month_str == 'JAN'
            or month_str == 'jan'
            or month_str == 'jAN'
            or month_str == 'January'
            or month_str == 'january'
            or month_str == 'JANUARY'
            or month_str == 'jANUARY'
            or month_str == '01'
            or month_str == '1'
            ): 
            qtrStr = '1'
        elif (
            month_str == 'Feb'
            or month_str == 'FEB'
            or month_str == 'feb'
            or month_str == 'fEB'
            or month_str == 'February'
            or month_str == 'february'
            or month_str == 'FEBRUARY'
            or month_str == 'fEBRUARY'
            or month_str == '02'
            or month_str == '2'
            ):
            qtrStr = '1'
        elif (
            month_str == 'Mar'
            or month_str == 'MAR'
            or month_str == 'mar'
            or month_str == 'mAR'
            or month_str == 'March'
            or month_str == 'march'
            or month_str == 'MARCH'
            or month_str == 'mARCH'
            or month_str == '03'
            or month_str == '3'
            ):
            qtrStr = '1'
        elif (
            month_str == 'Apr'
            or month_str == 'APR'
            or month_str == 'apr'
            or month_str == 'aPR'
            or month_str == 'April'
            or month_str == 'april'
            or month_str == 'APRIL'
            or month_str == 'aPRIL'
            or month_str == '04'
            or month_str == '4'
            ):
            qtrStr = '2'
        elif (
            month_str == 'May'
            or month_str == 'MAY'
            or month_str == 'may'
            or month_str == 'mAY'
            or month_str == '05'
            or month_str == '5'
            ):
            qtrStr = '2'
        elif (
            month_str == 'Jun'
            or month_str == 'JUN'
            or month_str == 'jun'
            or month_str == 'jUN'
            or month_str == 'June'
            or month_str == 'june'
            or month_str == 'JUNE'
            or month_str == 'jUNE'
            or month_str == '06'
            or month_str == '6'
            ):
            qtrStr = '2'
        elif (
            month_str == 'Jul'
            or month_str == 'JUL'
            or month_str == 'jul'
            or month_str == 'jUL'
            or month_str == 'July'
            or month_str == 'july'
            or month_str == 'JULY'
            or month_str == 'jULY'
            or month_str == '07'
            or month_str == '7'
            ):
            qtrStr = '3'
        elif (
            month_str == 'Aug'
            or month_str == 'AUG'
            or month_str == 'aug'
            or month_str == 'aUG'
            or month_str == 'August'
            or month_str == 'august'
            or month_str == 'AUGUST'
            or month_str == 'aUGUST'
            or month_str == '08'
            or month_str == '8'
            ):
            qtrStr = '3'
        elif (
            month_str == 'Sep'
            or month_str == 'SEP'
            or month_str == 'sep'
            or month_str == 'sEP'
            or month_str == 'September'
            or month_str == 'september'
            or month_str == 'SEPTEMBER'
            or month_str == 'sEPTEMBER'
            or month_str == '09'
            or month_str == '9'
            ):
            qtrStr = '3'
        elif (
            month_str == 'Oct'
            or month_str == 'OCT'
            or month_str == 'oct'
            or month_str == 'oCT'
            or month_str == 'October'
            or month_str == 'october'
            or month_str == 'OCTOBER'
            or month_str == 'oCTOBER'
            or month_str == '10'
            ):
            qtrStr = '4'
        elif (
            month_str == 'Nov'
            or month_str == 'NOV'
            or month_str == 'nov'
            or month_str == 'nOV'
            or month_str == 'November'
            or month_str == 'november'
            or month_str == 'NOVEMBER'
            or month_str == 'nOVEMBER'
            or month_str == '11'
            ):
            qtrStr = '4'
        elif (
            month_str == 'Dec'
            or month_str == 'DEC'
            or month_str == 'dec'
            or month_str == 'dEC'
            or month_str == 'December'
            or month_str == 'december'
            or month_str == 'DECEMBER'
            or month_str == 'dECEMBER'
            or month_str == '12'
            ):
            qtrStr = '4'
    
        return qtrStr;


    mmm_dict = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun', '07': 'Jul', \
        '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'}
