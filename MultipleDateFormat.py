import pandas as pd
import os
import openpyxl
import re
import datetime
from datetime import date

pathSource = input("Enter file with complete path:")

pathOutput = input("Enter output folder path:-")
SheetName = input("Enter sheet Name:-")
ColumnName = input("Enter column name to correct date:-")

# data = pd.read_excel('../Input/SPCS Consolidate Tracker _Nov_2022.xlsx',sheet_name='Nov 01')
data = pd.read_excel(pathSource, sheet_name=SheetName)


def currectDateFormat(availableDate):
    availableDate = str(availableDate)
    mat = re.search("(\d+/\d+/\d+\s\d+:\d+\s\w+)", availableDate)
    mat2 = re.search("(\d+-\d+-\d+\s\d+:.*)", availableDate)
    mat3 = re.search("(\d+/\d+)", availableDate)
    try:
        if mat:
            result = mat.group()
            format = '%m/%d/%Y %H:%M %p'  # The format
            datetime_str = datetime.datetime.strptime(result, format)
            output = datetime.datetime.strftime(datetime_str, "%d/%m/%Y %H:%M:%S %p")
            return output
        if mat2:
            result2 = mat2.group()
            format2 = '%Y-%m-%d %H:%M:%S'
            datetime_str = datetime.datetime.strptime(result2, format2)
            output = datetime.datetime.strftime(datetime_str, "%d/%m/%Y %H:%M:%S %p")
            return output
        if mat3:
            todayDate = date.today()
            currentYear = todayDate.year

            result3 = mat3.group()
            result3 = str(result3)+ "/"+str(currentYear)
            format3 = '%m/%d/%Y'
            datetime_str = datetime.datetime.strptime(result3, format3)
            output = datetime.datetime.strftime(datetime_str, "%d/%m/%Y %H:%M:%S %p")
            return output
    except:
        return availableDate

    # except:
    #     return availableDate

# data["Received Date_created"] = data['Report_Date'].apply(currectDateFormat)
data["Received Date_created"] = data[ColumnName].apply(currectDateFormat)
# data.to_excel("output.xlsx")
data.to_excel(pathOutput + "\output.xlsx",index=False)


