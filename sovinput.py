import Tkinter
import os
import sys
from xlrd import open_workbook
from unidecode import unidecode
from tkFileDialog import askopenfilename

# Global variables
userhome = os.path.expanduser('~/Desktop/')
# Get the relative path of the current user, i.e. C:/Users/jmarkman/Desktop


def ask():
    """opens the file explorer allowing user to choose SOV to parse

    Precondition: File is an SOV with a .xlsx or .xls file format
    Returns: a singular item array(for possibility to later loop through multiple files, TBD)
    with the aboslute path of the file

    This function uses Tkinter to present the user with askopenfilename() to get the SoV.
    The function returns said file to the rest of the program.
    """
    root = Tkinter.Tk()
    root.withdraw()
    sov = askopenfilename(initialdir=userhome)
    sovFile = [sov]
    return sovFile


def findSheetName(files):
    """find the sheet with the SOV through hierarchy guessing

    Parameter file: array of absolute path
    Returns: Sheet Object where the SOV should be on

    This function takes the returned file from ask() and finds the sheet where all the
    info we need will be. It starts specifically and ends with simply accessing the first
    sheet in the workbook. It returns the sheet that meets one of these conditions.
    """
    for item in files:
        try:
            wb = open_workbook(item)
            try:
                sheet1 = wb.sheet_by_name("SOV")
            except:
                try:
                    sheet1 = wb.sheet_by_name("SOV-APP")
                except:
                    try:
                        sheet1 = wb.sheet_by_name("AmRisc SOV")
                    except:
                        try:
                            sheet1 = wb.sheet_by_name("BREAKDOWN")
                        except:
                            try:
                                sheet1 = wb.sheet_by_name("Property Schedule")
                            except:
                                try:
                                    sheet1 = wb.sheet_by_name("2015 Schedule")
                                except:
                                    try:
                                        sheet1 = wb.sheet_by_name("Sheet1")
                                    except:
                                        sheet1 = wb.sheet_by_index(0)
        except IOError:
            # print "Closing!"
            sys.exit(0)
        return sheet1


def loopAllRows(sheet):
    """Loops through all rows in the current sheet, formats the data

    Parameter: A sheet object
    Returns: Pure value data from the current sheet in
    {rowNumber: ['value1','value2','value3']} format
    """
    ROWCAP = sheet.nrows

    totalPureData = {}
    i = 0
    while i < ROWCAP:
        if i < ROWCAP:
            row = sheet.row(i)
            rowPureData = []
            for item in row:
                rowPureData.append(item.value)
            totalPureData[i] = rowPureData
        i += 1
    return totalPureData


def identifyHeaderRow(numValsDict, comparisonDic):
    """Identifies the header row

    Parameter: Pure sheet data in a dictionary where key=rowNumber value=value array
    Precondition: Header captions are in a singular row, and within the first 30 rows of the file
    """
    maxCap = 30
    rowNumber = 1
    headerRowNum = 0
    while (rowNumber < maxCap) and (rowNumber < len(numValsDict)):
        row = numValsDict[rowNumber]
        matchCounter = 0
        for value in row:
            if isinstance(value, unicode) == unicode:
                value = unidecode(value)
            value = str(value).lower()

            if value in comparisonDic:
                matchCounter += 1
        if matchCounter > headerRowNum:
            headerRowNum = rowNumber

        rowNumber += 1

    headerNum_Vals = {}
    headerNum_Vals[headerRowNum] = numValsDict[headerRowNum]
    return headerNum_Vals
