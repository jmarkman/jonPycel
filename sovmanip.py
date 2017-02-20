import os
import sys
import xlwt
import subprocess
from xlrd import open_workbook
from unidecode import unidecode
import sovinput as pycelInput

# Global variables
# Get the relative path of the current user
userhome = os.path.expanduser('~/Desktop/')


def isEmptyRow(row):
    """
    If row is entirely empty

    Parameter: array of values ['one','two','three']
    """
    maxLength = len(row)
    emptyCheck = []
    for value in row:  # for each value in the array "row"
        if value == '':
            emptyCheck.append(value)
    if len(emptyCheck) == maxLength:
        return True


def sliceSubHeaderData(headerRowDict, sheet):
    """Gets relevant data below header row

    Parameter: headerRowDict: dictionary,
    Parameter sheet: a Sheet object

    Returns: data from below the header row
    """
    dataRowStartNumber = headerRowDict.keys()[0] + 1
    allData = pycelInput.loopAllRows(sheet)
    subHeadData = {}
    for key in allData:
        isEmp = False
        if isEmptyRow(allData[key]) == True:
            isEmp = True
        else:
            if key in range(dataRowStartNumber, len(allData)) and isEmp == False:
                subHeadData[key] = allData[key]
    return subHeadData


def combine(headerRowDict, subHeaderData):
    """Combines the header row and subheader row into a singular dictionary

    Parameter: two dictionaries
    Precondition: each dicitonary in {RowNumber: ['Value1','Value2','Value3'.....]}

    Returns: combined dictionary in format like above^
    """
    headSubCombined = headerRowDict.copy()
    headSubCombined.update(subHeaderData)
    return headSubCombined


def findFileName(absPath):
    """Identifies name of SOV given absolute path to file on local machine

    Parameter: Absolute path to file i.e  C://Greg/Desktop/SOV_Name.xls
    Returns: SOV_Name.xls
    """
    slash = absPath.rfind('/')
    dot = absPath.rfind('.')
    name = absPath[slash + 1:dot]
    return name


def getFileExtension(file):
    """Extracts the file extension from the given file. Somewhat necessary(?) for
    .xlsx / .xls inconsistencies.

    Parameter: Absolute path to the file
    Returns: File extension, i.e., .xlsx
    """
    dot = file.rfind('.')
    extension = file[dot:]
    return extension


def comp_converter(comparisonDic):
    """Removes unwanted characters from dictionary, see decompress
    """
    new = {}
    for key in comparisonDic:
        val = comparisonDic[key]
        key = decompress(key)
        new[key] = val
    return new


def decompress(value):
    """'does the reconstruction of each value'

    Greg ended up just setting each valued passed as an argument to
    the function to be cast to a string.

    Returns: param value returned as a string.
    """
    value = str(value)
    return value.lower().strip()


def head_matcher(compressedDict, headerRow, fileName):
    """
    Parameters: compressedDict, headerRow, fileName
        - compressedDict is the comparison dictionary
        - headerRow is the dictionary returned from identifyHeaderRow
        - fileName is the file we're working with, see ask()
    """
    decompressedRow = []
    for header in headerRow.itervalues().next():
        header = decompress(header)
        if header in compressedDict:
            decompressedRow.append(compressedDict[header])
        else:
            decompressedRow.append("XX" + header)
    return headerRow


def adjustments(final):
    """MASTER CALLER FOR ADJUSTMENETS
    Adjusts the data based on the requirements for the workstation.
    """
    locNumFix(final, 'Loc #')

    if final['Physical Building #'] == [] or isColEmpty(final['Physical Building #']) == True:
        physicalBuildingNum(final, 'Physical Building #')
    if final['Single Physical Building #'] == [] or isColEmpty(final['Single Physical Building #']) == True:
        physicalBuildingNum(final, "Single Physical Building #")
    writeRawStreet(final)
    autoLocNum(final)
    autoBldgNum(final)
    stripStreetNum(final, "Street 1")

    if final["Wiring Year"] == [] or isColEmpty(final['Wiring Year']) == True:
        wprhY(final, "Wiring Year")
    if final["Plumbing Year"] == [] or isColEmpty(final['Plumbing Year']) == True:
        wprhY(final, 'Plumbing Year')
    if final["Roofing Year"] == [] or isColEmpty(final['Roofing Year']) == True:
        wprhY(final, 'Roofing Year')
    if final["Heating Year"] == [] or isColEmpty(final['Heating Year']) == True:
        wprhY(final, 'Heating Year')

    if final["Fire Alarm Type"] == [] or isColEmpty(final['Fire Alarm Type']) == True:
        nonePlacer(final, "Fire Alarm Type")
    if final["Burglar Alarm Type"] == [] or isColEmpty(final['Burglar Alarm Type']) == True:
        nonePlacer(final, "Burglar Alarm Type")
    if final["Sprinkler Alarm Type"] == [] or isColEmpty(final['Sprinkler Alarm Type']) == True:
        sprinkAlarmType(final)
    sprinkExtent(final)
    convertBasements(final)
    convertConstructionType(final)
    # stripStreet2(final)
    populationCounter(final, 'State')

    return final


def sprinkExtent(final):
    """Adjusts the Sprinkler Extent to the specified values. AmRisc deals in three percentages:
    0%, 50%, and 100%. Not enough data from other sheet types to decide whether to change the
    conditions to include ranges of numbers instead of hardcoded 0/50/100.
    """
    sprinkExtent = final['Sprinkler Extent'][0]
    locationNum = final['Loc #'][0]
    for itemIndex in range(1, len(sprinkExtent)):
        try:
            if sprinkExtent[itemIndex] == 0.0 or sprinkExtent[itemIndex] == "0%":
                sprinkExtent[itemIndex] = "None"
            elif sprinkExtent[itemIndex] == 0.0 or sprinkExtent == "0" and locationNum[itemIndex] == "":
                sprinkExtent[itemIndex] = ""
            elif sprinkExtent[itemIndex] == 0.5 or sprinkExtent[itemIndex] == "50%":
                sprinkExtent[itemIndex] = "50%"
            elif sprinkExtent[itemIndex] > 0.5 and sprinkExtent[itemIndex] < 1.0 and sprinkExtent[itemIndex] != 0.0 and sprinkExtent[itemIndex] != "":
                sprinkExtent[itemIndex] = "> 50%"
            elif sprinkExtent[itemIndex] == 1.0 or sprinkExtent[itemIndex] == "100%":
                sprinkExtent[itemIndex] = "100%"
            else:
                sprinkExtent[itemIndex] = ""
        except IndexError:
            nonePlacer(final, "Sprinkler Extent")


def sprinkAlarmType(final):
    """Adjusts the Sprinkler Alarm Type column based on whether or not sprinklers are present.
    AmRisc denotes sprinklers being present with a 'Y' or 'N', and other SoVs follow this.
    May have to change this later to work based off of the actual percentage since yes/no
    isn't universal.

    Update 2/13/2017: Some AmRisc sheets don't even have a column "Sprinklered (Y/N)" for
    the program to pull information from. Implemented a bigger try/except workaround that
    defaults to the percentage sprinklered to get some kind of input regardless.
    Might pose a TA issue?
    """
    try:
        sprinkYN = final["Sprinkler Wet/Dry"][0]
        sprnkList = ["Sprinkler Alarm Type"]
        sprnkTypeList = final["Sprinkler Alarm Type"].append(sprnkList)
        sprinkType = final["Sprinkler Alarm Type"][0]
        for cond in range(1, len(sprinkYN)):
            try:
                if sprinkYN[cond].lower() == 'y':
                    sprinkYN[cond] = "Wet"
                    sprinkType.append("Local")
                elif sprinkYN[cond].lower() == 'n':
                    sprinkYN[cond] = "None"
                    sprinkType.append("None")
                else:
                    sprinkYN[cond] = ""
            except IndexError:
                nonePlacer(final, "Sprinkler Alarm Type")
    except IndexError:
        sprinkYN = final["Sprinkler Extent"][0]
        sprnkATList = ["Sprinkler Alarm Type"]
        sprnkWDList = ["Sprinkler Wet/Dry"]
        sprnkTypeList = final["Sprinkler Alarm Type"].append(sprnkATList)
        sprnkExtList = final["Sprinkler Wet/Dry"].append(sprnkWDList)
        sprinkType = final["Sprinkler Alarm Type"][0]
        sprinkWetDry = final["Sprinkler Wet/Dry"][0]
        for cond in range(1, len(sprinkYN)):
            try:
                if sprinkYN[cond] >= 0.5 or sprinkYN[cond] == "50%" or sprinkYN[cond] == "100%":
                    sprinkWetDry.append("Wet")
                    sprinkType.append("Local")
                else:
                    sprinkWetDry.append("None")
                    sprinkType.append("None")
            except IndexError:
                nonePlacer(final, "Sprinkler Alarm Type")


def writeRawStreet(final):
    """Takes the raw street 1 input value and puts it into the delete column and
    attempts to rename the column itself.
    """
    try:
        delCol = final["Delete"]
        fullAddr = ["Full Street Address"]
        delCol.append(fullAddr)
        street1 = final["Street 1"][0]
        fullStreetAddr = final["Delete"][0]
        for street in range(1, len(street1)):
            st = street1[street]
            fullStreetAddr.append(st)
    except IndexError:
        pass


def stripStreetNum(final, street1):
    # strips off the number from the street if it is there

    street1 = final[street1][0]
    for index in range(len(street1)):
        space = street1[index].find(' ')
        posNumber = street1[index][:space]

        if len(posNumber) > 0:
            if posNumber[0].isdigit():
                try:
                    posNumber = posNumber.replace('-', '')
                    street1[index] = street1[index][space:].strip()
                except:
                    pass
            else:
                pass
        street1[0] = "Street 1"


# def stripStreet2(final):
#     st1Arr = final["Street 1"][0]
#     st2Arr = final["Street 1"][0][:]
#     st2Arr.pop(0)
#     tempStreet2 = ["strt2"]
#     variations = ["blg", "suite", "suites", "ste", "bldg", "bld", '#']
#     try:
#         for place, address in enumerate(st2Arr):
#             for var in variations:
#                 if address.lower().find(var) != -1:
#                     idx = address.lower().find(var.lower())
#                     substringStreet2 = address[idx:]
#                     st1Arr[place] = address[:idx]
#                     if substringStreet2 in tempStreet2:
#                         tempStreet2.append("")
#                     else:
#                         tempStreet2.append(substringStreet2)
#                 # else:
#                 #     tempStreet2.append("")
#     except Exception as e:
#         print e
#     final["Street 2"] = [tempStreet2]


def physicalBuildingNum(final, caption):
    """Takes the street numbers from the address contained within (for AmRisc) "*Street Address",
    splits them, and based on the provided caption, places them in a list which is used to
    populate the column that corresponds with the caption.

    Param "final": Dictionary
    Param "caption": String
    """
    try:
        streetArr = final["Street 1"][0][:]
        streetArr.pop(0)
        numTracker = [caption]
        for val in streetArr:
            if len(val) > 0:
                space = val.find(" ")
                dash = val.find("-")
                streetNum = val[:space]
                if dash != -1:
                    num = val[:dash]
                    if not num.isdecimal():
                        num = ""
                elif dash == -1:
                    num = val[:space]
                    if not num.isdecimal():
                        num = ""
                try:
                    if caption == "Single Physical Building #" and dash != -1:
                        numTracker.append(num)
                    elif caption == "Single Physical Building #":
                        numTracker.append(num)
                    elif caption == "Physical Building #":
                        if streetNum[0].isdigit():
                            numTracker.append(streetNum)
                        else:
                            numTracker.append("")
                    else:
                        numTracker.append("")
                except ValueError:
                    pass
            final[caption] = [numTracker]
    except Exception as ex:
        print ex
        pass


def checkIfSame(addrList):
    """Uses the built-in all() function to check if a list contains ONLY a certain string.
    Used to check if there is only one repeat address in the "Full Street Address" column.

    Returns: boolean
    """
    return all(x == addrList[0] for x in addrList)


def autoLocNum(final):
    """Adjusts the "Loc #" column to be "1" if all the addresses in "Full Street Address" are
    the same.
    """
    rawStreets = final["Delete"][0]
    cleanRawStreets = [x for x in rawStreets if x != ""]
    cleanRawStreets.remove("Full Street Address")
    dupesBool = checkIfSame(cleanRawStreets)
    if dupesBool == True:
        bldgNum = final["Loc #"][0]
        for item in range(1, len(cleanRawStreets) + 1):
            bldgNum[item] = 1
    else:
        pass


def autoBldgNum(final):
    """Adjusts the building numbers to count off in ascending order if all the addresses in
    "Full Street Address" are the same.
    """
    rawStreets = final["Delete"][0]
    cleanRawStreets = [x for x in rawStreets if x != ""]
    cleanRawStreets.remove("Full Street Address")
    dupesBool = checkIfSame(cleanRawStreets)
    if dupesBool == True:
        bldgNum = final["Bldg #"]
        number = ["Bldg"]
        bldgNum.append(number)
        wsBldgNum = final["Bldg #"][0]
        i = 1
        for item in cleanRawStreets:
            wsBldgNum.append(i)
            i += 1
    else:
        pass


def convertBasements(final):
    """Takes the basement values provided on the SoV and converts it to either 0 for no basement
    or 1 for a present basement.
    """
    bsmtArray = final['# Basements'][0]
    for item in range(1, len(bsmtArray)):
        try:
            if bsmtArray[item][:1] == "2":
                bsmtArray[item] = 0
            elif bsmtArray[item][:1] == "1":
                bsmtArray[item] = 1
            elif bsmtArray[item][:1] == "3":
                bsmtArray[item] = 1
            else:
                bsmtArray[item] = ""
        except ValueError:
            nonePlacer(final, "# Basements")


def convertConstructionType(final):
    """Takes the values within the Construction Type column and converts them to
    the workstation's standard format for construction types.
    """
    sovConstTypes = final["Construction Type"][0]
    constTypeDictionary = {
        "brick frame": "1. Frame",
        "frame": "1. Frame",
        "frame ": "1. Frame",
        "brick veneer": "1. Frame",
        "frame block": "1. Frame",
        "heavy timber": "1. Frame",
        "masonry frame": "1. Frame",
        "masonry wood": "1. Frame",
        "metal building": "1. Frame",
        "sheet metal": "1. Frame",
        "wood": "1. Frame",
        "metal/aluminum": "1. Frame",
        "brick": "2. Joisted Masonry",
        "brick steel": "2. Joisted Masonry",
        "cd": "2. Joisted Masonry",
        "cement": "2. Joisted Masonry",
        "concrete": "2. Joisted Masonry",
        "masonry": "2. Joisted Masonry",
        "masonry timbre": "2. Joisted Masonry",
        "stone": "2. Joisted Masonry",
        "stucco": "2. Joisted Masonry",
        "joist masonry": "2. Joisted Masonry",
        "tilt-up": "2. Joisted Masonry",
        "jm": "2. Joisted Masonry",
        "joisted masonry": "2. Joisted Masonry",
        "joisted mason": "2. Joisted Masonry",
        "j/masonry": "2. Joisted Masonry",
        "jm / cbs": "2. Joisted Masonry",
        "cb": "3. Non-Combustible",
        "concrete block": "3. Non-Combustible",
        "icm": "3. Non-Combustible",
        "iron clad metal": "3. Non-Combustible",
        "steel concrete": "3. Non-Combustible",
        "steel cmu": "3. Non-Combustible",
        "non-comb.": "3. Non-Combustible",
        "non-comb": "3. Non-Combustible",
        "pole": "3. Non-Combustible",
        "superior nc": "3. Non-Combustible",
        "non-combustible": "3. Non-Combustible",
        "non-combustib": "3. Non-Combustible",
        "cement block": "4. Masonry, Non-Combustible",
        "cbs": "4. Masonry, Non-Combustible",
        "mnc": "4. Masonry, Non-Combustible",
        "ctu": "4. Masonry, Non-Combustible",
        "concrete tilt-up": "4. Masonry, Non-Combustible",
        "pre-cast com": "4. Masonry, Non-Combustible",
        "reinforced concrete": "4. Masonry, Non-Combustible",
        "masonry nc": "4. Masonry, Non-Combustible",
        "masonry non-c": "4. Masonry, Non-Combustible",
        "masonry non-combustible": "4. Masonry, Non-Combustible",
        "mfr": "5. Modified Fire Resistive",
        "modified fire resistive": "5. Modified Fire Resistive",
        "aaa": "6. Fire Resistive",
        "fire res": "6. Fire Resistive",
        "fire resistive": "6. Fire Resistive",
        "cinder block": "6. Fire Resistive",
        "steel": "6. Fire Resistive",
        "steel frame": "6. Fire Resistive",
        "superior": "6. Fire Resistive",
        "w/r": "6. Fire Resistive",
        "fire resist": "6. Fire Resistive",
        "wind resistive": "6. Fire Resistive",
        "fire resistiv": "6. Fire Resistive",
        "fr": "6. Fire Resistive",
        "f.r.": "6. Fire Resistive",
        "fr/wr": "6. Fire Resistive",
    }

    for index, item in enumerate(sovConstTypes):
        for key in constTypeDictionary:
            if sovConstTypes[index].lower().strip() in constTypeDictionary.iterkeys():
                sovConstTypes[index] = constTypeDictionary[
                    item.lower().strip()]


def populationCounter(final, caption):
    """Returns the number of non empty rows in the column specified by caption

    Param "final": Dictionary
    Param "caption": String
    """
    print "POPULATION COUNTER RUNNING"
    column = final[caption][0]
    nonEmptyCount = 0
    for item in column:
        if item != '':
            nonEmptyCount += 1
    return nonEmptyCount


def locNumFix(final, headerName):
    """Splices the location numbers to the appropriate length based on state
    """
    toWriteCount = populationCounter(final, "State")
    final[headerName][0] = final[headerName][0][:toWriteCount]


def nonePlacer(final, headerName):
    """Places None in the entire column
    NOTE: to retrieve appropriate length it relies on the assumption that the STATE column is full

    Parameter headerName: Name of the workstation header to add the Nones to. i.e 'Fire Alarm Type'
    """
    toWriteCount = populationCounter(final, "State")
    toWriteRow = []
    for item in range(toWriteCount):
        toWriteRow.append('None')
    toWriteRow[0] = headerName
    final[headerName].append(toWriteRow)


def wprhY(final, headerName):
    """For wiring, plumbing, roofing, heating year. If emtpy fill their contents with the
    data from Year Built, else does nothing
    """
    if final['Year Built'] != []:
        arr = final['Year Built'][:]
        final[headerName] = arr

def isColEmpty(columnVals):
    """If there's no data in the column return True, if there is data return False

    Parameter: Takes an array of values
    Precondition:inputs can be a double array[['AddressNum', '', '', '']] or a single empty array []
    """
    print "\nExecuting isColEmpty"
    if columnVals == []:
        return True
    notEmptyCount = 0
    for ArrayOrVal in columnVals:
        if type(ArrayOrVal) == list:
            for value in ArrayOrVal:
                value = str(value)
                if value != "":
                    notEmptyCount += 1
    if notEmptyCount <= 1:
        return True
    else:
        print "Column has more than 1 value in it, not empty"
        return False


def setnwrite(headSubCombined, fileName):
    """Formats the data from the parameter headSubCombined and calls writer()
    to actually wrtie the data to the Excel sheet. This is where the templates
    for the SoVs come into play.

    TODO: Better method for switching between the templates. Take the formatting
    code and shove it in another function, then call it based on what SoV is being processed?
    """

    # Get instance of workbook and sheet
    workbook = open_workbook(fileName[0])
    sheet = workbook.sheet_by_index(0)
    # Declare hardcoded locations on the sheet that signify what kind of sheet
    # it should be
    # "AmRisc Application / Schedule of Values"
    amriscCell1 = sheet.cell_value(0, 0)
    # Same as above but for when brokers "de-format" the spreadsheet somehow
    amriscCell2 = sheet.cell_value(0, 1)
    amriscCell3 = sheet.cell_value(7, 0)  # Who makes these sheets?
    crcSwettCell = sheet.cell_value(10, 0)  # "LOCATION INFORMATION"
    amriscID1 = "AmRisc Application / Schedule of Values"
    amriscID2 = "Starred * information is needed to process the account."
    crcSwettID = "LOCATION INFORMATION"

    workHeaderRow = [
        'Loc #',
        'Bldg #',
        'Delete',
        'Physical Building #',
        'Single Physical Building #',
        'Street 1',
        'Street 2',
        'City',
        'State',
        'Zip',
        'County',
        'Validated Zip',
        'Building Value',
        'Business Personal Property',
        'Business Income',
        'Misc Real Property',
        'TIV',
        '# Units',
        'Building Description',
        'ClassCodeDesc',
        'Construction Type',
        'Dist. To Fire Hydrant (Feet)',
        'Dist. To Fire Station (Miles)',
        'Prot Class',
        '# Stories',
        '# Basements',
        'Year Built',
        'Sq Ftg',
        'Wiring Year',
        'Plumbing Year',
        'Roofing Year',
        'Heating Year',
        'Fire Alarm Type',
        'Burglar Alarm Type',
        'Sprinkler Alarm Type',
        'Sprinkler Wet/Dry',
        'Sprinkler Extent',
        'Roof Covering',
        'Roof Geometry',
        'Roof Anchor',
        'Cladding Type',
        'Roof Sheathing Attachment',
        'Frame-Foundation Connection',
        'Residential Appurtenant Structures'
        ]

    final = {key: [] for key in workHeaderRow}

    work = {
        'Loc #': 0,
        'Bldg #': 1,
        'Delete': 2,
        'Physical Building #': 3,
        'Single Physical Building #': 4,
        'Street 1': 5,
        'Street 2': 6,
        'City': 7,
        'State': 8,
        'Zip': 9,
        'County': 10,
        'Validated Zip': 11,
        'Building Value': 12,
        'Business Personal Property': 13,
        'Business Income': 14,
        'Misc Real Property': 15,
        'TIV': 16,
        '# Units': 17,
        'Building Description': 18,
        'ClassCodeDesc': 19,
        'Construction Type': 20,
        'Dist. To Fire Hydrant (Feet)': 21,
        'Dist. To Fire Station (Miles)': 22,
        'Prot Class': 23,
        '# Stories': 24,
        '# Basements': 25,
        'Year Built': 26,
        'Sq Ftg': 27,
        'Wiring Year': 28,
        'Plumbing Year': 29,
        'Roofing Year': 30,
        'Heating Year': 31,
        'Fire Alarm Type': 32,
        'Burglar Alarm Type': 33,
        'Sprinkler Alarm Type': 34,
        'Sprinkler Wet/Dry': 35,
        'Sprinkler Extent': 36,
        'Roof Covering': 37,
        'Roof Geometry': 38,
        'Roof Anchor': 39,
        'Cladding Type': 40,
        'Roof Sheathing Attachment': 41,
        'Frame-Foundation Connection': 42,
        'Residential Appurtenant Structures': 43
    }

    amrisc = {
        "Percent Sprinklered": "Sprinkler Extent",
        "Sprinklered (Y/N)": "Sprinkler Wet/Dry",
        "Sprinkler Alarm Type": "Sprinkler Alarm Type",
        "Sprinkler Wet/Dry": "Sprinkler Wet/Dry",
        "Sprinkler Extent": "Sprinkler Extent",
        "Physical Building #": "Physical Building #",
        "Single Physical Building #": "Single Physical Building #",
        "Street 1": "Street 1",
        "Fire Alarm Type": "Fire Alarm Type",
        "Burglar Alarm Type": "Burglar Alarm Type",
        "Full Street Address": "Full Street Address",
        "Delete": "Full Street Address",
        "*Year Roof covering last fully replaced": "Roofing Year",
        "* Bldg No.": "Loc #",
        "Bldg": "Bldg #",
        "Loc": "Loc #",
        "*Orig Year Built": "Year Built",
        "*Square Footage": "Sq Ftg",
        "*# of Stories": "# Stories",
        "AddressNum": "Physical Building #",
        "*Street Address": "Street 1",
        "*City": "City",
        "*State Code": "State",
        "*Zip": "Zip",
        "County": "County",
        "*Real Property Value ($)": "Building Value",
        "Personal Property Value ($)": "Business Personal Property",
        "personal property value  ($)": "Business Personal Property",
        "Personal Property Value ($) ": "Business Personal Property",
        "Other Value $ (outdoor prop & Eqpt must be sch'd)": "Misc Real Property",
        "BI/Rental Income ($)": "Business Income",
        "*Total TIV": "TIV",
        "*Occupancy": "Building Description",
        "Construction Description ": "Construction Type",
        "Construction Description (provide further details on construction features)": "Construction Type",
        "ISO Prot Class": "Prot Class",
        "strt2":"Street 2",
        "*# of Units": "# Units",
        "Fire Alarm Type": "Fire Alarm Type",
        "Burglar Alarm Type": "Burglar Alarm Type",
        "*Basement": "# Basements",
        "Basement": "# Basements"
    }

    crcSwett = {
        "Loc  #": "Loc #",
        "Location Street Address:": "Street 1",
        "City": "City",
        "State": "State",
        "Zip Code": "Zip",
        "Building Value": "Building Value",
        "Content": "Business Personal Property",
        "BI w/ EE": "Business Income",
        "Total TIV": "TIV",
        "# Apt  Units": "# Units",
        "Building Occupancy": "Building Description",
        "Construction": "Construction Type",
        "# of Stories": "# Stories",
        "Yr Built Gut/Reh": "Year Built",
        "Total Building  Area": "Sq Ftg",
        "Plumbing": "Plumbing Year",
        "Heating": "Heating Year",
        "strt2":"Street 2",
        "Electrical": "Wiring Year",
        "Roof": "Roofing Year",
        "Sprinkler %": "Sprinkler Extent",
        "Sprinkler Alarm Type": "Sprinkler Alarm Type",
        "Sprinkler Wet/Dry": "Sprinkler Wet/Dry",
        "Sprinkler Extent": "Sprinkler Extent",
        "Physical Building #": "Physical Building #",
        "Fire Alarm Type": "Fire Alarm Type",
        "Burglar Alarm Type": "Burglar Alarm Type",
        "Single Physical Building #": "Single Physical Building #",
        "Full Street Address": "Full Street Address",
        "Delete": "Full Street Address",
        "Street 1": "Street 1"
    }

    rps = {
        "Loc. #": "Loc #",
        "Bldg.": "Bldg #",
        "Address": "Street 1",
        "City": "City",
        "County": "County",
        "ST": "State",
        "Zip Code": "Zip",
        "Building": "Building Value",
        "Contents ": "Business Personal Property",
        "Business Income (Incl EE)": "Business Income",
        "Other": "Misc Real Property",
        "Total": "TIV",
        "# of Units": "# Units",
        "Occupancy": "Building Description",
        "ISO Construction Code": "Construction Type",
        "Protection Class Code": "Prot Class",
        "No. Stories": "# Stories",
        "% Sprinkler": "Sprinkler Extent",
        "Year Wiring Updated": "Wiring Year",
        "Year Plumbing Updated": "Plumbing Year",
        "Year Roof Replaced": "Roofing Year",
        "Year Heating updated": "Heating Year",
        "Basement": "# Basements",
        "Sprinkler Alarm Type": "Sprinkler Alarm Type",
        "Sprinkler Wet/Dry": "Sprinkler Wet/Dry",
        "Sprinkler Extent": "Sprinkler Extent",
        "Physical Building #": "Physical Building #",
        "Single Physical Building #": "Single Physical Building #",
        "Full Street Address": "Full Street Address",
        "Delete": "Full Street Address",
        "strt2":"Street 2",
        "Street 1": "Street 1"
    }
    if amriscCell1.find(amriscID1) != -1 or amriscCell2.find(amriscID1) != -1 or amriscCell3.find(amriscID2) != -1:
        minimum = min(headSubCombined, key=headSubCombined.get) - 1
        headerRow = headSubCombined[minimum]
        sov_index = {}
        for itemIndex in range(len(headerRow)):
            if headerRow[itemIndex] in amrisc:
                sov_index[itemIndex] = True

        columnDict = {}
        for index in sov_index:
            column = []
            for key, item in headSubCombined.iteritems():
                column.append(item[index])
            columnDict[column[0]] = column

        for key in columnDict:
            if amrisc[key] in work:
                final[amrisc[key]].append(columnDict[key])
        final = adjustments(final)

        writer(final, work, workHeaderRow, amrisc, fileName)
    elif crcSwettCell.find(crcSwettID) != -1:
        minimum = min(headSubCombined, key=headSubCombined.get) - 1
        headerRow = headSubCombined[minimum]
        sov_index = {}
        for itemIndex in range(len(headerRow)):
            if headerRow[itemIndex] in crcSwett:
                sov_index[itemIndex] = True

        columnDict = {}
        for index in sov_index:
            column = []
            for key, item in headSubCombined.iteritems():
                column.append(item[index])
            columnDict[column[0]] = column

        for key in columnDict:
            if crcSwett[key] in work:
                final[crcSwett[key]].append(columnDict[key])
        final = adjustments(final)
        writer(final, work, workHeaderRow, crcSwett, fileName)


def writer(final, workDict, workHeaderRow, template, sovFileName):
    """Actually writes our data to the spreadsheet and produces our finished product
    """
    workbook = xlwt.Workbook()
    colWidth = 365 * 18
    rowHeight = 256
    wordWrapHeader = xlwt.easyxf('align: horiz center; font: bold on')
    wordWrap = xlwt.easyxf('align: horiz center')

    # HOW TO DO A CELL OVERWRITE AS A LAST RESORT IF NEEDED - greg
    # sheet = workbook.add_sheet("WKFC_Sheet1", cell_overwrite_ok=True)
    sheet = workbook.add_sheet("WKFC_Sheet1")

    for key, values in final.iteritems():
        colIndex = workDict[key]

        if values == []:
            sheet.write(0, colIndex, key, wordWrapHeader)
            sheet.col(colIndex).width = 0
            sheet.row(rowIndex).height = rowHeight
        else:
            valueArr = values[0]
            for rowIndex in range(len(valueArr)):
                if valueArr[rowIndex] in template:
                    sheet.write(rowIndex, colIndex, key, wordWrapHeader)
                    sheet.col(colIndex).width = colWidth
                    sheet.row(rowIndex).height = rowHeight
                else:
                    sheet.write(rowIndex, colIndex, valueArr[rowIndex], wordWrap)
                    sheet.col(colIndex).width = colWidth
                    sheet.row(rowIndex).height = rowHeight

    # TODO: Try/catch for IOError if SoV is re-run but results are still open
    sovCheck = os.path.isfile(sovFileName[0])
    if sovCheck == False:
        workbook.save(sovFileName[0])
        print "FILE WRITTEN"
        os.startfile(userhome + sovFileName[0])
        print "FILE NOW OPEN"
    else:
        sovName = findFileName(sovFileName[0])
        precedingName = '[Pycel_Extracted]_'
        fileType = getFileExtension(sovFileName[0])
        newFileName = precedingName + sovName + fileType
        workbook.save(userhome + newFileName)
        print "FILE WRITTEN"
        os.startfile(userhome + newFileName)
        print "FILE NOW OPEN"
