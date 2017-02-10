import os, sys, xlwt, logging, subprocess
from xlrd import open_workbook
from unidecode import unidecode
import sovinput as input

"""
This program will take a an SOV of a specific template (right now just amrisc)
and convert its data into a workstation file

This file will be written to the desktop as testDoc.xls (right now)

To export this data to a database, run export.py and select this file  

"""
# Global variables
userhome = os.path.expanduser('~/Desktop/') # Get the relative path of the current user, i.e. C:/Users/jmarkman/Desktop
"""
The "desktop" variable is hackneyed because os.path.expanduser can take a longer filepath argument, i.e., a Windows-based system can send '~/Documents/Folder1/Folder2/Folder3' and the value stored in the variable will be a relative filepath into Folder 3.

http://stackoverflow.com/questions/2953828/accessing-relative-path-in-python
"""

def isEmptyRow(row):
	"""If row is entirely empty 

	Parameter: array of values ['one','two','three']
	"""
	
	maxLength = len(row)
	emptyCheck = []
	for value in row: # for each value in the array "row"
		if value == '':
			emptyCheck.append(value)
	if len(emptyCheck) == maxLength:
		return True

def sliceSubHeaderData(headerRowDict, sheet):
	"""
	Gets relevant data below header row

	Parameter: headerRowDict: dictionary,
	Parameter sheet: a Sheet object

	Returns: data from below the header row
	"""

	dataRowStartNumber=headerRowDict.keys()[0]+1
	allData = input.loopAllRows(sheet)
	subHeadData={}
	for key in allData:
		isEmp=False
		if isEmptyRow(allData[key])==True:
			isEmp=True
		else:
			if key in range(dataRowStartNumber, len(allData)) and isEmp==False:
				subHeadData[key]=allData[key]
	return subHeadData

def combine(headerRowDict,subHeaderData):
	"""
	Combines the header row and subheader row into a singular dictionary

	Parameter: two dictionaries
	Precondition: each dicitonary in {RowNumber: ['Value1','Value2','Value3'.....]}

	Returns: combined dictionary in format like above^
	"""

	headSubCombined = headerRowDict.copy()
	headSubCombined.update(subHeaderData)
	return headSubCombined


def findFileName(absPath):
	"""
	Identifies name of SOV given absolute path to file on local machine

	Parameter: Absolute path to file i.e  C://Greg/Desktop/SOV_Name.xls
	Returns: SOV_Name.xls
	"""

	slash = absPath.rfind('/')
	dot = absPath.rfind('.')
	name = absPath[slash+1:dot]
	return name

def getFileExtension(file):
	"""
	Extracts the file extension from the given file. Somewhat necessary(?) for .xlsx / .xls inconsistencies.

	Parameter: Absolute path to the file
	Returns: File extension, i.e., .xlsx
	"""
	dot = file.rfind('.')
	extension = file[dot:]
	return extension

def comp_converter(comparisonDic):
	"""Removes unwanted characters from dictionary
	--see decompress
	
	"""
	new={}
	for key in comparisonDic:
		val = comparisonDic[key]
		key = decompress(key)
		new[key] = val
		# with open('E:\Work\Pycel\jonPycel\output\comp_converter().txt', 'w') as cc:
		# 	cc.write("The contents of the dictionary new\n\n")
		# 	for key, value in new.iteritems():
		# 		cc.write('{0}: {1}\n'.format(key, value))
	return new

def decompress(value):
	"""does the reconstruction of each value
	
	Greg ended up just setting each valued passed as an argument to the function to be cast to a string.

	Returns: param value returned as a string.
	"""
	value = str(value)
	return value.lower().strip()


def head_matcher(compressedDict, headerRow, fileName):
	'''
	Parameters: compressedDict, headerRow, fileName
		- compressedDict is the comparison dictionary
		- headerRow is the dictionary returned from identifyHeaderRow
		- fileName is the file we're working with, see ask()
	'''
	decompressedRow=[]
	# file=open('unmatches.txt','a+')
	for header in headerRow.itervalues().next():
		header=decompress(header)
		if header in compressedDict:
			decompressedRow.append(compressedDict[header])
		else:
			decompressedRow.append("XX"+header)
	return headerRow


def adjustments(final):
	"""
	MASTER CALLER FOR ADJUSTMENETS
	Adjusts the data
	"""
	locNumFix(final,'Loc #')

	if final['Physical Building #']==[] or isColEmpty(final['Physical Building #'])==True: physicalBuildingNum(final, 'Physical Building #')
	if final['Single Physical Building #']==[] or isColEmpty(final['Single Physical Building #'])==True: physicalBuildingNum(final, "Single Physical Building #")
	
	street1Fix(final, "Street 1")

	if final["State"]==[] or isColEmpty(final['State'])==True: statesConverter(final, "State")

	if final["Wiring Year"]==[] or isColEmpty(final['Wiring Year'])==True:     wprhY(final, "Wiring Year")
	if final["Plumbing Year"]==[]or isColEmpty(final['Plumbing Year'])==True:  wprhY(final, 'Plumbing Year')
	if final["Roofing Year"]==[]or isColEmpty(final['Roofing Year'])==True:    wprhY(final, 'Roofing Year')
	if final["Heating Year"]==[]or isColEmpty(final['Heating Year'])==True:   wprhY(final, 'Heating Year')

	if final["Fire Alarm Type"]==[] or isColEmpty(final['Fire Alarm Type'])==True: nonePlacer(final, "Fire Alarm Type")
	if final["Burglar Alarm Type"]==[] or isColEmpty(final['Burglar Alarm Type'])==True: nonePlacer(final, "Burglar Alarm Type")
	if final["Sprinkler Alarm Type"]==[] or isColEmpty(final['Sprinkler Alarm Type'])==True: sprinkAlarmType(final)
	# sprinkWetDry(final)
	sprinkExtent(final)
	populationCounter(final, 'State')

	return final

# sprinkWetDry() is most likely a deprecated function, do not uncomment until further notice
'''
def sprinkWetDry(final):
	"""In amrisc, Column sprinkered Wet/Dry only accepts a Dry, Wet, None... this correct it to ____ and None
	TODO: is it safe to just mark it as none if empty?
	TODO: Amrisc has some limitations here, what should be the switch?"""

	# for now only adjust the N setting in amrisc

	sprinkWetDry=final['Sprinkler Wet/Dry'][0]
	for itemIndex in range(len(sprinkWetDry)):
		if sprinkWetDry[itemIndex] =="N":
			sprinkWetDry[itemIndex]="None"
'''

def sprinkExtent(final):
	"""
	Adjusts the sprinkExtent to NONE FOR RIGHT NOW, NEED TO UNDERSTAND BUSINESS RULES FOR THIS
	In amrisc, Percent Sprinklered takes anything but dan wants it in  100%, >50%, <=50%, how to manipulate thresholds for non amrisc?
	"""

	sprinkExtent = final['Sprinkler Extent'][0]
	locationNum = final['Loc #'][0]
	for itemIndex in range(1, len(sprinkExtent)):
		try:
			if sprinkExtent[itemIndex] == 0.0:
				sprinkExtent[itemIndex] = "None"
			elif sprinkExtent[itemIndex] == 0.0 and locationNum[itemIndex] == "":
				sprinkExtent[itemIndex] = ""
			elif sprinkExtent[itemIndex] == 0.5:
				sprinkExtent[itemIndex] = "50%"
			elif sprinkExtent[itemIndex] == 1.0:
				sprinkExtent[itemIndex] = "100%"
			else:
				sprinkExtent[itemIndex] = ""
		except IndexError:
			nonePlacer(final, "Sprinkler Extent")


def sprinkAlarmType(final):
	sprinkYN = final['Sprinkler Wet/Dry'][0]
	sprnkList = ['Sprinkler Alarm Type']
	sprnkTypeList = final['Sprinkler Alarm Type'].append(sprnkList)
	sprinkType = final['Sprinkler Alarm Type'][0]
	for cond in range(1, len(sprinkYN)):
		try:
			if sprinkYN[cond] == 'Y' or sprinkYN[cond] == 'y':
				sprinkYN[cond] = "Wet"
				sprinkType.append("Local")
			elif sprinkYN[cond] == 'N' or sprinkYN[cond] == 'n':
				sprinkYN[cond] = "None"
				sprinkType.append("None")		
			else:
				sprinkYN[cond] = ""
		except IndexError:
			nonePlacer(final, "Sprinkler Alarm Type")

def street1Fix(final, street1):
	#strips off the number from the street if it is there

	street1 = final[street1][0]
	for index in range(len(street1)):
		space = street1[index].find(' ')
		posNumber = street1[index][:space]
		try:
			posNumber=posNumber.replace('-',' ')
			street1[index]=street1[index][space:].strip()
		except:
			pass
	street1[0] = "Street 1"


def physicalBuildingNum(final, caption):
	"""
	Takes the street numbers from the address contained within (for AmRisc) "*Street Address", 
	splits them, and based on the provided caption, places them in a list which is used to
	populate the column that corresponds with the caption.
	
	Param "final": Dictionary
	Param "caption": String
	"""
	# print 'PHYSICAL BULDING NUMBER IS RUNNING '
	# print caption
	try:
		streetArr = final["Street 1"][0][:]
		streetArr.pop(0)
		numTracker = [caption]
		for val in streetArr:
			if len(val)>0:
				space = val.find(" ")
				dash = val.find("-")
				if dash > 0:
					holder = val[:space]
					num1 = holder[:dash]
					num2 = holder[dash+1:]
					#print "the numval %s" %(num1)
					#print "the numval %s" %(num2)
				else:
					num = val[:space]
					#print "the numval %s" %(num)
				try:
					if caption == 'Single Physical Building #' and dash != -1: 
						int(num2)
						numTracker.append(str(num2))
					elif caption == 'Single Physical Building #':
						int(num)
						numTracker.append(str(num))
					else:
						int(num1)
						numTracker.append(str(num1))
				except ValueError:
					pass
			final[caption] = [numTracker]
	except:
		pass

# def populateBasements(final):
"""
Add entry to dictionary? "Basement":"# Basements"
"""


def populationCounter(final, caption):
	"""
	Returns the number of non empty rows in the column specified by caption
	
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
	"""splices the location numbers to the appropriate length based on state"""
	toWriteCount=populationCounter(final,"State")
	final[headerName][0]=final[headerName][0][:toWriteCount]
	
def nonePlacer(final, headerName):
	"""Places None in the entire column 
	NOTE: to retrieve appropriate length it relies on the assumption that the STATE column is full

	Parameter headerName: Name of the workstation header to add the Nones to. i.e 'Fire Alarm Type' """
	toWriteCount=populationCounter(final, "State")
	toWriteRow=[]
	for item in range(toWriteCount):
		toWriteRow.append('None')
	toWriteRow[0]=headerName
	final[headerName].append(toWriteRow)

def wprhY(final, headerName):
	"""For wiring, plumbing, roofing, heating year. If emtpy fill their contents with the 
	data from Year Built, else does nothing """

	if final['Year Built']!=[]:
		arr=final['Year Built'][:]
		final[headerName]=arr

def statesConverter(final, headerName):
	"""Converts state abbreviation into full name """
     # TODO test this
	 # Amrisc SOV itself only accepts abbreviations, test on other templates

	states={'mississippi': 'MS', 'oklahoma': 'OK', 'wyoming': 'WY', 'minnesota': 'MN', 'alaska': 'AK', 'arkansas': 'AR', 'new mexico': 'NM', 'indiana': 'IN', 'maryland': 'MD', 'louisiana': 'LA', 'texas': 'TX', 'tennessee': 'TN', 'iowa': 'IA', 'wisconsin': 'WI', 'arizona':'AZ', 'michigan': 'MI', 'kansas': 'KS', 'utah': 'UT', 'virginia': 'VA', 'oregon': 'OR', 'connecticut': 'CT', 'district of columbia': 'DC', 'new hampshire': 'NH', 'idaho': 'ID', 'west virginia': 'WV', 'south carolina': 'SC', 'california': 'CA', 'massachusetts': 'MA', 'vermont': 'VT', 'georgia': 'GA', 'north dakota': 'ND', 'pennsylvania': 'PA', 'puerto rico': 'PR', 'florida': 'FL', 'hawaii': 'HI', 'kentucky': 'KY', 'rhode island': 'RI', 'nebraska': 'NE', 'missouri': 'MO', 'ohio': 'OH', 'alabama': 'AL', 'illinois': 'IL', 'virgin islands': 'VI', 'south dakota': 'SD', 'colorado': 'CO', 'new jersey': 'NJ', 'washington':'WA', 'north carolina': 'NC', 'maine': 'ME', 'new york': 'NY', 'montana': 'MT','nevada': 'NV', 'delaware': 'DE'}
	if headerName.lower().strip() in states:
		final["State"]=states[headerName.lower().strip()]

def isColEmpty(columnVals):
	"""If there's no data in the column return True, if there is data return False

	Parameter: Takes an array of values
	Precondition:inputs can be a double array[['AddressNum', '', '', '']] or a single empty array []
											 ^--these two would return empty   ------------------------^         
	"""

	print "\nExecuting isColEmpty"
	if columnVals==[]: 
		# print "%s is totally empty" %(columnVals)
		return True

	notEmptyCount=0
	for ArrayOrVal in columnVals:
		if type(ArrayOrVal)==list:
			# print "nested array: %s" %(ArrayOrVal)
			for value in ArrayOrVal:
				value=str(value)
			 	if value!="":
			 		# print "the Value: %s" %(value)
			 		notEmptyCount+=1		
	if notEmptyCount<=1:
		return True
	else:
		print "Column has more than 1 value in it, not empty"
	 	return False


def setnwrite (headSubCombined, fileName):
	"""formats data and calls writer to actually wrtie the data"""

	workbook = open_workbook(fileName[0])
	sheet = workbook.sheet_by_index(0)
	amriscCell1 = sheet.cell_value(0,0)
	amriscCell2 = sheet.cell_value(0,1)
	crcSwettCell = sheet.cell_value(10,0)
	amriscID = "AmRisc Application / Schedule of Values"
	crcSwettID = "LOCATION INFORMATION"
	
	workHeaderRow = ['Loc #', 'Bldg #', 'Delete', 'Physical Building #', 'Single Physical Building #', 'Street 1', 'Street 2', 'City', 'State', 'Zip', 'County', 'Validated Zip', 'Building Value', 'Business Personal Property', 'Business Income', 'Misc Real Property', 'TIV', '# Units', 'Building Description', 'ClassCodeDesc', 'Construction Type','Dist. To Fire Hydrant (Feet)', 'Dist. To Fire Station (Miles)', 'Prot Class', '# Stories', '# Basements', 'Year Built', 'Sq Ftg', 'Wiring Year', 'Plumbing Year', 'Roofing Year', 'Heating Year', 'Fire Alarm Type', 'Burglar Alarm Type', 'Sprinkler Alarm Type', 'Sprinkler Wet/Dry', 'Sprinkler Extent', 'Roof Covering', 'Roof Geometry', 'Roof Anchor', 'Cladding Type', 'Roof Sheathing Attachment', 'Frame-Foundation Connection', 'Residential Appurtenant Structures']

	final = {key: [] for key in workHeaderRow}

	work={
		'Loc #':0,
		'Bldg #':1,
		'Delete':2,
		'Physical Building #':3,
		'Single Physical Building #':4,
		'Street 1':5,
		'Street 2':6,
		'City':7,
		'State':8,
		'Zip':9,
		'County':10,
		'Validated Zip':11,
		'Building Value':12,
		'Business Personal Property':13,
		'Business Income':14,
		'Misc Real Property':15,
		'TIV':16,
		'# Units':17,
		'Building Description':18,
		'ClassCodeDesc':19,
		'Construction Type':20,
		'Dist. To Fire Hydrant (Feet)':21,
		'Dist. To Fire Station (Miles)':22,
		'Prot Class':23,
		'# Stories':24,
		'# Basements':25,
		'Year Built':26,
		'Sq Ftg':27,
		'Wiring Year':28,
		'Plumbing Year':29,
		'Roofing Year':30,
		'Heating Year':31,
		'Fire Alarm Type':32,
		'Burglar Alarm Type':33,
		'Sprinkler Alarm Type':34,
		'Sprinkler Wet/Dry':35,
		'Sprinkler Extent':36,
		'Roof Covering':37,
		'Roof Geometry':38,
		'Roof Anchor':39,
		'Cladding Type':40,
		'Roof Sheathing Attachment':41,
		'Frame-Foundation Connection':42,
		'Residential Appurtenant Structures':43
		}

	amrisc={
		"Percent Sprinklered":"Sprinkler Extent",
		"Sprinklered (Y/N)":"Sprinkler Wet/Dry",
		"*Year Roof covering last fully replaced":"Roofing Year",
		"* Bldg No.":"Loc #",
		"*Orig Year Built":"Year Built",
		"*Square Footage":"Sq Ftg",
		"*# of Stories":"# Stories",
		"AddressNum":"Physical Building #",
		"*Street Address":"Street 1", 
		"*City":"City", 
		"*State Code":"State", 
		"*Zip":"Zip", 
		"County":"County", 
		"*Real Property Value ($)":"Building Value", 
		"Personal Property Value ($)":"Business Personal Property", 
		"personal property value  ($)":"Business Personal Property", 
		"Personal Property Value ($) ":"Business Personal Property",
		"Other Value $ (outdoor prop & Eqpt must be sch'd)":"Misc Real Property",
		"BI/Rental Income ($)":"Business Income",
		"*Total TIV":"TIV", 
		"*Occupancy":"Building Description", 
		"Construction Description ":"Construction Type", 
		"Construction Description (provide further details on construction features)":"Construction Type",
		"ISO Prot Class":"Prot Class",
		"*# of Units":"# Units"
		}

	crcSwett = {
		"Loc  #":"Loc #",
		"Location Street Address:":"Street 1",
		"City":"City",
		"State":"State",
		"Zip Code":"Zip",
		"Building Value":"Building Value",
		"Content":"Business Personal Property",
		"BI w/ EE":"Business Income",
		"Total TIV":"TIV",
		"# Apt  Units":"# Units",
		"Building Occupancy":"Building Description",
		"Construction":"Construction Type",
		"# of Stories":"# Stories",
		"Yr Built Gut/Reh":"Year Built",
		"Total Building  Area":"Sq Ftg",
		"Plumbing":"Plumbing Year",
		"Heating":"Heating Year",
		"Electrical":"Wiring Year",
		"Roof":"Roofing Year",
		"Sprinkler %":"Sprinkler Extent"
		}

	
	if amriscCell1.find(amriscID) != -1 or amriscCell2.find(amriscID) != -1:
		minimum= min(headSubCombined, key=headSubCombined.get)-1
		headerRow= headSubCombined[minimum]
		sov_index={}
		for itemIndex in range(len(headerRow)):
			if headerRow[itemIndex] in amrisc:
				sov_index[itemIndex]= True

		columnDict={}
		for index in sov_index:
			column=[]
			for key,item in headSubCombined.iteritems():
				column.append(item[index])
			columnDict[column[0]]=column

		for key in columnDict:
			if amrisc[key] in work:
				final[amrisc[key]].append(columnDict[key])
		final = adjustments(final)

		#Debug
		# with open('E:\Work\Pycel\jonPycel\output\sheetFinal.txt', 'w') as stfnl:
		# 	stfnl.write("The contents of the dictionary final\n\n")
		# 	for key, value in final.iteritems():
		# 		stfnl.write('{0}: {1}\n'.format(key, value))

		writer(final, work, workHeaderRow, amrisc, fileName)
	elif crcSwettCell.find(crcSwettID) != -1:
		minimum= min(headSubCombined, key=headSubCombined.get)-1
		headerRow= headSubCombined[minimum]
		sov_index={}
		for itemIndex in range(len(headerRow)):
			if headerRow[itemIndex] in crcSwett:
				sov_index[itemIndex]= True

		columnDict={}
		for index in sov_index:
			column=[]
			for key,item in headSubCombined.iteritems():
				column.append(item[index])
			columnDict[column[0]]=column

		for key in columnDict:
			if crcSwett[key] in work:
				final[crcSwett[key]].append(columnDict[key])
		final = adjustments(final)
		writer(final, work, workHeaderRow, crcSwett, fileName)

def writer(final, workDict, workHeaderRow, template, sovFileName):
	"""Does the writing"""
	workbook = xlwt.Workbook()
	colWidth = 256 * 20

	# HOW TO DO A CELL OVERWRITE AS A LAST RESORT IF NEEDED
	# sheet = workbook.add_sheet("WKFC_Sheet1", cell_overwrite_ok=True)
	sheet = workbook.add_sheet("WKFC_Sheet1")

	for key, values in final.iteritems():
		colIndex=workDict[key]
		wordWrap = xlwt.easyxf('align: wrap on, horiz center')

		if values==[]:
			sheet.write(0, colIndex, key)
			sheet.col(colIndex).width = 365 * (16)
		else:
			valueArr = values[0]
			for rowIndex in range(len(valueArr)):
				if valueArr[rowIndex] in template:
					sheet.write(rowIndex, colIndex, key, wordWrap)
					sheet.col(colIndex).width = 365 * (16)
				else:
					sheet.write(rowIndex, colIndex, valueArr[rowIndex], wordWrap)
					sheet.col(colIndex).width = 365 * (16)
	
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