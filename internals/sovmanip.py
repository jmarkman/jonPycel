import os
import sys
import xlwt
import Tkinter
import logging
import subprocess
from xlrd import open_workbook
from unidecode import unidecode
from tkFileDialog import askopenfilename

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
def ask():
	"""opens the file explorer allowing user to choose SOV to parse

	Precondition: File is an SOV with a .xlsx or .xls file format
	Returns: a singular item array(for possibility to later loop through multiple files, TBD) with the aboslute path of the file
	
	This function uses Tkinter to present the user with askopenfilename() to get the SoV. The function returns said file to the rest of the program.	
	"""

	root=Tkinter.Tk()
	root.withdraw()
	file = askopenfilename(initialdir = userhome)
	file = [file]
	print "FILE NAME: " + file[0]
	return file

def findSheetName(files):
	"""find the sheet with the SOV through hierarchy guessing

	Parameter file: array of absolute path
	Returns: Sheet Object where the SOV should be on

	This function takes the returned file from ask() and finds the sheet where all the info we need will be. It starts specifically and ends with simply accessing the first sheet in the workbook. It returns the sheet that meets one of these conditions in the try/except nest.
	"""
	# has to guess through them
	for item in files:
		print ("\n-------------------------------NEW FILE ---------------------------")
		wb=open_workbook(item)
		try:
			sheet1=wb.sheet_by_name("SOV")
		except:
			try:
				sheet1=wb.sheet_by_name("SOV-APP")
			except:
				try:
					sheet1=wb.sheet_by_name("AmRisc SOV")
				except:	
					try:
						sheet1=wb.sheet_by_name("BREAKDOWN")	
					except:
						try:
							sheet1=wb.sheet_by_name("Property Schedule")
						except:
							try:
								sheet1=wb.sheet_by_name("2015 Schedule")
							except:
								try:
									sheet1=wb.sheet_by_name("Sheet1")
								except:
									sheet1=wb.sheet_by_index(0)
		return sheet1
	
def loopAllRows(sheet):
	"""Loops through all rows in the current sheet, formats the data

	Parameter: A sheet object
	Returns: Pure value data from the current sheet in {rowNumber: ['value1','value2','value3']} format
	
	The function starts with a solid definition for the number of rows in the sheet. Originally the ROWCAP variable used to be the total number of rows in the sheet via the sheet.nrows function, but was set to a static number for size/speed concerns. This will be used for our sentinel value for the definite loop used to iterate through the sheet returned from findSheetName(). After, it creates an empty dictionary assigned to the variable "totalPureData", where we will store an association with a number (zero-based) and a row's contents - see Greg's format comment above the declaration of the dictionary.

	Once that is done, we launch into the loop. While the sentinel value "i" is less than the value assigned to ROWCAP, and if "i" is less than ROWCAP, store the row object at position sheet.row(i) in the variable "row". Create an array called "rowPureData" and for each cell object in the row object, store the value of that cell object in the array "rowPureData". 

	Finally, assign the contents of that array to the key "i" in the dictionary "totalPureData". Increment "i" and continue the loop until the condition is met. Once the condition is met, return the dictionary.	
	"""
		
	#ROWCAP=212
	#maximum number of rows in the current sheet
	ROWCAP=sheet.nrows
	# jk for amrisc just cap at row 200
	# SET THIS TO CHANGE THE ROW LOOP CAP
	# ROWCAP=50


	# format:   {0:['data1','data2'....]}
	totalPureData = {}
	i = 0
	while i < ROWCAP:
		#max number of rows in document for diagnostics
		if (i < (ROWCAP)):
			"""
			Debug: print the current row number the loop is on to stdout

			# print "\nCurrent Row Number: " + str(i+1)
			
			These are Cell objects, see xlrd documentation. Addendum by Jon: it's not just limited to xlrd, but is key to how Excel works. Each cell in the workbook is an object with a long list of attributes. When we enter a line of text, we're using a mutator to change the attributes of the cell object by adding a value to the "string" part of the cell object, and the format options in Excel try to cast an attribute of one type to another attribute of another type.
			"""	
			row = sheet.row(i)
			# holds this specific rows pure data
			rowPureData = []
			for item in row:
				rowPureData.append(item.value)
			totalPureData[i] = rowPureData
		i += 1
	"""
	Debug: Print the contents of the totalPureData dictionary to see what's inside of it

	# for key,value in totalPureData.iteritems():
	# 	print key,value
	"""
	# with open('E:\Work\Pycel\pycel2\output\loopAllRows().txt', 'w') as tpd:
	# 	tpd.write("The contents of the dictionary totalPureData\n\n")
	# 	for key, value in totalPureData.iteritems():
	# 		tpd.write('{0}: {1}\n'.format(key, value))
	return totalPureData

def identifyHeaderRow(numValsDict,comparisonDic):
	"""Identifies the header row

	Parameter: Pure sheet data in a dictionary where key=rowNumber value=value array
	Precondition: Header captions are in a singular row, and within the first 30 rows of the file
	
	maxCap refers to the total number of columns in the workstation that we can be sure of getting on the SoV. rowNumber is set to 1 because Excel works on a 1-based index instead of a programming-based index of 0. headerRowNum most likely deals with the dictionary numValsDict. 

	The loop says "while the rowNumber is less than the maxCap and that same rowNumber is less than the length of the dictionary that was passed in as an argument to the identifyHeaderRow function:
		- set the variable "row" equal to the element at [key] from the numValsDict dictionary
		- declare a counter variable called "matchCounter" and set it to 0
		- for each value in the row supplied, make sure that if it's in unicode, use the unidecode library to pre-clean the input just in case there's some rogue input
		- after cleaning that value, if that value exists within the comparisonDic (key-value relationship), increment matchCounter by 1
		- at that point, if matchCounter is greater than headerRowNum, set headerRowNum equal to rowNumber
		- finally, increment rowNumber by 1
	
	Once the while loop runs its course, we spawn a new dictionary and assign the key to headerRowNum and put stuff from numValsDict inside of it. Finally, return that dictionary because that contains all of our header row numbers and values.
	
	"""

	# THIS HINGES ON THIS SUBSTITUTION DICTIONARY
	# TO SEE THE MATCHING ENABLE ALL PRINT STATEMENTS
	# This matching is an assumption

	maxCap = 30 
	rowNumber = 1
	headerRowNum = 0 # 
	while ((rowNumber < maxCap) and (rowNumber < len(numValsDict))):
		row=numValsDict[rowNumber]
		matchCounter = 0
		for value in row:
			# because the dictionary is strings and lowercase keys and to deal with unicode
			if type(value) == unicode:
				# print "converting unicode"
				value=unidecode(value)
			value=str(value).lower()

			# print "judging value: " +value
			if value in comparisonDic:
				matchCounter+=1
				# print "matched row num: " +str(rowNumber)+" value: "+ value
		if matchCounter>headerRowNum:
			headerRowNum=rowNumber
		
		# increment
		rowNumber+=1
		# print "current Header row number: " +str(headerRowNum)


	# print headerRowNum  #the header row number NOTE THIS IS 0 INDEXED
	# print numValsDict[headerRowNum]    #data at headerRowNum

	headerNum_Vals={}
	headerNum_Vals[headerRowNum]=numValsDict[headerRowNum]
	# print "\n\nHEADER ROW NUM AND VALUES"
	# print headerNum_Vals
	return headerNum_Vals	

def isEmptyRow(row):
	"""If row is entirely empty 

	Parameter: array of values ['one','two','three']
	"""
	
	maxLength = len(row)
	emptyCheck = []
	# print row
	for value in row: # for each value in the array "row"
		if value == '':
			emptyCheck.append(value)
	# print maxLength, len(emptyCheck)
	if len(emptyCheck) == maxLength:
		# print "Whole row empty"
		return True

def sliceSubHeaderData(headerRowDict, sheet):
	"""
	Gets relevant data below header row

	Parameter: headerRowDict: dictionary,
	Parameter sheet: a Sheet object

	Returns: data from below the header row
	"""

	dataRowStartNumber=headerRowDict.keys()[0]+1
	allData=loopAllRows(sheet)
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

	headSubCombined=headerRowDict.copy()
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
		val=comparisonDic[key]
		key=decompress(key)
		new[key]= val
	return new

def decompress(value):
	"""does the reconstruction of each value
	
	Greg ended up just setting each valued passed as an argument to the function to be cast to a string.

	Returns: param value returned as a string.
	"""
	value=str(value)
	# value=''.join(value.split())
	# value=value.replace("*", "")
	# value=value.replace("/", "")
	# value=value.replace("-", "")
	# value=value.replace(".", "")
	# value=value.lower().strip()
	return value


def head_matcher(compressedDict, headerRow, fileName):
	'''
	Parameters: compressedDict, headerRow, fileName
		- compressedDict is the comparison dictionary
		- headerRow is the dictionary returned from identifyHeaderRow
		- fileName is the file we're working with, see ask()

		This function basically takes in the comparison dictionary and the header dictionary generated from identifyHeaderRow and begins the matching process. The function starts by creating a new array called decompressedRow.

		For each header item in headerRow, using an itervalues().next() to go through the dictionary, use the decompress function to sanitize each item. Then, if that header item exists in the comparison dictionary that was run through comp_converter(), add that header to the decompressedRow array. Otherwise, add a distinguishing label to decompressedRow to mark that the header wasn't around.

		Finally, return headerRow
	'''
	decompressedRow=[]
	# file=open('unmatches.txt','a+')
	for header in headerRow.itervalues().next():
		header=decompress(header)
		if header in compressedDict:
			decompressedRow.append(compressedDict[header])
		else:
			# print ("Not matched %s" %header)
			# file.write (str(fileName[0]))
			# file.write("\n"+header+ "\n")
			decompressedRow.append("XX"+header)
	# print headerRow
	# for key in headerRow:
		# headerRow[key]=decompressedRow
	with open('E:\Work\Pycel\pycel2\output\headerRow_after.txt', 'w') as decompRow:
		for header in decompressedRow:
			decompRow.write('{0}'.format(header) + ' ')
	return headerRow


def adjustments(final):
	"""
	MASTER CALLER FOR ADJUSTMENETS
	Adjusts the data
	"""

	# for key, value in final.iteritems():
	# 	print key,value
	print final['Physical Building #']
	print final['Single Physical Building #'] 
	print final['Single Physical Building #'] ==[]

	locNumFix(final,'Loc #')

	if final['Physical Building #']==[] or isColEmpty(final['Physical Building #'])==True:physicalBuildingNum(final, 'Physical Building #')
	if final['Single Physical Building #']==[] or isColEmpty(final['Single Physical Building #'])==True:physicalBuildingNum(final, "Single Physical Building #")
	
	street1Fix(final, "Street 1")

	if final["State"]==[] or isColEmpty(final['State'])==True: statesConverter(final, "State")

	if final["Wiring Year"]==[] or isColEmpty(final['Wiring Year'])==True:     wprhY(final, "Wiring Year")
	if final["Plumbing Year"]==[]or isColEmpty(final['Plumbing Year'])==True:  wprhY(final, 'Plumbing Year')
	if final["Roofing Year"]==[]or isColEmpty(final['Roofing Year'])==True:    wprhY(final, 'Roofing Year')
	if final["Heating Year"]==[]or isColEmpty(final['Heating Year'])==True:   wprhY(final, 'Heating Year')

	if final["Fire Alarm Type"]==[] or isColEmpty(final['Fire Alarm Type'])==True: nonePlacer(final, "Fire Alarm Type")
	if final["Burglar Alarm Type"]==[] or isColEmpty(final['Burglar Alarm Type'])==True: nonePlacer(final, "Burglar Alarm Type")
	if final["Sprinkler Alarm Type"]==[] or isColEmpty(final['Sprinkler Alarm Type'])==True: nonePlacer(final, "Sprinkler Alarm Type")
	# sprinkWetDry(final)
	sprinkExtent(final)
	populationCounter(final, 'State')

	return final

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

	sprinkExtent=final['Sprinkler Extent'][0]
	for itemIndex in range(len(sprinkExtent)):
		if sprinkExtent[itemIndex]==0.0:
			sprinkExtent[itemIndex]="None"

def street1Fix(final, street1):
	"""strips off the number from the street if it is there

	FIGURE OUT HOW TO DO STREET ABBREVIATIONS SWITCHES"""
	street1=final[street1][0]
	print street1
	for index in range(len(street1)):
		space=street1[index].find(' ')
		posNumber=street1[index][:space]
		try:
			# there is a number
			posNumber=posNumber.replace('-','')
			print "i am the possible number %s" %(posNumber)
			# if this fails it will catch
			# int(posNumber)
			# print street1[index][space:].strip()
			street1[index]=street1[index][space:].strip()
			# final[street1[index]]=final[street1[index]][space:]
		except:
			# no number there
			pass
	street1[0]="Street 1"


def physicalBuildingNum(final, caption):
	"""Identifies the number associated with the street1 and populates a colum with this number
	 if this column is Single Physical Building Number, it will copy Physical Building #
	
	 TODO test this more"""

	print 'PHYSICAL BULDING NUMBER IS RUNNING '
	print caption
	try:
		streetArr=final["Street 1"][0][:]
		# print streetArr, caption
		streetArr.pop(0)
		numTracker=[caption]
		for val in streetArr:
			if len(val)>0:
				space=val.find(" ")
				dash=val.find("-")
				# dash is found
				if dash!=-1:
					num=val[:dash]
				else:
					num=val[:space]
				print "the numval %s" %(num)
				# if this piece is actually a number
				try:
					# if num is actually a number
					int(num)
					numTracker.append(str(num))
				except ValueError:
					# shouldnt append if its not a number
					pass
		# Contents of Physcial Building # becomes Single Physical Building #
		# -----Is this how it should work in terms of business rules?
		if caption=='Single Physical Building #':
			# print 'EXECUTING'
			final[caption]=final['Physical Building #'][:]
		else:
			final[caption]=[numTracker]
		print final[caption]
	except:
		pass
		
def populationCounter(final, caption):
	"""returns the number of non empty rows filled in the given column"""
	print "POPULATION COUNTER RUNNING"
	column=final[caption][0]
	nonEmptyCount=0
	for item in column:
		if item!='':
			nonEmptyCount+=1
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


	# # print 'EXECUTING NONEPLACER'
	# rowToAppend=[]
	# # add the header name to index 0 of to be appended column
	# rowToAppend.append(headerName)
	# # assumption that state is going to be full
	# for row in range(len(final["State"][0])):
	# 	rowToAppend.append("None")
	# # drop a None because of the offset with the headerName
	# rowToAppend.pop(-1)
	# # realign
	# final[headerName].append(rowToAppend)
	# # print final[headerName]

def wprhY(final, headerName):
	"""For wiring, plumbing, roofing, heating year. If emtpy fill their contents with the 
	data from Year Built, else does nothing """

	if final['Year Built']!=[]:
		# copy the list because pointers
		arr=final['Year Built'][:]
		final[headerName]=arr
		# no need to return, modifes the hashtable directly

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
		print "%s is totally empty" %(columnVals)
		return True

	# if it's a double array
	notEmptyCount=0
	for ArrayOrVal in columnVals:
		if type(ArrayOrVal)==list:
			print "nested array: %s" %(ArrayOrVal)
			for value in ArrayOrVal:
				value=str(value)
			 	if value!="":
			 		print "the Value: %s" %(value)
			 		notEmptyCount+=1		
	if notEmptyCount<=1:
		return True
	else:
		print "Column has more than 1 value in it, not empty"
	 	return False


def setnwrite (headSubCombined, fileName):
	"""formats data and calls writer to actually wrtie the data"""

	workHeaderRow = ['Loc #', 'Bldg #', 'Delete', 'Physical Building #', 'Single Physical Building #', 'Street 1', 'Street 2', 'City', 'State', 'Zip', 'County', 'Validated Zip', 'Building Value', 'Business Personal Property', 'Business Income', 'Misc Real Property', 'TIV', '# Units', 'Building Description', 'ClassCodeDesc', 'Construction Type','Dist. To Fire Hydrant (Feet)', 'Dist. To Fire Station (Miles)', 'Prot Class', '# Stories', '# Basements', 'Year Built', 'Sq Ftg', 'Wiring Year', 'Plumbing Year', 'Roofing Year', 'Heating Year', 'Fire Alarm Type', 'Burglar Alarm Type', 'Sprinkler Alarm Type', 'Sprinkler Wet/Dry', 'Sprinkler Extent', 'Roof Covering', 'Roof Geometry', 'Roof Anchor', 'Cladding Type', 'Roof Sheathing Attachment', 'Frame-Foundation Connection', 'Residential Appurtenant Structures']

	final = {key: [] for key in workHeaderRow}

	work={'Loc #':0,'Bldg #':1,'Delete':2,'Physical Building #':3,'Single Physical Building #':4,'Street 1':5,'Street 2':6,'City':7,'State':8,'Zip':9,'County':10,'Validated Zip':11,'Building Value':12,'Business Personal Property':13,'Business Income':14,'Misc Real Property':15,'TIV':16,'# Units':17,'Building Description':18,'ClassCodeDesc':19,'Construction Type':20,'Dist. To Fire Hydrant (Feet)':21,'Dist. To Fire Station (Miles)':22,'Prot Class':23,'# Stories':24,'# Basements':25,'Year Built':26,'Sq Ftg':27,'Wiring Year':28,'Plumbing Year':29,'Roofing Year':30,'Heating Year':31,'Fire Alarm Type':32,'Burglar Alarm Type':33,'Sprinkler Alarm Type':34,'Sprinkler Wet/Dry':35,'Sprinkler Extent':36,'Roof Covering':37,'Roof Geometry':38,'Roof Anchor':39,'Cladding Type':40,'Roof Sheathing Attachment':41,'Frame-Foundation Connection':42,'Residential Appurtenant Structures':43}

	amrisc={"Percent Sprinklered":"Sprinkler Extent","Sprinklered (Y/N)":"Sprinkler Wet/Dry","*Year Roof covering last fully replaced":"Roofing Year", "* Bldg No.":"Loc #","*Orig Year Built":"Year Built","*Square Footage":"Sq Ftg","*# of Stories":"# Stories","AddressNum":"Physical Building #", "*Street Address":"Street 1", "*City":"City", "*State Code":"State", "*Zip":"Zip", "County":"County", "*Real Property Value ($)":"Building Value", "Personal Property Value ($)":"Business Personal Property","Other Value $ (outdoor prop & Eqpt must be sch'd)":"Misc Real Property","BI/Rental Income ($)":"Business Income", "*Occupancy":"Building Description", "Construction Description ":"Construction Type", "Construction Description (provide further details on construction features)":"Construction Type","ISO Prot Class":"Prot Class","*# of Units":"# Units"}

	crcSwett = {"Loc  #":"Loc #","Location Street Address:":"Street 1","City":"City","State":"State","Zip Code":"Zip","Building Value":"Building Value","Content":"Business Personal Property","BI w/ EE":"Business Income","Total TIV":"TIV","# Apt  Units":"# Units","Building Occupancy":"Building Description","Construction":"Construction Type","# of Stories":"# Stories","Yr Built Gut/Reh":"Year Built","Total Building  Area":"Sq Ftg","Plumbing":"Plumbing Year","Heating":"Heating Year","Electrical":"Wiring Year","Roof":"Roofing Year","Sprinkler %":"Sprinkler Extent"}
	# AMRISC TODO: workstation columns that don't have a 1:1 with an amrisc won't autoflip, fill these with whatever is needed (most of the last columns)

	# ADD FILE TEMPLATES HERE, TODO: FIND A WAY TO SELECT TEMPLATES OR AUTOFIND THEM

	minimum= min(headSubCombined, key=headSubCombined.get)-1
	headerRow= headSubCombined[minimum]
	#sov_index will look like - <SOVIndex>:True - if that columns needs to end up in the workstation
	sov_index={}
	# Identify Fixed SOVHeader Rows needed to switch
	for itemIndex in range(len(headerRow)):
		# if header value belongs in workstation
		if headerRow[itemIndex] in amrisc:
			# flag the column index to be appended
			sov_index[itemIndex]= True

	# compresses sov data that needs to be written into column format
	columnDict={}
	for index in sov_index:
		column=[]
		for key,item in headSubCombined.iteritems():
			column.append(item[index])
		columnDict[column[0]]=column

	# combines SOV columns with workstation columns
	for key in columnDict:
		if amrisc[key] in work:
			final[amrisc[key]].append(columnDict[key])

	# does the adjustments
	final = adjustments(final)
	# for notice can use isColEmpty to see if column is empty(defined as header but no data)

	# Jon: added file name to writer params for usage in writer()'s save function since there
	# was no way to access the name of the spreadsheet otherwise in that method. Maybe this can
	# be set to a global variable 
	writer(final, work, workHeaderRow, amrisc, fileName)


def writer(final, workDict, workHeaderRow, amrisc, sovFileName):
	"""Does the writing"""
	workbook = xlwt.Workbook() 

	# HOW TO DO A CELL OVERWRITE AS A LAST RESORT IF NEEDED
	# sheet = workbook.add_sheet("WKFC_Sheet1", cell_overwrite_ok=True)
	sheet = workbook.add_sheet("WKFC_Sheet1")

	# see all the data
	for key, values in final.iteritems():
		colIndex=workDict[key]
		wordWrap = xlwt.easyxf('align: wrap on, horiz center') # provides access to formatting features

		if values==[]:
			# if no data found in SOV,  put the workstation header name instead
			sheet.write(0,colIndex,key, wordWrap)
			sheet.col(colIndex).width = 365 * (16)
		else:
			# because values is a double array
			valueArr = values[0]
			for rowIndex in range(len(valueArr)):
				# write the header from Workstation rather than SOV
				if valueArr[rowIndex] in amrisc:
					# print "I matched in the template!"
					sheet.write(rowIndex,colIndex, key, wordWrap)
					sheet.col(colIndex).width = 365 * (16)
				else:
					# write the rest of the data
					sheet.write(rowIndex,colIndex, valueArr[rowIndex], wordWrap)
					sheet.col(colIndex).width = 365 * (16)
	
	# Notes on word wrap: There is no default function in xlwt for automatically setting the width of a column
	# However, according to a StackOverflow answer, the default length of a column is 2962 units, which Excel
	# recognizes as 8.11 units. Divide 2962 by 8.11 and the answer is somewhere around the 365-367 range. So
	# for the easiest and most accurate results without output looking hella weird, we can use sheet.col().width
	# and set it to 365 * 16 to in essence double the size of the column so every header fits to an aesthetic extent	

	# Check if the file exists
	sovCheck = os.path.isfile(sovFileName[0])
	if sovCheck == False: # if it doesn't...
		workbook.save(sovFileName[0]) # save the file
		print "FILE WRITTEN"
		os.startfile(userhome + sovFileName[0]) # open the file 
		print "FILE NOW OPEN"
	else:
		sovName = findFileName(sovFileName[0])
		precedingName = '[Pycel_Extracted]_'
		fileType = getFileExtension(sovFileName[0])
		newFileName = precedingName + sovName + fileType # build a new filename 
		workbook.save(userhome + newFileName)
		print "FILE WRITTEN"
		os.startfile(userhome + newFileName)
		print "FILE NOW OPEN"
		

#MAIN CONTROLLER FOR THE PROGRAM
def run():
	"""
	Master Controller
	
	Contains the following functions:
		- ask()
		- findSheetName()
		- loopAllRows()
		- identifyHeaderRow()
		- head_matcher()
		- sliceSubHeaderData()
		- combine()
		- setnwrite()

	Contains the megadict called comparisonDic. This is where all the magic happens and is similar to the main method in C#/Java/etc.
	
	"""
	# clear the terminal screen
	clear = lambda: os.system('cls')
	clear()

	fileName = ask()
	sheet = findSheetName(fileName)
	data = loopAllRows(sheet)


	headerRow = identifyHeaderRow(data,comparisonDic)
	headerRow = head_matcher(comp_converter(comparisonDic), headerRow, fileName)
	subHeadData = sliceSubHeaderData(headerRow, sheet)
	headSubCombined = combine(headerRow,subHeadData)
	setnwrite(headSubCombined,fileName)
	print "Original fileName %s" % (fileName)


run()