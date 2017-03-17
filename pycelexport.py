import os
import re
import sys
import pyodbc
import uuid
import Tkinter
import tkSimpleDialog
from xlrd import open_workbook
from tkFileDialog import askopenfilename
from tkinter import messagebox

"""
This program is only for exporting the data received from the pycel converter

It take in a pycel converted file, ask for the control number, 
and write the data to a database
"""

def openPromptValidate(caption, description):
	"""Prompts and validates control number from the user
	
	Parameter caption: appears as window name
	Parameter description: text that appears above input box """
	val = tkSimpleDialog.askstring(caption,description)
	# value can only be a number with min length 6 and max length 7
	val = re.search(r"^\d{6,7}$", val)
	# if val passes the appropriate regex... (is a match object)
	if val!=None:
		# val is a match object, this is how you access the data
		final=val.group(0)
		return final
	# val is None if no match, bad user input, recurse until correct
	else:
		messagebox.showinfo(message="The control number entered was not valid as input.\n Check your input and try again.\n\nIf problems persist, send an email to: support@rsgta.zohosupport.com")
		final=openPromptValidate(caption,description)
	return final

def getControlNumber():
	controlNumEntered =  openPromptValidate("Control Number?", "Enter the Control Number")
	return controlNumEntered

def insertStatement(cnxn,row):
	"""Inserts individual record into pycelSOV table
	
	Parameter cnxn: connection object connected to database:HermesLocationsBuildTest with correct write permissions
	Parameter row: a single Record object"""
	# print row.distFireStation
	cursor = cnxn.cursor()

	cursor.execute("INSERT INTO pycelSOV (ControlNoIMS, LocationNo, BuildingNo, StreetAddressRaw, PhysicalBldgNum, SinglePhysicalBldgNum, Address1, Address2, City, State, Zip, County, BuildingValue, BusinessPersonalProperty, BusinessIncome, MiscRealProperty, Units, BuildingDescription, ClassCodeDescRaw, ConstructionTypeRaw, ProtectionCode, Stories, Basements, YearBuilt, SqFootage, WiringYear, PlumbingYear, RoofingYear, HeatingYear, FireAlarmTypeRaw, BurglarAlarmTypeRaw, SprinklerAlarmTypeRaw, SprinklerWetDryRaw, SprinklerExtentRaw) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (row.CONTROLNUMBER, row.locNum, row.bldgNum, row.rawAddr, row.physicalBuild, row.singlePhysical, row.street1, row.street2,  row.city, row.state, row.zipCode, row.county, row.buildVal, row.busPers, row.busIncome, row.miscReal, row.numUnits, row.buildDescrip, row.classCodeDesc, row.constType, row.protClass, row.numStories, row.numBasement, row.yearBuilt, row.sqFtg, row.wireY, row.plumbY, row.roofY, row.heatY, row.fireAlarmType, row.burgAlarmType, row.sprinklerAlarmType, row.sprinklerWetDry, row.sprinklerExtent))
	# need this commit statement or else nothing goes through, ACID principle
	cnxn.commit()
	print "Statement executed and committed"


def commitToDatabase(records):
	"""connects to database and signals the execution and committing of the record to the database
	
	Parameter records: array of Record objects"""
	try:
		# cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.10.11.112;DATABASE=ABBYY_AppData;UID=svc-flexicap;PWD=svcflex')
        	cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.0.0.46;DATABASE=ABBYY;UID=abbyy01;PWD=20@dEvAbbYY17!')
		print "Connection established"
	except:
		print "Error: Failure to connect to the database, check your internet connection"
		return
	for row in records:
		# only individual record executions
		insertStatement(cnxn, row)

class Record(object):
	"""" represents each row as an object where each attribute corresponds to a column from the workstation"""
	def __init__(self, locNum, bldgNum, rawAddr, physicalBuild, singlePhysical, street1, street2, city, state, zipCode, county, buildVal, busPers, busIncome, miscReal, numUnits, buildDescrip, classCodeDesc, constType, protClass,numStories, numBasement, yearBuilt, sqFtg, wireY,plumbY,roofY,heatY, fireAlarmType, burgAlarmType, sprinklerAlarmType, sprinklerWetDry, sprinklerExtent, roofCovering, roofGeo, roofAnchor, cladType, roofSheath, frameConnection, resAppurtenant):
		self.locNum = locNum
		self.bldgNum = bldgNum
		self.rawAddr = rawAddr
		self.physicalBuild = physicalBuild
		self.singlePhysical = singlePhysical
		self.street1 = street1
		self.street2 = street2
		self.city = city
		self.state = state
		self.zipCode = zipCode
		self.county = county
		self.buildVal = buildVal
		self.busPers = busPers
		self.busIncome = busIncome
		self.miscReal = miscReal
		self.numUnits = numUnits
		self.buildDescrip = buildDescrip
		self.classCodeDesc = classCodeDesc
		self.constType = constType
		self.protClass = protClass
		self.numStories = numStories
		self.numBasement = numBasement
		self.yearBuilt = yearBuilt
		self.sqFtg = sqFtg
		self.wireY = wireY
		self.plumbY = plumbY
		self.roofY= roofY
		self.heatY = heatY
		self.fireAlarmType = fireAlarmType
		self.burgAlarmType = burgAlarmType
		self.sprinklerAlarmType = sprinklerAlarmType
		self.sprinklerWetDry = sprinklerWetDry
		self.sprinklerExtent= sprinklerExtent
		self.roofCovering = roofCovering
		self.roofGeo = roofGeo
		self.roofAnchor = roofAnchor
		self.cladType = cladType
		self.roofSheath = roofSheath
		self.frameConnection = frameConnection
		self.resAppurtenant = resAppurtenant
		# to be inputted by user
		self.CONTROLNUMBER = None

	def __str__(self):
		return("Record object:\n"
               "  Loc #  = {0}\n"
               "  Building # = {1}\n"
               "  Raw Address = {2}\n"
               "  Physical Building # = {3}\n"
               "  Single Physical Building # = {4}\n"
               "  Street 1 = {5}\n"
               "  Street 2 = {6}\n"
               "  City = {7}\n"
               "  State = {8}\n"
               "  Zip = {9}\n"
               "  County = {10}\n"
               "  Business Value = {11}\n"
               "  Business Personal Property = {12}\n"
               "  Business Income 	= {13}\n"
               "  Misc Real Property = {14}\n"
               "  # of Units = {15}\n"
               "  Building Description = {16}\n"
               "  ClassCodeDesc = {17}\n"
               "  Construction Type = {18}\n"
               "  Prot Class = {19}\n"
               "  # of Stories = {20}\n"
               "  # of Basements = {21}\n"
               "  Year Built = {22}\n"
               "  Sq Ftg = {23}\n"
               "  Wiring Year = {24}\n"
               "  Plumbing Year = {25}\n"
               "  Roofing Year = {26}\n"
               "  Heating Year = {27}\n"
               "  Fire Alarm Type = {28}\n"
               "  Burglar Alarm Type = {29}\n"
               "  Sprinkler Alarm Type 	= {30}\n"
               "  Sprinkler Wet/Dry = {31}\n"
               "  Sprinkler Extent 	= {32}\n"
               "  Roof Covering = {33}\n"
               "  Roof Geometry = {34}\n"
               "  Roof Anchor = {35}\n"
               "  Cladding Type = {36}\n"
               "  Roof Sheathing Attachment = {37}\n"
               "  Frame Foundation Connection  = {38}\n"
               "  Residential Appurtenant Structures = {39}\n"
               "  Control Number = {40}\n"
               .format(self.locNum,self.bldgNum, self.rawAddr, self.physicalBuild, self.singlePhysical ,self.street1, self.street2, self.city, self.state, self.zipCode, self.county, self.buildVal, self.busPers, self.busIncome, self.miscReal, self.numUnits, self.buildDescrip, self.classCodeDesc, self.constType, self.protClass, self.numStories, self.numBasement, self.yearBuilt, self.sqFtg, self.wireY, self.plumbY, self.roofY, self.heatY, self.fireAlarmType, self.burgAlarmType, self.sprinklerAlarmType, self.sprinklerWetDry, self.sprinklerExtent, self.roofCovering, self.roofGeo, self.roofAnchor, self.cladType, self.roofSheath, self.frameConnection, self.resAppurtenant, self.CONTROLNUMBER)
               )

def getRecords(file, controlNum):
	"""pulls data from pycel converted SOV

	Parameter file: the user selected file in a doubly nested array
	Parameter controlNum: Validated user entered IMS controlNumber

	Returns: Array of record objects"""

	wb = open_workbook(str(file))
	records = []
	for sheet in wb.sheets():
	    number_of_rows = sheet.nrows
	    number_of_columns = sheet.ncols
	    # use these ^^ because we know the data is not too large
	    rows = []
	    for row in range(1, number_of_rows):
	        values = []
	        for col in range(number_of_columns):
	            value  = (sheet.cell(row,col).value)
	            try:
	                value = str(int(value))
	            except ValueError:
	                pass
	            finally:
	                values.append(value)
			
	        # Takes found data and converts it into a Record object        
	        fullRow = Record(*values)
	        # apply the inputted control number to the instance
	        fullRow.CONTROLNUMBER=controlNum
	        # container for all the Record objects
	        records.append(fullRow)
	return records


def run(pycelFile):
	"""Master caller"""
	file = pycelFile
	controlNum = getControlNumber()
	records = getRecords(pycelFile, controlNum)
	commitToDatabase(records)

