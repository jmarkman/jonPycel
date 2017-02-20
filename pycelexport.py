import os
import re
import sys
import pyodbc
import Tkinter
import tkSimpleDialog
from xlrd import open_workbook
from tkFileDialog import askopenfilename

"""
This program is only for exporting the data received from the pycel converter

It take in a pycel converted file, ask for the control number, 
and write the data to a database
"""


def ask():
	"""opens the file explorer allowing user to choose pycel converted SOV to parse"""

	root=Tkinter.Tk()
	root.withdraw()
	userhome = os.path.expanduser('~')
	desktop = userhome + '/Desktop/'
	file = askopenfilename(initialdir=desktop)
	# file = askopenfilename(initialdir="C:\Users\gregory.schultz\Desktop\pycel_1_6_17\examples")
	file=[file]
	print "FILE NAME: " +file[0]
	return file

def openPromptValidate(caption, description):
	"""Prompts and validates control number from the user
	
	Parameter caption: appears as window name
	Parameter description: text that appears above input box """

	val=tkSimpleDialog.askstring(caption,description)
	# value can only be a number with min length 6 and max length 7
	val=re.search(r"^\d{6,7}$", val)
	# if val passes the appropriate regex... (is a match object)
	if val!=None:
		# val is a match object, this is how you access the data
		final=val.group(0)
		return final
	# val is None if no match, bad user input, recurse until correct
	else:
		final=openPromptValidate(caption,"Err! Enter only 6 or 7 digits!")
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


	# TODOS: For some reason the DistToFireStation is failing to input because of the numeric input, find out why
    # 		 Figure out how to put in the keys correctly, throwing an err
	cursor.execute("INSERT INTO pycelSOV (ControlNoIMS, SOVID, LocationNo, BuildingNo, PhysicalBldgNum, SinglePhysicalBldgNum, Address1, Address2, City, State, Zip, County, BuildingValue, BusinessPersonalProperty, BusinessIncome, MiscRealProperty, TIV, Units, BuildingDescription, ClassCodeDesc,  ConstructionType,  DistToFireHydrant, DistToFireStation, ProtectionCode, Stories, Basements, YearBuilt, SqFootage, WiringYear, PlumbingYear, RoofingYear, HeatingYear, FireAlarmType, BurglarAlarmType, SprinklerAlarmType, SprinklerWetDry, SprinklerExtent)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(row.CONTROLNUMBER, 1, row.locNum, row.bldgNum, row.physicalBuild, row.singlePhysical, row.street1, row.street2, row.city, row.state, row.zipCode, row.county, row.buildVal, row.busPers, row.busIncome, row.miscReal, row.TIV, row.numUnits, row.buildDescrip, row.classCodeDesc, row.constType,  row.distFireHydrant, 1111 , row.protClass, row.numStories, row.numBasement, row.yearBuilt, row.sqFtg, row.wireY, row.plumbY, row.roofY, row.heatY, row.fireAlarmType, row.burgAlarmType, row.sprinklerAlarmType, row.sprinklerWetDry, row.sprinklerExtent))	
	# cursor.execute("INSERT INTO pycelSOV (ControlNoIMS, SOVID, LocationNo, BuildingNo, PhysicalBldgNum, SinglePhysicalBldgNum, Address1, Address2, City, State, Zip, County, BuildingValue, BusinessPersonalProperty, BusinessIncome, MiscRealProperty, TIV, Units, BuildingDescription, ClassCodeDesc,  ConstructionType,  DistToFireHydrant, DistToFireStation, ProtectionCode, Stories, Basements, YearBuilt, SqFootage, WiringYear, PlumbingYear, RoofingYear, HeatingYear, FireAlarmType, BurglarAlarmType, SprinklerAlarmType, SprinklerWetDry, SprinklerExtent)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(row.CONTROLNUMBER, 1, row.locNum, row.bldgNum, row.physicalBuild, row.singlePhysical, row.street1, row.street2, row.city, row.state, row.zipCode, row.county, row.buildVal, row.busPers, row.busIncome, row.miscReal, row.TIV, row.numUnits, row.buildDescrip, row.classCodeDesc, row.constType,  row.distFireHydrant, row.distFireStation , row.protClass, row.numStories, row.numBasement, row.yearBuilt, row.sqFtg, row.wireY, row.plumbY, row.roofY, row.heatY, row.fireAlarmType, row.burgAlarmType, row.sprinklerAlarmType, row.sprinklerWetDry, row.sprinklerExtent))	
	# need this commit statement or else nothing goes through, ACID principle
	cnxn.commit()
	print "Statement executed and committed"


def commitToDatabase(records):
	"""connects to database and signals the execution and committing of the record to the database
	
	Parameter records: array of Record objects"""
	try:
		cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.10.11.112;DATABASE=HermesLocationsBuildTest;UID=svc-flexicap;PWD=svcflex')
		# cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.10.11.112;DATABASE=HermesLocationsBuildTest;UID=svc-flexicap;PWD=svcflex')
		print "Connection established"
	except:
		print "Error: Failure to connect to the database, check your internet connection"
		return
	for row in records:
		# only individual record executions
		insertStatement(cnxn, row)
	print "All Statements Executed Successfully, displaying contents:"

	# display total database contents in terminal window for debug
	# show = cnxn.cursor()
	# show.execute("select * from dbo.pycelSOV")	
	# rows = show.fetchall()
	# print "Database Contents: \n "
	# for row in range(len(rows)):
	# 	print "Row number %s : %s" %( row, rows[row])

class Record(object):
	"""" represents each row as an object where each attribute corresponds to a column from the workstation"""
	def __init__(self, locNum, bldgNum, delete,physicalBuild, singlePhysical, street1, street2, city, state, zipCode, county, valZip, buildVal, busPers,
		busIncome, miscReal, TIV, numUnits, buildDescrip, classCodeDesc, constType, distFireHydrant,distFireStation, protClass,numStories,
		numBasement, yearBuilt, sqFtg, wireY,plumbY,roofY,heatY, fireAlarmType, burgAlarmType, sprinklerAlarmType, sprinklerWetDry, 
		sprinklerExtent, roofCovering, roofGeo, roofAnchor, cladType, roofSheath, frameConnection, resAppurtenant):
		self.locNum = locNum
		self.bldgNum = bldgNum
		self.delete = delete
		self.physicalBuild=physicalBuild
		self.singlePhysical =singlePhysical
		self.street1 = street1
		self.street2 = street2
		self.city = city
		self.state = state
		self.zipCode = zipCode
		self.county = county
		self.valZip = valZip 
		self.buildVal = buildVal
		self.busPers = busPers
		self.busIncome = busIncome
		self.miscReal = miscReal
		self.TIV = TIV
		self.numUnits = numUnits
		self.buildDescrip = buildDescrip
		self.classCodeDesc = classCodeDesc
		self.constType = constType
		self.distFireHydrant = distFireHydrant
		self.distFireStation = distFireStation
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
		self.CONTROLNUMBER =None

	def __str__(self):
		return("Record object:\n"
               "  Loc #  = {0}\n"
               "  Building # = {1}\n"
               "  Delete = {2}\n"
               "  Physical Building # = {3}\n"
               "  Single Physical Building # = {4}\n"
               "  Street 1 = {5}\n"
               "  Street 2 = {6}\n"
               "  City = {7}\n"
               "  State = {8}\n"
               "  Zip = {9}\n"
               "  County = {10}\n"
               "  Validated Zip = {11}\n"
               "  Business Value = {12}\n"
               "  Business Personal Property = {13}\n"
               "  Business Income 	= {14}\n"
               "  Misc Real Property = {15}\n"
               "  TIV	= {16}\n"
               "  # of Units = {17}\n"
               "  Building Description = {18}\n"
               "  ClassCodeDesc = {19}\n"
               "  Construction Type = {20}\n"
               "  Dist. to Fire Hydrant (feet) = {21}\n"
               "  Dist. to Fire Station (miles) = {22}\n"
               "  Prot Class = {23}\n"
               "  # of Stories = {24}\n"
               "  # of Basements = {25}\n"
               "  Year Built = {26}\n"
               "  Sq Ftg = {27}\n"
               "  Wiring Year = {28}\n"
               "  Plumbing Year = {29}\n"
               "  Roofing Year = {30}\n"
               "  Heating Year = {31}\n"
               "  Fire Alarm Type = {32}\n"
               "  Burglar Alarm Type = {33}\n"
               "  Sprinkler Alarm Type 	= {34}\n"
               "  Sprinkler Wet/Dry = {35}\n"
               "  Sprinkler Extent 	= {36}\n"
               "  Roof Covering = {37}\n"
               "  Roof Geometry = {38}\n"
               "  Roof Anchor = {39}\n"
               "  Cladding Type = {40}\n"
               "  Roof Sheathing Attachment = {41}\n"
               "  Frame Foundation Connection  = {42}\n"
               "  Residential Appurtenant Structures = {43}\n"
               "  Control Number = {44}\n"
               .format(self.locNum,self.bldgNum, self.delete, self.physicalBuild, self.singlePhysical ,self.street1, self.street2, self.city, self.state, self.zipCode, self.county, self.valZip, self.buildVal, self.busPers, self.busIncome, self.miscReal, self.TIV, self.numUnits, self.buildDescrip, self.classCodeDesc, self.constType, self.distFireHydrant, self.distFireStation, self.protClass, self.numStories, self.numBasement, self.yearBuilt, self.sqFtg, self.wireY, self.plumbY, self.roofY, self.heatY, self.fireAlarmType, self.burgAlarmType, self.sprinklerAlarmType, self.sprinklerWetDry, self.sprinklerExtent, self.roofCovering, self.roofGeo, self.roofAnchor, self.cladType, self.roofSheath, self.frameConnection, self.resAppurtenant, self.CONTROLNUMBER)
               )

def getRecords(file,controlNum):
	"""pulls data from pycel converted SOV

	Parameter file: the user selected file in a doubly nested array
	Parameter controlNum: Validated user entered IMS controlNumber

	Returns: Array of record objects"""

	wb = open_workbook(file[0])
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

	# Terminal display
	for rowIndex in range(len(records)):
		print "Row %s:\n %s" %(rowIndex, records[rowIndex])

		# Now everything is accessible as an object (yes!), so for example to get the street1 do records[rowIndex].street1
		# Easy for finding the value for the SQL kickout

		# or if you want to just print out everything...
		# for attr, value in records[rowIndex].__dict__.iteritems():
			# print attr, value

	return records


def run():
	"""Master caller"""
	file=ask()
	controlNum=getControlNumber()
	records=getRecords(file,controlNum)
	commitToDatabase(records)

# Lego
run()


