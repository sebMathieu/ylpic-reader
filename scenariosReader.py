## Classes and methods needed to convert the series of Excel files with the scenarios data
# Requires xlrd which can be installed with "pip3 install xlrd".
#@author Sebastien MATHIEU

import xlrd
import datetime, math
import networkMaker

# Constants
## File with the information on the scenarios.
SCENARIOS_FILE="scenarios GREDOR_substation V3.xlsx"
## File with the calendar information.
CALENDAR_FILE="YearCalendar 2015-2020-2030-2050 for profile.xlsx"
## File with the load profiles/
LOAD_PROFILES_FILE="catalogue charge V3.xlsx"

# Buffers of excel files and their path
networkMaker.BUFFER_CALENDAR = None
networkMaker.BUFFER_CALENDAR_FILE = None
networkMaker.BUFFER_PROFILES = None
networkMaker.BUFFER_PROFILES_FILE = None
networkMaker.BUFFER_SCENARIOS = None
networkMaker.BUFFER_SCENARIOS_FILE = None

## Add loads to the corresponding network graph from the scenarios files.
# @param folderPath Path to the folder with the Excel data files.
# @param year Year of the scenarios.
# @param day Day of the year.
# @param graph Multigraph to add the loads.
# @param scenarioType Type of scenario. Usually 'L' or 'H'.
def readScenarios(folderPath,year,day,graph,scenarioType='H'):
	readScenariosExcel(folderPath+SCENARIOS_FILE,year,graph,scenarioType)

	profilesType=readCalendarExcel(folderPath+CALENDAR_FILE,year,day)
	readLoadProfilesExcel(folderPath+LOAD_PROFILES_FILE,day,graph,profilesType)
	#networkMaker.detectEndBuses(graph)

## Read and attached the load profiles to each buses.
# @param filePath Path to the load profiles excel file.
# @param day Day of the year.
# @param graph Graph to add the loads.
# @param profilesType Type of profiles for the day.
def readLoadProfilesExcel(filePath,day,graph,profilesType):
	# Open excel (buffered)
	if filePath != networkMaker.BUFFER_PROFILES_FILE:
		networkMaker.BUFFER_PROFILES = xlrd.open_workbook(filePath,on_demand=True)
		networkMaker.BUFFER_PROFILES_FILE = filePath
	xl = networkMaker.BUFFER_PROFILES

	# Constants
	# Row of the headers profile
	PROFILES_HEADERS_ROW={'R':5,'HP':8,'I1':8,'I2':9,'I3':9,'EC':8,'PV':0,'Wind':0,'IEP':1,'CHP':0}
	# Sign (+1 or -1) of the profile, positive for a production.
	PROFILES_SIGN={'R':-1, 'HP':-1,'I1':-1,'I2':-1,'I3':-1,'EC':-1,'PV':1,'Wind':1,'IEP':-1,'CHP':1}
	# Power factors of each profiles.
	PROFILES_POWER_FACTOR={'HP':0.91}
	# Default power factor.
	PROFILES_DEFAULT_POWER_FACTOR=0.95

	# Obtain the base profiles
	baseActiveProfiles={}
	baseReactiveProfiles={}

	# Single day not calendar dependent
	for pType in ['EC']:
		sheet=xl.sheet_by_name(pType)
		baseActiveProfiles[pType]=sheet.col_values(1, start_rowx=PROFILES_HEADERS_ROW[pType]+2, end_rowx=PROFILES_HEADERS_ROW[pType]+2+96)

	# Day dependent profiles
	for pType in ['PV','Wind','IEP']:
		sheet=xl.sheet_by_name(pType)
		baseActiveProfiles[pType]=sheet.col_values(2, start_rowx=PROFILES_HEADERS_ROW[pType]+1+(day-1)*96, end_rowx=PROFILES_HEADERS_ROW[pType]+1+day*96)

	# Obtain the base profile dependent on the calendar
	for pType in profilesType:
		sheet=xl.sheet_by_name(pType)

		# Get the profile column
		c=1
		while c < sheet.ncols:
			p=sheet.cell_value(PROFILES_HEADERS_ROW[pType],c).lower()
			if profilesType[pType]==p:
				break
			c+=1
		if p == '':
			raise Exception('Profile type "%s" not found for the load type "%s" in "%s"'%(profilesType[pType],pType,filePath))

		# Get the base profile
		if pType in ['R','HP']:
			baseActiveProfiles[pType]=sheet.col_values(c, start_rowx=PROFILES_HEADERS_ROW[pType]+2, end_rowx=PROFILES_HEADERS_ROW[pType]+2+96)
		elif pType in ['I1','I2','I3','CHP']:
			baseActiveProfiles[pType]=sheet.col_values(c, start_rowx=PROFILES_HEADERS_ROW[pType]+2, end_rowx=PROFILES_HEADERS_ROW[pType]+2+96)
			baseReactiveProfiles[pType]=sheet.col_values(c+1, start_rowx=PROFILES_HEADERS_ROW[pType]+2, end_rowx=PROFILES_HEADERS_ROW[pType]+2+96)
		else:
			raise Exception('Reading of the base profile of type %s not handled.'%pType)

	# Set the correct sign to each base profile
	for pType, profile in baseActiveProfiles.items():
		profile=list(map(lambda x: x*PROFILES_SIGN[pType],profile))
		baseActiveProfiles[pType]=profile

		if pType in baseReactiveProfiles:
			baseReactiveProfiles[pType]=list(map(lambda x: x*PROFILES_SIGN[pType],baseReactiveProfiles[pType]))
		else:
			if pType in PROFILES_POWER_FACTOR:
				baseReactiveProfiles[pType]=list(map(lambda x: x*math.tan(math.acos(PROFILES_POWER_FACTOR[pType])),profile))
			else:
				baseReactiveProfiles[pType]=list(map(lambda x: x*math.tan(math.acos(PROFILES_DEFAULT_POWER_FACTOR)),profile))

	# Iterate over the nodes
	for n,ndata in graph.nodes(data=True):
		# Filter nodes without data
		if len(ndata)==0 or ndata['load'] is None:
			continue
		load=ndata['load']
		for refType,refPower in load.refPowers.items():
			if refType == 'load':
				if load.loadType in ['R']:
					continue
				elif load.loadType in ['I1','I2','I3','IEP']:
					load.activeProfiles[refType]=list(map(lambda x: load.refPowers[refType]*x,baseActiveProfiles[load.loadType]))
					load.reactiveProfiles[refType]=list(map(lambda x: load.refPowers[refType]*x,baseReactiveProfiles[load.loadType]))
				else:
					raise Exception('Unhandled load type: "%s".'%load.loadType)
			elif refType == 'inhab':
				if not load.loadType in ['R']:
					raise Exception('Inhabitant not handled with load type "%s".'%load.loadType)
				load.activeProfiles['load']=list(map(lambda x: load.refPowers[refType]*x,baseActiveProfiles[load.loadType]))
				load.reactiveProfiles['load']=list(map(lambda x: load.refPowers[refType]*x,baseReactiveProfiles[load.loadType]))
			elif refType in baseActiveProfiles:
				load.activeProfiles[refType]=list(map(lambda x: load.refPowers[refType]*x,baseActiveProfiles[refType]))
				load.reactiveProfiles[refType]=list(map(lambda x: load.refPowers[refType]*x,baseReactiveProfiles[refType]))
			else:
				raise Exception('Production/consumption of type "%s" not handled.'%p)

## Read the calendar excel file and obtain the day type given a load type.
# @param filePath Path to the load profiles excel file.
# @param year Year.
# @param day Day of the year as a number.
# @return Dictionary of the profile type.
def readCalendarExcel(filePath,year,day):
	# Start cell of the month in the calendar excel.
	CALENDAR_CELLS={'January':(3,2),'February':(18,2),'March':(33,2),
					'April':(3,10),'May':(18,10),'June':(33,10),
					'July':(3,18),'August':(18,18),'September':(33,18),
					'October':(3,26),'November':(18,26),'December':(33,26)
					}
	# Width of a month in the calendar excel.
	CALENDAR_WIDTH=7
	# Height of a month in the calendar excel.
	CALENDAR_HEIGHT=12
	# Differences of ordinal day count from excel and python.
	ORDINAL_EXCEL_DIF=737425-43831
	# Day type abbreviation map column.
	CALENDAR_ABBRV_MAP_ROW=1
	# Day type abbreviation map column.
	CALENDAR_ABBRV_MAP_COL=34
	# Load type of the calendar.
	CALENDAR_LOAD_TYPES=['R','HP','I1','I2','I3','CHP']


	# Get the date corresponding to the day of the year
	date=datetime.date(year,1,1)
	date=datetime.date.fromordinal(date.toordinal() - 1 + day)
	month=date.strftime("%B")
	excelOrdinal=int(date.toordinal()-ORDINAL_EXCEL_DIF)

	# Open excel (buffered)
	if filePath != networkMaker.BUFFER_CALENDAR_FILE:
		networkMaker.BUFFER_CALENDAR = xlrd.open_workbook(filePath,on_demand=True)
		networkMaker.BUFFER_CALENDAR_FILE = filePath
	xl = networkMaker.BUFFER_CALENDAR

	# Read from it
	profilesType={}
	for loadType in CALENDAR_LOAD_TYPES:
		sheet=xl.sheet_by_name("%s %s"%(loadType,year))

		# Find the cell
		found=False
		l0,c0=CALENDAR_CELLS[month]
		l,c=l0-2,c0
		lEnd=l+CALENDAR_HEIGHT
		cEnd=c+CALENDAR_WIDTH
		while l < lEnd and not found:
			c=c0
			l+=2
			while c < cEnd:
				v=sheet.cell_value(l,c)
				if v != '' and int(v)==excelOrdinal:
					found=True
					break
				c+=1

		if not found:
			raise Exception('Day %s %s %s not found in calendar file for the load type %s.'%(date.day,month,year,loadType))

		# Get the day type with its abbreviation
		shortDayType=sheet.cell_value(l+1,c)

		# Get the day type
		fullDayType=None
		l=CALENDAR_ABBRV_MAP_ROW
		abbrv=sheet.cell_value(l,CALENDAR_ABBRV_MAP_COL)
		while abbrv != "":
			# Check
			if abbrv == shortDayType:
				fullDayType=sheet.cell_value(l,CALENDAR_ABBRV_MAP_COL+1)
				break

			# iterate
			l+=1
			abbrv=sheet.cell_value(l,CALENDAR_ABBRV_MAP_COL)

		if fullDayType is None:
			raise Exception('Day type %s not found in calendar file for the load type %s and the day %s %s %s.'%(shortDayType,loadType,date.day,month,year))

		profilesType[loadType]=fullDayType.lower()

	return profilesType

## Read the scenarios excel file.
# @param filePath Path to the scenarios excel file.
# @param year Year of the scenarios.
# @param graph Graph to add the loads.
# @param scenarioType Type of scenario. Usually 'L' or 'H'.
def readScenariosExcel(filePath,year,graph,scenarioType='H'):
	# Open excel (buffered)
	if filePath != networkMaker.BUFFER_SCENARIOS_FILE:
		networkMaker.BUFFER_SCENARIOS = xlrd.open_workbook(filePath,on_demand=True)
		networkMaker.BUFFER_SCENARIOS_FILE = filePath
	xl = networkMaker.BUFFER_SCENARIOS

	# Constants
	# Number of rows of the header of the scenario sheet.
	SCENARIOS_HEADER_SIZE=6
	# Number of the column of bus.
	SCENARIOS_BUS_COLUMN=0
	# Number of the column of the profile type.
	SCENARIOS_TYPE_COLUMN=2
	# Reference powers of the load.
	SCENARIOS_REFERENCE_POWERS={'load':3,'inhab':4,'EC':5,'HP':6,'PV':7,'CHP':8,'Wind':9}

	# Read the file
	loadCount=1
	unknownBuses=set()
	doubleLoads=set()
	sheet=xl.sheet_by_name("scenarios %s %s"%(year,scenarioType))
	for row in range(SCENARIOS_HEADER_SIZE,sheet.nrows):

		# Get bus and check existence
		bus=int(sheet.cell_value(row,SCENARIOS_BUS_COLUMN))
		if bus==0: # Dummy line
			continue
		elif not graph.has_node(bus):
			unknownBuses.add(bus)
			continue
		# Fetch informations
		loadType=sheet.cell_value(row,SCENARIOS_TYPE_COLUMN)

		# Add to the load list of the bus
		loadData=LoadData(loadCount,bus,loadType)
		try:
			if graph.node[bus]["load"] is not None:
				doubleLoads.add(bus)
			else:
				graph.node[bus]["load"]=loadData
				loadCount+=1
		except KeyError:
			pass

		# Add load information
		for t,c in SCENARIOS_REFERENCE_POWERS.items():
			v=float(sheet.cell_value(row,c))
			if v != 0:
				loadData.refPowers[t]=v

	if len(unknownBuses) > 0:
		print("%s unknown buses in the scenarios file \"%s\":\n\t%s"%(len(unknownBuses),filePath,unknownBuses))
	if len(doubleLoads) > 0:
		print("%s double loads in the scenarios file \"%s\" for the following buses:\n\t%s"%(len(doubleLoads),filePath,doubleLoads))

## Load of the network which may be production and/or consumption.
class LoadData:
	## Constructor.
	# @param internalId Internal id of the load.
	# @param bus Id of the bus the load is attached to.
	# @param loadType Type of the load.
	def __init__(self,internalId,bus,loadType):
		self.internalId=internalId
		self.bus=bus
		self.loadType=loadType
		self.refPowers={}
		self.activeProfiles={}
		self.reactiveProfiles={}

	## @var internalId
	# Internal id of the load.
	## @var bus
	# Id of the bus the load is attached to.
	## @var loadType
	# Type of the load.
	## @var refPowers
	# Dictionary with the reference powers.
	## @var activeProfiles
	# Dictionary with the active profiles. Productions are positive.
	## @var reactiveProfiles
	# Dictionary with the reactive profiles.
