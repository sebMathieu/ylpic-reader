## Convert the series of Excel files with the characteristics of the network to a DSIMA instance.
# Requires xlrd which can be installed with "pip install xlrd".
#@author Sebastien MATHIEU

import sys,os,shutil,datetime,random, copy
import xlrd

import networkMaker
import scenariosReader
from dotConverter import makeNetworkDot

## Default numerical tolerance
EPS=0.001
## Number of periods
T=96
## Periods set
periods=range(1,T+1)
## Maximal powers of heat pumps (HP) and electric cars (EC) in MW.
maxPowers={'HP':3.0976390423/1000,'EC':1.25/1000}
## Relative upward flexibility by load profile. For the TSO, flex in MW.
flexU={'I1':0,'I2':0.75,'I3':0.75,'EC':0.8,'HP':0.1,'R':0, 'TSO': 5}
## Relative downward flexibility by load profile. For the TSO, flex in MW.
flexD={'I1':0.2,'I2':0,'I3':0,'EC':0.8,'HP':0.25,'R':0, 'TSO': 5}

## Entry point of the program.
# @param argv Program parameters.
def main(argv):
	# Parse arguments
	if len(argv) < 4 :
		displayHelp()
		sys.exit(2)
	inputPath=argv[0] if argv[0].endswith(("/","\\")) else argv[0]+"/"
	if not os.path.exists(inputPath):
		raise Exception('Folder \"%s\" does not exists.' % inputPath)
	outputPath=argv[1] if argv[1].endswith(("/","\\")) else argv[1]+"/"
	year=int(argv[2])
	scenario=argv[3]

	# Static mode?
	static=len(argv) == 5 and argv[4] == "--static"
	if static:
		print("STATIC MODE")
		for k in flexU.keys():
			flexU[k]=0
		for k in flexD.keys():
			flexD[k]=0

	#days=range(1,366)
	days={26,35,51,65,106,138,142,175,263,305,344,360}

	# Create output folder
	if not os.path.exists(outputPath):
		os.makedirs(outputPath)
	for d in days:
		dayDirectory='%s/%s'%(outputPath,d)
		if not os.path.exists(dayDirectory):
			os.makedirs(dayDirectory)

	# Read graph
	graph=networkMaker.makeNetwork(inputPath)
	addRootBus(graph,8001)
	makeNetworkCSV(outputPath,days,graph)

	# Read prices data
	pricesData = readPricesData('%sprices.xlsx' % inputPath)

	# Read day by day
	dotGenerated=False
	for d in days:
		print("\t%s" % d)
		g = copy.deepcopy(graph)
		scenariosReader.readScenarios(inputPath,year,d,g,scenario)

		if not dotGenerated:
			makeNetworkDot(g)
			dotGenerated=True

		makeProducers("%s/%s/producers"%(outputPath,d), g)
		makeRetailers("%s/%s/retailers"%(outputPath,d), g)
		makeQualificationIndicators("%s/%s"%(outputPath,d),g)
		makeTSO("%s/%s"%(outputPath,d))
		makePrices("%s/%s/prices.csv" % (outputPath,d), pricesData, year, d)

## Display help of the program.
def displayHelp():
	text="Usage :\n\tpython3 dsimaConverter.py dataFolder outputFolder year scenario [--static]\n"
	text+="\nExample:\n\tpython3 dsimaConverter.py ylpic 2020H 2020 H\n"
	print(text)

## Add a root bus to the graph.
# @param graph Graph with the data.
# @param connectionId Original id of the bus to connect the root to.
def addRootBus(graph, connectionId):
	# Create the node
	graph.add_node(0,{"internalId":0,"id":0,"baseVoltage":graph.node[connectionId]["baseVoltage"],"transformers":[],"load":None})

	# Compute capacity of edges connected to the connection bus.
	C=0
	for n in graph.neighbors_iter(connectionId):
		for e, edata in graph[connectionId][n].items():
			C+=edata['pMax']

	# Add the line to the list
	graph.add_edge(0,connectionId,0,{"id":0,"length":1,"R1":0.0001,"X1":0.0001,"C1":0.0001,"pMax":C,"internalId":0,"closed":True})

## Make a CSV file for the TSO with its parameters.
# @param outputPath Output path of the retailers files.
def makeTSO(outputPath):
	with open('%s/tso.csv'%outputPath, 'w') as file:
		file.write('# T, pi^S+, pi^S-\n')
		file.write('%s,%s,%s\n'%(T,45.0,-45.0))

		file.write('# t, R+, R-, E\n')
		for t in range(T):
			upwardFlexNeeds=flexU['TSO']
			downwardFlexNeeds=-flexD['TSO']

			# Generate and imbalance of the same sign than the annual data
			externalImbalance=random.uniform(downwardFlexNeeds,upwardFlexNeeds)
			file.write('%s,%s,%s,%s\n'%(t+1,upwardFlexNeeds,downwardFlexNeeds,externalImbalance))

## Make the CSV files with the flexibility qualification indicator.
# @param outputPath Output path of the retailers files.
# @param graph Graph with the data.
def makeQualificationIndicators(outputPath,graph):
	with open('%s/qualified-flex.csv'%outputPath, 'w') as file:
		file.write('# N\n%s\n'%len(graph.node))
		file.write('# n, d+, d-\n')
		for n,ndata in graph.nodes(data=True):
			if graph.node[n]['internalId']==0:
				file.write('%s,%s,%s\n'%(n,EPS,EPS))
			else:
				# Compute the flex
				maxFlexU=0
				maxFlexD=0

				try:
					loadData=ndata['load']
					if loadData is not None:
						for p, v in loadData.refPowers.items():
							maxFlexU+=v*flexU[p]
							maxFlexD+=v*flexD[p]
				except KeyError:
					pass

				# Write
				dU=EPS+1.0/(maxFlexU/10.0+maxFlexD/10.0+EPS)
				dD=EPS+1.0/(maxFlexU/1.0+maxFlexD/10.0+EPS)
				file.write('%s,%s,%s\n'%(graph.node[n]['internalId'],dU,dD))

## Make the retailers.
# @param outputPath Output path of the retailers files.
# @param graph Graph with the data.
def makeRetailers(outputPath, graph):
	# Create the output if it doesn't exist
	if not os.path.exists(outputPath):
		os.makedirs(outputPath)

	# Periods
	T = 96
	periods = range(1,T+1)

	# Make a retailer for each consumption type
	for c in ['R','I1','I2','I3','EC','HP']:
		# Get profile name
		p=c
		if p in ['R','I1','I2','I3']:
			p='load'

		# Obtain the list of nodes
		pNodes=set()
		for n,ndata in graph.nodes(data=True):
			# Obtain the load data
			loadData=None
			try:
				loadData=ndata['load']
			except KeyError:
				pass
			if loadData is None:
				continue

			# Check profile in the node
			if p not in loadData.activeProfiles.keys() or (p == 'load' and loadData.loadType != c):
				continue

			# Add to the list
			pNodes.add(n)

		if len(pNodes) == 0:
			continue

		with open('%s/%s.csv'%(outputPath,c), 'w') as file:
			file.write('# T, pi^r, pi^f\n%s, 5, 60\n'%T) #TODO other 2values ?
			file.write('# N set\n%s\n'%', '.join(map(str,map(lambda x: graph.node[x]['internalId'],pNodes))))

			file.write('# n, t, p^min, p^max\n')
			g={}
			for n in pNodes:
				g[n]=0
				load = graph.node[n]['load']
				for t in periods:
					d = load.activeProfiles[p][t-1]/1000
					pMin = d*(1.0+flexD[c])
					if p in ['HP','EC']:
						pMin = max(pMin, -maxPowers[p]*load.refPowers[p])
					pMax = min(0,d*(1.0-flexU[c]))
					if pMin > pMax:
						print('Warning: Min > max (%s > %s) with base power %s and profile %s and a reference %s.'%(pMin,pMax,d,c, load.refPowers[p]))
						pMin = pMax
					g[n] = min(g[n], pMin)
					file.write('%s,%s,%s,%s\n'%(graph.node[n]['internalId'], t, pMin, pMax))

			# Write maximal access requirement
			file.write('# n, V, g, G\n')
			for n in pNodes:
				load = graph.node[n]['load']
				V = sum(load.activeProfiles[p])/4000 # Divided by 4 coz quarters
				if p != 'load':
					g[n] = min(g[n], -maxPowers[p]*load.refPowers[p] if p in ['HP','EC'] else -graph.node[n]['load'].refPowers[p]/1000)
				file.write('%s,%s,%s,0\n'%(graph.node[n]['internalId'], V, g[n]))

			# Write external imbalance
			file.write('# t, E\n')
			for t in periods:
				# Assume 0 external imbalance
				file.write('%s,%s\n'%(t,0))


## Make the producers.
# @param outputPath Output path of the producers files.
# @param graph Graph with the data.
def makeProducers(outputPath,graph):
	# Create the output if it doesn't exist
	if not os.path.exists(outputPath):
		os.makedirs(outputPath)

	# Periods
	T = 96
	periods = range(1,T+1)

	# Make a producer for each production type
	profiles=['Wind','PV','CHP']
	marginalCosts={'Wind':-65.0,'PV':-65.0,'CHP':60.0}
	for p in profiles:
		# Obtain the list of nodes
		pNodes=set()
		for n,ndata in graph.nodes(data=True):
			# Obtain the load data
			loadData=None
			try:
				loadData=ndata['load']
			except KeyError:
				pass
			if loadData is None:
				continue

			# Check profile in the node
			if p not in loadData.refPowers.keys():
				continue

			# Add to the list
			pNodes.add(n)

		if len(pNodes) == 0:
			continue

		# Write the producer
		with open('%s/%s.csv'%(outputPath,p), 'w') as file:
			file.write('# T\n%s\n'%T)
			file.write('# N set\n%s\n'%', '.join(map(str,map(lambda x: graph.node[x]['internalId'],pNodes))))

			file.write('# n, t, p^min, p^max, c, pi^r\n')
			for n in pNodes:
				for t in periods:
					file.write('%s,%s,%s,%s,%s,0.001\n'%(graph.node[n]['internalId'],t,0,graph.node[n]['load'].activeProfiles[p][t-1]/1000,marginalCosts[p]))

			# Write external imbalance
			file.write('# t, E\n')
			for t in periods:
				# Assume 0 external imbalance
				file.write('%s,%s\n'%(t,0))

			# Write maximal access requirement
			file.write('# n, g, G\n')
			for n in pNodes:
				file.write('%s,0,%s\n'%(graph.node[n]['internalId'],graph.node[n]['load'].refPowers[p]/1000))


## Make the prices csv file.
# @param outputPath Output path of the prices csv file.
# @param pricesData Prices data.
# @param year Year.
# @param day Day.
def makePrices(outputPath, pricesData, year, day):
	with open(outputPath, 'w') as file:
		file.write('# T, EPS, pi^l, dt\n')
		file.write('%s, %s, %s, %s\n'%(T, EPS, 100.0, 0.25)) # Number of periods, accuracy, local imbalance penalty [\euro/MWh], period size [h]

		file.write('# t, pi^E, pi^I+, pi^I-\n')
		for t in range(T):
			piE=pricesData[day]['energy price'][t]
			minImbalancePrice=max([EPS,piE+EPS,-piE+EPS])

			piIU=pricesData[day]['upward imbalance price'][t]
			piID=pricesData[day]['downward imbalance price'][t]

			if piIU < minImbalancePrice:
				piIU=minImbalancePrice
			if piID < minImbalancePrice:
				piID=minImbalancePrice
			file.write('%s,%s,%s,%s\n'%(t+1,piE,piIU,piID))

## Create the CSV file with the network.
# @param outputPath Folder with the instances to output
# @param days Days to consider.
# @param graph Graph.
def makeNetworkCSV(outputPath,days,graph):
	unknownBuses=[]
	Sb=100 # Base power in MVA
	Vb=10 # Base voltage in kV
	Yb=Sb/(Vb*Vb) # Base admittance

	# Create the tmp file
	tmpFile="network.csv"
	with open(tmpFile, 'w') as file:
		file.write("# File describing the network topology.\n")
		file.write("# Buses, Links, pi^VSP, pi^VSC, Sb [MVA], Vb [kV], default Q/P\n")
		file.write("%s,%s,500,1000, %s, %s, 0.1\n"%(len(graph.nodes()),len(graph.edges()), Sb, Vb))
		file.write("# Link id, from bus, to bus, Conductance [pu.], Susceptance [pu.], C [MVA], original id\n")

		for u,v,edata in graph.edges(data=True):
			if not edata["closed"]:
				pass
			fromBus=graph.node[u]['internalId']
			toBus=graph.node[v]['internalId']
			z=complex(edata['R1'], edata['X1'])
			y=(1/z)/Yb
			file.write("%s,%s,%s,%s,%s,%s,%s\n"%(edata['internalId']+1,fromBus,toBus,y.real,y.imag,edata['pMax']/1000000,edata['id']))


		file.write("# Bus id, Base voltage [kV], Vmin [pu.], Vmax [pu.], original id\n")
		for n,ndata in graph.nodes(data=True):
			file.write("%s,%s,0.95,1.05,%s\n"%(ndata['internalId'],ndata['baseVoltage']/1000,ndata['id']))

	if len(unknownBuses) > 0:
		raise Exception("%s unknown buses:\n\t%s"%(len(unknownBuses),unknownBuses))

	# Copy to the daily folders
	for d in days:
		oPath='%s/%s/network.csv'%(outputPath,d)
		shutil.copyfile(tmpFile,oPath)

	# Clean
	os.remove(tmpFile)

## Read prices excel file.
# @param filePath Path to the excel file.
# @return Time series of the prices file in a dictionary which keys are days number.
def readPricesData(filePath):
	# Constants of the transformers excel files
	# Number of header lines.
	HEADER_LINES=2
	# Number of the column with the energy prices
	ENERGY_PRICES_COLUMN=2
	# Number of the column with the upward imbalance prices
	UPWARD_IMBALANCE_PRICES_COLUMN=3
	# Number of the column with the downward imbalance prices
	DOWNWARD_IMBALANCE_PRICES_COLUMN=4
	# Number of the column with the system imbalance
	SYSTEM_IMBALANCE_COLUMN=5
	# Prices labels
	PRICES_LABELS={"energy price":ENERGY_PRICES_COLUMN, "upward imbalance price":UPWARD_IMBALANCE_PRICES_COLUMN, "downward imbalance price":DOWNWARD_IMBALANCE_PRICES_COLUMN,"system imbalance":SYSTEM_IMBALANCE_COLUMN}

	# Open the file
	workbook=xlrd.open_workbook(filePath,on_demand=True)
	sheet=workbook.sheet_by_index(0)

	# Get the initial date
	initialDayTuple = xlrd.xldate_as_tuple(sheet.cell_value(HEADER_LINES,0),workbook.datemode)
	initialDay = datetime.datetime(initialDayTuple[0], 1, 1, 0, 0, 0).toordinal()

	# Parse content
	dayRaw = None
	dayTimeSeries = {}
	timeSeries = {}
	for r in range(HEADER_LINES,sheet.nrows):
		row = sheet.row_values(r, 0, sheet.ncols)

		# New day?
		if row[0] != dayRaw:
			# Check the previous one is full
			if dayRaw is not None and len(dayTimeSeries['energy price']) != 96:
				dayTuple = xlrd.xldate_as_tuple(dayRaw,workbook.datemode) # Gregorian (year, month, day, hour, minute, nearest_second)
				d = datetime.datetime(dayTuple[0], dayTuple[1], dayTuple[2], dayTuple[3], dayTuple[4], dayTuple[5])
				raise Exception("Laking prices for day %s, only %s entries instead of 96." % (d, len(dayTimeSeries['energy price'])))

			# New day, cast it as a day and convert it to day number
			dayRaw = row[0]
			dayTuple = xlrd.xldate_as_tuple(row[0],workbook.datemode) # Gregorian (year, month, day, hour, minute, nearest_second)
			day = datetime.datetime(dayTuple[0], dayTuple[1], dayTuple[2], dayTuple[3], dayTuple[4], dayTuple[5])
			day = (day.toordinal()-initialDay)+1

			# Initialize the time series
			dayTimeSeries = {}
			for p in PRICES_LABELS.keys():
				dayTimeSeries[p] = []
			timeSeries[day] = dayTimeSeries

		# Add the data
		for p,col in PRICES_LABELS.items():
			v = float(row[col])
			dayTimeSeries[p].append(v)

	return timeSeries

# Starting point from python #
if __name__ == "__main__":
	main(sys.argv[1:])