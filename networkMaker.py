## Classes and methods needed to convert the series of Excel files with the characteristics of the network to a CSV file.
# Requires xlrd and networkx which can be installed with "pip3 install xlrd" and "pip install networkx".
#@author Sebastien MATHIEU

import xlrd
import networkx

# Constants
## File with the information on the buses.
BUSES_FILE="Noeuds.xlsx"
## File with the information about the cable type.
CABLES_FILE="cables.xlsx"
## File with the information about the connection between the buses.
LINES_FILE="feedersMT.xlsx"
## File with the information about the MV-LV transformers.
LV_TRANSFORMERS_FILE="transfo MTBT.xlsx"

## Read the buses excel file.
# @param filePath Path to the buses excel file.
# @param graph Networkx graph to which the buses should be added as nodes.
def readBusesExcel(filePath,graph):
	# Constants of the buses excel files
	# Number of the column of the id of the bus in the bus file.
	BUS_ID_COLUMN=0
	# Number of the column of the base voltage of the bus in the bus file.
	BUS_VOLTAGE_COLUMN=4
	# Number of the column of the bus bar
	BUS_BAR_COLUMN=3
	# Number of the column of the cell
	CELL_COLUMN=2
	# Number of the column of the cell status, which indicates if the connected line is open or closed on this side
	CELL_STATUS_COLUMN=7


	# Read the file
	xl = xlrd.open_workbook(filePath,on_demand=True)
	busCount=0
	for sheet in xl.sheets():
		for row in range(1,sheet.nrows):
			id=excelStr2int(sheet.cell_value(row, BUS_ID_COLUMN))
			v=sheet.cell_value(row, BUS_VOLTAGE_COLUMN)
			connection_info = (excelStr2int(sheet.cell_value(row, BUS_BAR_COLUMN)), excelStr2int(sheet.cell_value(row, CELL_COLUMN)))
			connection_closed = True if sheet.cell_value(row, CELL_STATUS_COLUMN) == 'F' else False
			if graph.has_node(id):
				graph.node[id]["count"]+=1
				graph.node[id]["connections"][connection_info] = connection_closed
			else:
				connections = {connection_info:connection_closed}
				graph.add_node(id,{"internalId":busCount,"id":id,"baseVoltage":v,"count":1,"transformers":[],"load":None, "connections":connections})
				busCount+=1


## Convert a string from excel to an int.
# @param s String.
# @return Integer.
def excelStr2int(s):
	v=s
	if type(s) is type("") and s[0]=="'":
		v=int(s[1:])
	else:
		v=int(s)
	return v

## Make the network structure from the parameters given in excel files.
# @param folderPath Folder with the excel files.
# @return Graph.
def makeNetwork(folderPath):
	graph=networkx.MultiGraph()

	readBusesExcel(folderPath+BUSES_FILE,graph)
	cables=readCablesExcel(folderPath+CABLES_FILE)
	readLinesExcel(folderPath+LINES_FILE,cables,graph)
	readLvTransformersExcel(folderPath+LV_TRANSFORMERS_FILE,graph)

	#TODO: check graph is radial when accounting for open and closed lines.
	#TODO: check sum of length of segments equals encoded line length

	return graph

## Convert the multigraph of the network to a graph.
# @param mg Multigraph.
# @return Graph.
def multiGraphToGraph(mg):
	g=networkx.Graph()

	for n,ndata in mg.nodes(data=True):
		g.add_node(n,ndata)

	lineCount=1
	for u,v,edata in mg.edges(data=True):
		if g.has_edge(u,v):
			# Replace the line if needed
			if g.get_edge_data(u,v)["pMax"] < edata["pMax"]:
				edata["number"]=g.get_edge_data(u,v)["number"]
				g.remove_edge(u,v)
				g.add_edge(u,v,edata)
		else:
			edata["number"]=lineCount
			lineCount+=1
			g.add_edge(u,v,edata)

	return g

## Read the cables excel file content.
# @param filePath Path to the cables excel file.
# @return Map of cables with their type as a key.
def readCablesExcel(filePath):
	# Constants of the cables excel files
	# Number of the column of the type of the cable.
	CABLE_TYPE_COLUMN=0
	# Number of the column of the maximum current in the cable.
	CABLE_R1_COLUMN=1
	# Number of the column of the maximum current in the cable.
	CABLE_X1_COLUMN=2
	# Number of the column of the maximum current in the cable.
	CABLE_C1_COLUMN=3

	# Number of the column of the maximum current in the cable.
	CABLE_IMAX_COLUMN=9

	# Open and parse the cable sheet
	sheet=xlrd.open_workbook(filePath,on_demand=True).sheet_by_index(0)
	cables={}
	for row in range(1,sheet.nrows):
		cableType=sheet.cell_value(row,CABLE_TYPE_COLUMN)
		R1=sheet.cell_value(row,CABLE_R1_COLUMN)
		X1=sheet.cell_value(row,CABLE_X1_COLUMN)
		C1=sheet.cell_value(row,CABLE_C1_COLUMN)
		iMax=sheet.cell_value(row,CABLE_IMAX_COLUMN)
		cables[cableType]=CableData(cableType,R1,X1,C1,iMax)

	return cables

## Class containing the data of a bus.
class CableData:
	## Constructor.
	# @param cableType Type of the cable.
	# @param R1 Direct conductance.
	# @param X1 Direct reactance.
	# @param C1 Direct capacitance.
	# @param iMax Maximum current in the cable.
	def __init__(self,cableType,R1=0,X1=0,C1=0,iMax=None):
		self.cableType=cableType
		self.R1=R1
		self.X1=X1
		self.C1=C1
		self.iMax=float(iMax)

	## @var cableType
	# Type of the cable.
	## @var iMax
	# Maximum current in the cable.
	## @var R1
	# Direct conductance.
	## @var X1
	# Direct reactance.
	## @var C1
	# Direct capacitance.

## Get the cable associated to a line.
# @param cables Cables data as a map.
# @param row Row of the line.
# @param sheet Excel sheet of the line file.
def getCable(cables, row, sheet):
	# Cable potential prefixes of their id.
	CABLE_ID_PREFIXES=['S','A','Câble']
	# Cable characteristics
	# Number of the column of the section.
	LINE_SECTION_COLUMN=19
	# Number of the column of the insulation.
	LINE_INSULATION_COLUMN=17
	# Number of the column of the insulation voltage.
	LINE_INSULATION_VOLTAGE_COLUMN=18
	# Number of the column of the core type.
	LINE_CORE_COLUMN=16

	section = sheet.cell_value(row, LINE_SECTION_COLUMN)
	insulation = sheet.cell_value(row, LINE_INSULATION_COLUMN)
	insulationVoltage = sheet.cell_value(row, LINE_INSULATION_VOLTAGE_COLUMN)
	core = sheet.cell_value(row, LINE_CORE_COLUMN)
	# Match the corresponding cable
	cable = None
	for prefix in CABLE_ID_PREFIXES:
		# Try one prefix
		try:
			cable = cables['%s-%s-%s-%s-%s' % (prefix, section, core, insulation, insulationVoltage)]
			break
		except KeyError:
			pass
	if cable is None:
		raise Exception(
			"Cable not found with the following characteristics:\n section:%s, core:%s, insulation:%s, insulation voltage: %s" % (section, core, insulation, insulationVoltage))

	return cable


## Read the lines excel file.
# @param filePath Path to the excel file with the information on the lines.
# @param cables Cables data as a map.
# @param graph Graph to add edges to.
def readLinesExcel(filePath,cables,graph):
	# Constants of the buses excel files
	# Number of the column of the id of the line.
	LINE_ID_COLUMN=0
	# Number of the column of the id of the "from" bus.
	LINE_FROM_COLUMN=1
	# Number of the column of the id of the "to" bus.
	LINE_TO_COLUMN=4
	# Number of the column of the bus bar on the "to" bus
	FROM_BUSBAR_COLUMN=2
	# Number of the column of the cell on the bus bar on the "from" bus
	FROM_CELL_COLUMN=3
	# Number of the column of the bus bar on the "to" bus
	TO_BUSBAR_COLUMN=5
	# Number of the column of the cell on the bus bar on the "to" bus
	TO_CELL_COLUMN=6
	# Number of the column of the length.
	LINE_LENGTH_COLUMN=20
	# Number of the column of the line voltage.
	LINE_VOLTAGE_COLUMN=13

	# Open and parse the line sheet
	lineCount=1
	sheet=xlrd.open_workbook(filePath,on_demand=True).sheet_by_index(0)
	unknownBuses=set()
	openLines=set()
	for row in range(1,sheet.nrows):
		# Line information
		id=excelStr2int(sheet.cell_value(row, LINE_ID_COLUMN))
		fromBus=excelStr2int(sheet.cell_value(row, LINE_FROM_COLUMN))
		toBus=excelStr2int(sheet.cell_value(row, LINE_TO_COLUMN))

		if fromBus == toBus:
			continue

		cable = getCable(cables, row, sheet)
		if graph.get_edge_data(fromBus,toBus,id) is None:
			# This is the first (and maybe only) segment of the link between these buses
			fromBusBar = excelStr2int(sheet.cell_value(row, FROM_BUSBAR_COLUMN))
			fromCell = excelStr2int(sheet.cell_value(row, FROM_CELL_COLUMN))
			toBusBar = excelStr2int(sheet.cell_value(row, TO_BUSBAR_COLUMN))
			toCell = excelStr2int(sheet.cell_value(row, TO_CELL_COLUMN))

			# Check source and destination existence
			previousUnknown=len(unknownBuses)
			if not graph.has_node(fromBus):
				unknownBuses.add(fromBus)
			if not graph.has_node(toBus):
				unknownBuses.add(toBus)
			if previousUnknown < len(unknownBuses):
				continue

			# Check line is closed.
			closed = False
			try:
				# Line is closed if both end cells are closed.
				closed = graph.node[fromBus]["connections"][(fromBusBar,fromCell)] and graph.node[toBus]["connections"][(toBusBar,toCell)]
			except KeyError:
				pass #Because fromBus or toBus is not in the file describing nodes (to be fixed there)

			if not closed:
				openLines.add(id)

			# Get the length and compute the characteristics
			lineAttr={}
			lineAttr["id"]=id
			lineAttr["length"]=float(sheet.cell_value(row, LINE_LENGTH_COLUMN))*1e-3 # in km
			voltage=float(sheet.cell_value(row, LINE_VOLTAGE_COLUMN))
			lineAttr["R1"]=lineAttr["length"]*cable.R1 # in Ohm
			lineAttr["X1"]=lineAttr["length"]*cable.X1 # in Ohm
			lineAttr["C1"]=lineAttr["length"]*cable.C1 # in microFarad
			lineAttr["pMax"]=voltage*cable.iMax # in VA
			lineAttr["internalId"]=lineCount
			lineAttr["closed"]=closed

			lineCount += 1

			# Add the line to the list
			graph.add_edge(fromBus,toBus,id,lineAttr)
		else:
			edgeData = graph.get_edge_data(fromBus,toBus,id)
			# We have several segments, compute the overall impedance!!!
			R1 = edgeData["R1"] # in Ohm
			X1 = edgeData["X1"] # in Ohm
			C1 = edgeData["C1"] # in microFarad

			from math import pi
			baseFrequency = 50
			omega = 2*pi*baseFrequency

			# First pi-model
			Z1 = R1 + 1j*X1
			B1 = omega*C1*1e-6
			#Second pi-model
			length = float(sheet.cell_value(row, LINE_LENGTH_COLUMN))*1e-3
			Z2 = length*cable.R1 + 1j *length*cable.X1
			B2 = omega*length*cable.C1 * 1e-6

			# Residual shunt admittance "in the middle"
			Y3 = 1j*(B1+B2)/2

			if abs(Y3) > 0:
				# Start to triangle
				Z3 = 1/Y3
				YY1 = Z3/(Z1*Z2+Z1*Z3+Z2*Z3)
				YY2 = Z2/(Z1*Z2+Z1*Z3+Z2*Z3)
				YY3 = Z1/(Z1*Z2+Z1*Z3+Z2*Z3)

				Y12 = YY2+1j*B1/2
				Y13 = YY3+1j*B2/2

				Yaverage23 = Y12+Y13/2
				edgeData["R1"] = YY1.real # in Ohm
				edgeData["X1"] = YY1.imag # in Ohm
				edgeData["C1"] = Yaverage23.imag/omega * 1e6 # in microFarad
			else:
				# Just sum up the two line impedances
				Z = Z1+Z2
				edgeData["R1"] = Z.real # in Ohm
				edgeData["X1"] = Z.imag # in Ohm
				edgeData["C1"] = 0 # in microFarad

			# Maximum capacity is the minimum of the capacity of the two segments
			voltage=float(sheet.cell_value(row, LINE_VOLTAGE_COLUMN))
			edgeData["pMax"] = min(edgeData["pMax"], voltage*cable.iMax)

	if len(unknownBuses) > 0:
		print("%s unknown buses in the lines file \"%s\":\n\t%s"%(len(unknownBuses),filePath,sorted(unknownBuses)))
	if len(openLines) > 0:
		print("%s open lines in the lines file \"%s\":\n\t%s"%(len(openLines),filePath,sorted(openLines)))

	# Check non-connected buses
	aloneBuses=set(filter(lambda n: len(graph.neighbors(n)) == 0, graph.nodes()))
	if len(aloneBuses) > 0:
		print("%s buses alone according to the lines file \"%s\":\n\t%s"%(len(aloneBuses),filePath,sorted(aloneBuses)))


## Read the lines excel file with the MV-LV transformers.
# @param filePath Path to the excel file with the information on the transformers.
# @param graph Graph to add edges to.
def readLvTransformersExcel(filePath,graph):
	# Constants of the transformers excel files
	# Number of the column of the bus the transformer is attached to.
	TRANSFORMER_BUS_COLUMN=0
	# Number of the column of the id.
	TRANSFORMER_ID_COLUMN=1
	# Number of the column of the maximum power.
	TRANSFORMER_PMAX_COLUMN=2

	# Open and parse the transformer sheet
	transformerCount=1
	sheet=xlrd.open_workbook(filePath,on_demand=True).sheet_by_index(0)
	unknownBuses=set()
	for row in range(1,sheet.nrows):
		# Get the bus and check existence
		bus=int(sheet.cell_value(row, TRANSFORMER_BUS_COLUMN))
		if not graph.has_node(bus):
			unknownBuses.add(bus)
			continue

		# Fetch information
		id=int(sheet.cell_value(row, TRANSFORMER_ID_COLUMN))
		pmax=sheet.cell_value(row, TRANSFORMER_PMAX_COLUMN)

		# Add to the transformer list of the bus
		try:
			graph.node[bus]["transformers"].append(TransformerData(transformerCount,id,bus,pmax))
			transformerCount+=1
		except KeyError:
			pass

	if len(unknownBuses) > 0:
		print("%s unknown buses in the transformers file \"%s\":\n\t%s"%(len(unknownBuses),filePath,unknownBuses))

## Data of a transformer.
class TransformerData:
	## Constructor
	# @param internalId Internal id of the transformer.
	# @param id Id of the transformer.
	# @param bus Bus to which the transformer is attached.
	# @param pmax Maximal power of the transformer.
	def __init__(self,internalId,id=None,bus=None,pmax=None):
		self.internalId=internalId
		self.id=id
		self.pmax=pmax
		self.bus=bus

	def __str__(self):
		if self.id is not None:
			return "Transformer%s"%self.id
		elif self.number is not None:
			return "Transformer%s"%self.number
		else:
			return "Transformer"

	## @var number
	# Number of the transformer.
	## @var id
	# Id of the transformer.
	## @var bus
	# Bus to which the transformer is attached.
	## @var pmax
	# Maximal power of the transformer.

## Detect the buses at the end of the network without loads.
def detectEndBuses(g):
	endBuses=set(filter(lambda n: len(g.neighbors(n)) == 1 and ('load' not in g.node[n] or g.node[n]['load'] is None), g.nodes()))
	print("End buses:\n\t%s"%endBuses)
