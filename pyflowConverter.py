## Convert the series of Excel files with the characteristics of the network to a pyflow file.
# Requires xlrd which can be installed with "pip3 install xlrd".
#@author Bertrand CORNELUSSE

# Last update 2015-09-29
#TODO Clean input reading
#TODO Generators
#TODO Transformers
#TODO Shunt admittances at buses

import sys,os
import networkx
import math

import networkMaker
import scenariosReader

## Entry point of the program.
# @param argv Program parameters: data folder path, slack bus id, period of interest (scenario hardcoded for now).
def main(argv):
	# Read instance folder
	if len(argv) < 2 :
		displayHelp()
		sys.exit(2)
	folderPath=argv[0] if argv[0].endswith(("/","\\")) else argv[0]+"/"

	# Provide slack bus external ID as second argument
	slackBusId = int(argv[1])

	# Provide period of the scenario as third argument
	if not 0<period<96:
		raise Exception('Invalid period %s'%period)

	if not os.path.exists(folderPath):
		raise Exception('Folder \"%s\" does not exists.' % folderPath)

	# Read network data (all but power information)
	graph=networkMaker.makeNetwork(folderPath)

	# Read power information, select time horizon, day number, and scenario type.
	scenariosReader.readScenarios(folderPath,2020,1,graph,'H')

	# Output
	makePyflowCSV('caseYlpic.py', graph, 'ylpic', slackBusId, period)

## Display help of the program.
def displayHelp():
	text="Usage :\n\tpython pyflowConverter.py dataFolder slackBusId period\n"
	print(text)

## Convenience function
def writeLine(outFile,str2Write):
	outFile.write("""    %s\n""" % str2Write)
	return

def makePyflowCSV(fileName, networkGraph, caseName, slackBusId, period):

	# Constants for conversion to per unit
	baseMVA = 100 # MVA
	baseKV = -1.0 # Dummy value to start with, will be read from data
	baseFrequency = 50 # Hz
	# Constants for operational limits
	VLimitDown = 0.95
	VLimitUp = 1.05

	outFile = open(fileName,'w')

	# write front matter
	outFile.write("""from numpy import array\n\ndef %s():\n\n""" % caseName)
	writeLine(outFile,"""ppc = {"version": '2'}""")
	writeLine(outFile,"""## system MVA base""")
	writeLine(outFile,"""ppc["baseMVA"] = %f""" % baseMVA)

	# Write buses in the format ["bus_i", "type", "Pd", "Qd", "Gs", "Bs", "area", "Vm", "Va", "baseKV", "zone", "Vmax", "Vmin"] (cf. pypower data format)
	# First, filtering the data and ordering it according to the internalId
	busData = {}
	for n,ndata in networkGraph.nodes(data=True):
		# Type, 1: load, 2: generator, 3: slack bus
		if n == slackBusId:
			type = 3
		else:
			#TODO Handle generators, for now assuming everything is a load
			type = 1

		# Injection data
		Pd = 0 # MW
		Qd = 0 # MVar
		try:
			loadData = ndata['load']
			if loadData is not None:
				for baseline in loadData.activeProfiles.values():
					Pd += baseline[period]/1e3 # Convert from W to MW
				for baseline in loadData.reactiveProfiles.values():
					Qd += baseline[period]/1e3 # Convert from VAr to MVAr
		except KeyError:
			pass

		Pd = round(Pd,6) # Round for writing to file
		Qd = round(Qd,6) # Round for writing to file

		# Shunt admittance
		Gs = 0 #TODO if any need one day
		Bs = 0 #TODO if any need one day

		# Base voltage
		if baseKV == -1:
			baseKV = ndata['baseVoltage']/1e3
		elif baseKV != ndata['baseVoltage']/1e3:
			print("Warning: several voltage levels: %f" % (ndata['baseVoltage']/1000))

		busData[ndata['internalId']] = [n, type, Pd, Qd, Gs, Bs, 1, 1, 0,baseKV,1,VLimitUp,VLimitDown]

	# Finally printing to the file by increasing internal id.
	writeLine(outFile,"""## Bus data""")
	writeLine(outFile,"""ppc["bus"] = array([""")
	writeLine(outFile,'    #["bus_i", "type", "Pd", "Qd", "Gs", "Bs", "area", "Vm", "Va", "baseKV", "zone", "Vmax", "Vmin"]')
	for n,data in sorted(busData.items()):
		writeLine(outFile,"""    %s,""" % data)
	writeLine(outFile,"""])""")

	# Write generators as ["bus","Pg","Qg","Qmax","Qmin","Vg","mBase","status","Pmax","Pmin","Pc1","Pc2","Qc1min","Qc1max","Qc2min","Qc2max","ramp_agc","ramp_10","ramp_30","ramp_q","apf"]
	writeLine(outFile,"""## Gen data""")
	writeLine(outFile,"""ppc["gen"] = array([""")
	writeLine(outFile,'    #["bus","Pg","Qg","Qmax","Qmin","Vg","mBase","status","Pmax","Pmin","Pc1","Pc2","Qc1min","Qc1max","Qc2min","Qc2max","ramp_agc","ramp_10","ramp_30","ramp_q","apf"]')
	# Write slack bus
	writeLine(outFile, """    %s,""" % [slackBusId, 0,   0, 300, -300, 1, 100, 1, 250, -250, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
	# Write generators
	generators = [] # As a list of values
	for gen in generators: 	# TODO
		writeLine(outFile,"""    %s,""" % gen)
	writeLine(outFile,"""])""")

	# Write branches as ["fbus", "tbus", "r", "x", "b", "rateA", "rateB", "rateC", "ratio", "angle", "status", "angmin", "angmax"]
	# First, filtering the data and ordering it according to the internalId
	branchData = {}
	for u,v,edata in networkGraph.edges(data=True):
		# TODO Handle case of multi-branches
		Zb = (baseKV*1e3)**2/(baseMVA*1e6) # in Ohm
		r = round(edata['R1'] / Zb,6) # in p.u.
		x = round(edata['X1'] / Zb,6) # in p.u.
		b = 0 if edata['C1'] == 0 else round(edata['C1']*1e-6 * Zb * baseFrequency,6) # in p.u. !!!
		branchData[edata['internalId']] = [u,v,r,x,b,round(edata['pMax']/1e6,6),0,0,0,0,int(edata['closed']),-360,360]

	# Writing to file
	writeLine(outFile,"""## Branch data""")
	writeLine(outFile,"""ppc["branch"] = array([""")
	writeLine(outFile,'    #["fbus", "tbus", "r", "x", "b", "rateA", "rateB", "rateC", "ratio", "angle", "status", "angmin", "angmax"]')
	for n,data in sorted(branchData.items()):
		writeLine(outFile,"""    %s,""" % data)
	writeLine(outFile,"""])""")

	# write back matter
	writeLine(outFile,"""return ppc""")
	outFile.close()

# Starting point from python #
if __name__ == "__main__":
	main(sys.argv[1:])