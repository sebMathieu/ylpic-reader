## ## Convert the series of Excel files into time series
# Requires xlrd which can be installed with "pip3 install xlrd".
#@author Sebastien MATHIEU

import sys
import datetime
import time
import copy

import xlwt

import networkMaker
import scenariosReader

## Entry point of the program.
# @param argv Program parameters: data folder path, slack bus id, period of interest (scenario hardcoded for now).
def main(argv):
	globalTic=time.time()

	# Default parameters
	year = 2020
	scenario = 'H'
	T = 96

	# Read instance folder
	if len(argv) < 1 :
		displayHelp()
		sys.exit(2)

	if len(argv) > 1:
		year=int(argv[0])
	if len(argv) > 2:
		scenario=argv[1]

	folderPath = argv[-1] if argv[-1].endswith(("/","\\")) else argv[-1]+"/"

	print("Compute time series %s %s" % (year, scenario))

	# Create work sheet
	wb = xlwt.Workbook()
	sheet = wb.add_sheet('timeseries')
	dateStyle = xlwt.easyxf(num_format_str='DD/MM/YYYY')

	# Header
	sheet.write(0,0,'Day')
	sheet.write(0,1,'Quarter')

	sheet.write(0,2,'Active production')
	sheet.write(1,2,'MW')

	sheet.write(0,3,'Active consumption')
	sheet.write(1,3,'MW')

	sheet.write(0,4,'Net active injection')
	sheet.write(1,4,'MW')

	sheet.write(0,5,'Reactive Production')
	sheet.write(1,5,'MVar')

	sheet.write(0,6,'Reactive Consumption')
	sheet.write(1,6,'MVar')

	sheet.write(0,7,'Net reactive injection')
	sheet.write(1,7,'MVar')

	# Read graph
	initGraph = networkMaker.makeNetwork(folderPath)

	# Read network data (all but power information)
	firstDay = datetime.date(year, 1, 1).toordinal()
	day = datetime.date(year, 1, 1)
	for d in range(0,365):
		tic = time.time()
		graph = copy.deepcopy(initGraph)

		# Obtain the corresponding date-time
		day=day.fromordinal(firstDay+d)

		# Read daily data
		scenariosReader.readScenarios(folderPath,year,d+1,graph,scenario)

		for t in range(0, T):
			l=2+d*T+t # Excel line

			# Write day and quarter
			sheet.write(l, 0, day, dateStyle)
			sheet.write(l, 1, t+1)

			# Find production & consumption
			activeProduction=0.0
			activeConsumption=0.0
			reactiveProduction=0.0
			reactiveConsumption=0.0

			for n, ndata in graph.nodes(data=True):
				Pd = 0 # MW
				Qd = 0 # MVar
				try:
					loadData = ndata['load']
					if loadData is not None:
						for label, baseline in loadData.activeProfiles.items():
							if len(baseline) < T:
								raise Exception("Error with label \"%s\" in node %s, baseline has %s periods in day %s." % (label, n, len(baseline),d+1))

						for baseline in loadData.activeProfiles.values():
							Pd += baseline[t]/1e3 # Convert from W to MW
						for baseline in loadData.reactiveProfiles.values():
							Qd += baseline[t]/1e3 # Convert from VAr to MVAr
				except KeyError:
					pass

				if Pd > 0:
					activeProduction+=Pd
					reactiveProduction+=Qd
				else:
					activeConsumption+=Pd
					reactiveConsumption+=Qd

			# Write
			sheet.write(l, 2, activeProduction)
			sheet.write(l, 3, activeConsumption)
			sheet.write(l, 4, activeProduction+activeConsumption)
			sheet.write(l, 5, reactiveProduction)
			sheet.write(l, 6, reactiveConsumption)
			sheet.write(l, 7, reactiveProduction+reactiveConsumption)

		print("\t day %s: %.3fs" % (d+1, time.time()-tic))

	# Save
	outputPath='timeseries-%s%s.xls' % (year,scenario)
	wb.save(outputPath)
	print('"%s" saved after %.2fs' % (outputPath, time.time()-globalTic))

## Display help of the program.
def displayHelp():
	text="Usage :\n\tpython timeseriesConverter.py [year scenario] dataFolder\n"
	print(text)

# Starting point from python #
if __name__ == "__main__":
	main(sys.argv[1:])
