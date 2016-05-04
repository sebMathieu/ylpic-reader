## Convert the series of Excel files with the characteristics of the network to a pyflow file.
# Requires xlrd which can be installed with "pip3 install xlrd".
#@author Bertrand CORNELUSSE

# Last update 2015-09-29
#TODO Clean input reading
#TODO Generators
#TODO Transformers
#TODO Shunt admittances at buses

import sys, os, copy
import networkx
import math

import networkMaker

## Entry point of the program.
# @param argv Program parameters: data folder path, slack bus id, period of interest (scenario hardcoded for now).
def main(argv):
    # Provide slack bus external ID as second argument
    if len(argv) == 1:
        folderPath=argv[0] if argv[0].endswith(("/","\\")) else argv[0]+"/"
        if not os.path.exists(folderPath):
            raise Exception('Folder \"%s\" does not exists.' % folderPath)

        # Read network data (all but power information)
        graph=networkMaker.makeNetwork(folderPath)

        # Output
        makeDGPFile('caseYlpic.csv', graph)
    else:
        print('Please provide data folder path')

## Convenience function
def writeLine(outFile,str2Write):
    outFile.write("""    %s\n""" % str2Write)
    return

## Make the file input for pyflow.
# @param fileName Output file name.
# @param networkGraph Graph with the daily scenario.
def makeDGPFile(fileName, networkGraph):

    # Convert multigraph into a simple graph, without multiedges
    simpleGraph = networkx.Graph()
    for u,vdata in networkGraph.nodes(data=True):
        simpleGraph.add_node(u,vdata)
    for u,v,edata in networkGraph.edges(data=True):
        if not simpleGraph.has_edge(u,v):
            simpleGraph.add_edge(u,v,edata)
        else:
            print(u,v)

    outFile = open(fileName,'w')

    # write front matter
    outFile.write("""Bus\n""")

    busData = {}
    for n,ndata in simpleGraph.nodes(data=True):
        outFile.write("""%d\n""" % n)

    # Write branches as ["fbus", "tbus", "r", "x", "b", "rateA", "rateB", "rateC", "ratio", "angle", "status", "angmin", "angmax"]
    # First, filtering the data and ordering it according to the internalId
    branchData = {}
    for u,v,edata in simpleGraph.edges(data=True):
        branchData[edata['internalId']] = [edata['internalId'],u,v,0,0,0,0,
                                           int(edata['closed']),float(edata['length'])]

    # Writing to file
    outFile.write("""id;from;to;r;x;b;Smax;closed;length\n""")
    for n,data in sorted(branchData.items()):
        outFile.write('%s\n' % (';'.join([str(d) for d in data])))

    outFile.close()

# Starting point from python #
if __name__ == "__main__":
    main(sys.argv[1:])