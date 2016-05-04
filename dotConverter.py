## Convert the series of Excel files with the characteristics of the network to a dot representation and the corresponding pdf figure.
# Requires xlrd and pydot2 which can be installed with "pip3 install xlrd" and "pip3 install pydot2".
# Also requires Graphviz to convert the dot file to a pdf.
#@author Sebastien MATHIEU

import sys,os,subprocess
import networkx

import networkMaker
import scenariosReader

## Option to display transformers.
SHOW_TRANSFORMERS = False

## Entry point of the program.
# @param argv Program parameters.
def main(argv):
    # Read instance folder
    if len(argv) < 1 :
        displayHelp()
        sys.exit(2)
    folderPath=argv[-1] if argv[-1].endswith(("/","\\")) else argv[-1]+"/"
    if not os.path.exists(folderPath):
        raise Exception('Folder \"%s\" does not exists.' % folderPath)

    # Read data
    graph=networkMaker.makeNetwork(folderPath)
    scenariosReader.readScenarios(folderPath,2020,1,graph,'H')

    # Create simple graph and draw
    makeNetworkDot(graph)

## Make a dot file from a network.
# @param networkGraph Graph of the network.
def makeNetworkDot(networkGraph):
    # Constants definition
    # List of load types with a generation icon.
    GENERATION_TYPES=['G']
    # List of load types with an industrial icon.
    INDUSTRIAL_TYPES=['I1','I2','I3']

    # Map where keys are bus ids and value graphviz ids.
    idToGvId={}

    # Style of the graph
    g=networkx.Graph()

    # Create nodes
    unknownCount=len(networkGraph.nodes())
    transformersCount=0
    for n,ndata in networkGraph.nodes(data=True):
        # Take care of unknown nodes
        id=None
        internalId=id
        try:
            internalId=ndata["id"]
            id="BUS%s"%ndata["internalId"]
        except KeyError:
            unknownCount-=1
            internalId="?"
            id="BUS%s"%(-unknownCount)
        idToGvId[n]=id

        # Add styled node
        g.add_node(id,{"xlabel":internalId,"shape":"box","style":"filled","color":"#000000","fixedsize":"true","width":0.5,"height":0.075,"label":" ","fontsize":10,"id":id})

        # Transformer
        tId=None
        try:
            if SHOW_TRANSFORMERS and len(ndata["transformers"]) >= 1:
                t=ndata["transformers"][0]
                transformersCount+=1

                tId="TF%s"%t.internalId
                g.add_node(tId,{"xlabel":t.id,"shape":"circle","style":"filled","color":"#000000","fixedsize":"true","width":0.15,"height":0.15,"label":" ","fontsize":10})
                g.add_edge(id,tId,{"weight":10000})
        except KeyError:
            pass

        # Load
        load=ndata["load"]
        if load is not None:
            loadId="LOAD%s"%load.internalId
            edgeAttr={"weight":10000}
            if load.loadType in GENERATION_TYPES:
                g.add_node(loadId,{"label": "~","shape":"circle","style":"bold","color":"#000000","fixedsize":"true", "penwidth":2, "width":0.2, "height":0.2,"fontsize":18})
            elif load.loadType in INDUSTRIAL_TYPES:
                g.add_node(loadId,{"xlabel":load.loadType,"label":" ","shape":"invtriangle","color":"#000000","fixedsize":"true", "penwidth":2, "width":0.2, "height":0.2,"portPos":"n","fontsize":10})
                edgeAttr["tailport"]="n"
            else:
                g.add_node(loadId,{"xlabel":load.loadType,"label":" ","style":"filled","shape":"invtriangle","color":"#000000","fixedsize":"true", "penwidth":2, "width":0.2, "height":0.2,"portPos":"n","fontsize":10})
                edgeAttr["tailport"]="n"

            # Connect the load
            if tId is None:
                g.add_edge(id,loadId,edgeAttr)
            else:
                g.add_edge(tId,loadId,edgeAttr)

    # Create edges
    for u,v,edata in networkGraph.edges(data=True):
        if not edata["closed"]:
            continue
        uId=idToGvId[u]
        vId=idToGvId[v]
        if not g.has_edge(uId,vId):
            w=10/(1+edata["length"])
            g.add_edge(uId,vId,{"label":"     .     ","fontsize":10, "id": "LINE%s"%(edata["internalId"]+1)})

    # Detect cycles
    try:
        cycles = networkx.cycle_basis(g)
        for cy in cycles:
            networkx.set_node_attributes(g,'color',dict(zip(cy,['#FF0000']*len(cy))))
            print([g.node[b]['xlabel'] for b in cy])
        print("Cycles: %d" %len(cycles))
    except Exception as e:
        print("Error %s" % e)

    if not networkx.is_connected(g):
        print("Warning: graph is not connected.")
        components = [c for c in sorted(networkx.connected_components(g), key=len, reverse=True)]
        print("Warning: graph contains %d connected components:" % len(components))
        print(components)

    networkx.write_dot(g,'network.gv')

    # Convert to pdf and gml
    with open(os.devnull, 'wb') as devnull:
        subprocess.check_call(['neato', '-Tsvg', '-onetwork.svg', 'network.gv'], stdout=devnull, stderr=subprocess.STDOUT)
        subprocess.check_call(['neato', '-Tpdf', '-onetwork.pdf', 'network.gv'], stdout=devnull, stderr=subprocess.STDOUT)
        subprocess.check_call(['gv2gml', '-onetwork.gml', 'network.gv'], stdout=devnull, stderr=subprocess.STDOUT)

## Display help of the program.
def displayHelp():
    text="Usage :\n\tpython3 dotConverter.py dataFolder\n"
    print(text)

# Starting point from python #
if __name__ == "__main__":
    main(sys.argv[1:])