# ylpic-reader

This repository contains the python file reading the Excel files with the data of the Ylpic network.
This code and the data are provided within the framework of the Gredor project of the public service of Wallonia – Department of Energy and Sustainable Building.

## Quick start
To use the code you need first need Python 3 (https://www.python.org/downloads/).
Once installed, some packages need to be installed using the following commands:
    pip3 install xlrd
    pip3 install networkx
    pip3 install pydot2

Place the Excel data files (not included in the repository for confidentiality) in a folder named "ylpic" next to the python files and in the terminal use:
    python3 dotConverter.py ylpic

Two files, "network.dot" and "network.pdf" should appear.

## Programmer instructions
The python code "dotConverter.py" is an example of use of the data.
The raw information contained in the Excel files is read by the codes "networkMaker.py" and "scenariosReader.py".
The two commands
	graph=networkMaker.makeNetwork(folderPath)
	scenariosReader.readScenarios(folderPath,2020,1,graph,'H')
create a NetworkX graph with the raw data attached to each node or edge as a dictionnary.
