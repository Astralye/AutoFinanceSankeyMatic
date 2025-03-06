from datetime import datetime 
import pandas as pd
import math

def getFile(fileName):
    try:
        # Data starts at E53
        # Read past this point
        return pd.read_excel(fileName,
                            sheet_name="Nodes",
                            thousands=",")
    except:
        print("Error! File does not exist!")
        return None

def createNodeStructure(df):


    # Indexable size
    indexableNodes = df.shape[0] - 1 

    # Check for empty size
    if indexableNodes == 0:
        return

    # Names
    name = "Name"
    sumValueAddress = "Cell_address/es"
    childCell = "Children_address"

    # String to add to file after creation
    fileString = ""

    # Gets overwritten every node
    tmpText = ""

    for i in range(indexableNodes):
        
        # Get object
        nodeObject = df.loc[i]

        # Retrieve values
        nodeName   = nodeObject[name]
        nodeValue  = nodeObject[sumValueAddress]
        nodeChild = nodeObject[childCell] # address

        # Check if empty node
        if pd.isna(nodeName): # dont know why nan is a string or to convert
            continue

        # check if not leaf node
        if pd.isna(nodeChild):
            continue

        childrenNodes = nodeChild.split(",")

        childrenObjects = []
        for cn in childrenNodes:
            childrenObjects.append(df.loc[ int(cn[1:]) - 2 ])

        for obj in childrenObjects:

            value = nodeValue if len(childrenNodes) == 1 else obj[sumValueAddress]

            fileString += f'{nodeName} [{int(value)}] {obj[name]}\n'
        
    return fileString

def createFile(fs):
    f = open(f'SankeyDiagram-{datetime.today().strftime('%Y%m%d')}.txt', "a")
    f.write(fs)
    f.close()

    print("Sankey Diagram successfully created")

def exitLog():
    input("Press enter to exit.")

def main():
    df = getFile('2025\February_2025.xlsx')

    if df is None:
        exitLog()
        return

    fs = createNodeStructure(df)

    if fs is None:
        exitLog()
        return

    createFile(fs)
    exitLog()
    
main()
