import pandas as pd
import os
import time # Adding this back to track performance

def runMyReportMerger(inputFolder, outputFolder):
    startTime = time.time() # Start the clock
    
    print("Looking for report files...")
    allFiles = [f for f in os.listdir(inputFolder) if f.endswith(('.xlsx', '.csv'))]
    
    if not allFiles:
        print("No files found!")
        return

    keyColumns = ['SAFENUMBER', 'SAFENAME', 'SAFEURLID']
    baseFile = next((f for f in allFiles if 'member' in f.lower()), allFiles[0])
    
    print(f"Starting with base: {baseFile}")
    pathToBase = os.path.join(inputFolder, baseFile)
    
    if baseFile.endswith('.csv'):
        mainDf = pd.read_csv(pathToBase)
    else:
        mainDf = pd.read_excel(pathToBase)

    for fileName in allFiles:
        if fileName == baseFile:
            continue
            
        print(f"Adding data from {fileName}...")
        currentFilePath = os.path.join(inputFolder, fileName)
        
        tempDf = pd.read_csv(currentFilePath) if fileName.endswith('.csv') else pd.read_excel(currentFilePath)
        mainDf = pd.merge(mainDf, tempDf, on=keyColumns, how='left')

    mainDf = mainDf.loc[:, ~mainDf.columns.duplicated()]

    # Front-load the ID columns
    allColumns = list(mainDf.columns)
    for k in reversed(keyColumns):
        if k in allColumns:
            allColumns.insert(0, allColumns.pop(allColumns.index(k)))
    mainDf = mainDf[allColumns]

    print("Saving the final report...")
    finalSavePath = os.path.join(outputFolder, 'CyberArkFinalReport.xlsx')
    mainDf.to_excel(finalSavePath, index=False)
    
    # Calculate how long it took
    endTime = time.time()
    totalTime = round(endTime - startTime, 2)
    
    print(f"Done! Report saved in {totalTime} seconds.")

# Your paths
myReports = r'c:\Users\n369560\Downloads\Reports'
myOutput = r'c:\Users\n369560\Downloads\CombinedReport'

runMyReportMerger(myReports, myOutput)
