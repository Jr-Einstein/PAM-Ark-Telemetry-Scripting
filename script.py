import pandas as pd # Import pandas to handle all the spreadsheet data manipulation
import os # Import os to manage file paths and folder scanning
import time # Import time to keep track of the script's run duration

def startCyberArkReportMerge(sourcePath, outputPath):
    # Capture the start time so I can calculate how long this takes at the end
    startTime = time.time()
    
    # First, I'm scanning my reports folder for any Excel or CSV files
    print("0% - Scanning directory for reports...")
    allFiles = [f for f in os.listdir(sourcePath) if f.endswith(('.xlsx', '.csv'))]
    
    # If the folder is empty or doesn't have the right files, I stop the script here
    if not allFiles:
        print("Error: No Excel or CSV files found.")
        return

    # These are the three core columns that link every CyberArk report together
    matchKeys = ['SAFENUMBER', 'SAFENAME', 'SAFEURLID']
    
    # I need to pick a starting file; I'm looking for the member report because it has the most rows
    print("10% - Locating member report for base mapping...")
    membersFile = next((f for f in allFiles if 'member' in f.lower()), allFiles[0])
    
    # I build the full path and load the member report into my master dataset
    print(f"25% - Reading base file: {membersFile}...")
    filePath = os.path.join(sourcePath, membersFile)
    if membersFile.endswith('.csv'):
        masterData = pd.read_csv(filePath) # Load as CSV
    else:
        masterData = pd.read_excel(filePath) # Load as Excel

    # Now I create a list of all the other files I need to merge into the master
    otherFiles = [f for f in allFiles if f != membersFile]
    totalFiles = len(otherFiles)

    # I loop through each remaining file and attach its columns to my master data
    for index, fileName in enumerate(otherFiles):
        # I calculate a progress percentage to show in the console while it works
        progress = 30 + int((index / totalFiles) * 50)
        print(f"{progress}% - Merging data from: {fileName}...")
        
        # Build path and read the current file
        currentFilePath = os.path.join(sourcePath, fileName)
        if fileName.endswith('.csv'):
            currentData = pd.read_csv(currentFilePath)
        else:
            currentData = pd.read_excel(currentFilePath)
        
        # This 'left' merge is the key: it maps the new info to the existing members
        # If a safe appears once in 'currentData', it replicates for every member in 'masterData'
        masterData = pd.merge(masterData, currentData, on=matchKeys, how='left')

    # After merging many files, I delete any duplicate columns that might have slipped in
    print("85% - Cleaning up duplicate columns and organizing layout...")
    masterData = masterData.loc[:, ~masterData.columns.duplicated()]
    
    # I'm moving my three primary ID columns to the very front so the report is easy to read
    cols = list(masterData.columns)
    for key in reversed(matchKeys):
        if key in cols:
            # Pop the column out and re-insert it at index 0
            cols.insert(0, cols.pop(cols.index(key)))
    masterData = masterData[cols]

    # Now I'm ready to save everything into a single master Excel sheet
    print("95% - Almost finished! Writing final Excel file...")
    finalPath = os.path.join(outputPath, 'CyberArk_Combined_Master_Report.xlsx')
    masterData.to_excel(finalPath, index=False) # Exporting without the index numbers
    
    # Calculate total time taken
    totalTime = round(time.time() - startTime, 2)
    print(f"100% - Done! Master report is ready at: {finalPath} (Process took {totalTime}s)")

# Setting my folder paths as raw strings to handle Windows backslashes
sourceDir = r'c:\Users\n369560\Downloads\Reports'
outputDir = r'c:\Users\n369560\Downloads\CombinedReport'

# Trigger the main function
startCyberArkReportMerge(sourceDir, outputDir)
