import pandas as pd
import os
import time

def startCyberArkReportMerge(sourcePath, outputPath):
    # 1. Identify all valid files
    print("0% - Scanning directory for reports...")
    allFiles = [f for f in os.listdir(sourcePath) if f.endswith(('.xlsx', '.csv'))]
    
    if not allFiles:
        print("Error: No Excel or CSV files found.")
        return

    # Constants for matching
    matchKeys = ['SAFENUMBER', 'SAFENAME', 'SAFEURLID']
    
    # 2. Pick the base file (safe_members2)
    # We use this as the foundation because it has the most rows (one per member)
    print("10% - Locating member report for base mapping...")
    membersFile = next((f for f in allFiles if 'member' in f.lower()), allFiles[0])
    
    # 3. Reading the base file
    print(f"25% - Reading base file: {membersFile}...")
    filePath = os.path.join(sourcePath, membersFile)
    if membersFile.endswith('.csv'):
        masterData = pd.read_csv(filePath)
    else:
        masterData = pd.read_excel(filePath)

    # 4. Loop through other files and merge
    otherFiles = [f for f in allFiles if f != membersFile]
    totalFiles = len(otherFiles)

    for index, fileName in enumerate(otherFiles):
        # Calculate progress between 30% and 80%
        progress = 30 + int((index / totalFiles) * 50)
        print(f"{progress}% - Merging data from: {fileName}...")
        
        currentFilePath = os.path.join(sourcePath, fileName)
        if fileName.endswith('.csv'):
            currentData = pd.read_csv(currentFilePath)
        else:
            currentData = pd.read_excel(currentFilePath)
        
        # This 'left' merge replicates safe1 data for every matching member row
        masterData = pd.merge(masterData, currentData, on=matchKeys, how='left')

    # 5. Clean up columns (Remove duplicates if the same column exists in multiple files)
    print("85% - Cleaning up duplicate columns and organizing layout...")
    masterData = masterData.loc[:, ~masterData.columns.duplicated()]
    
    # Move identifiers to the front
    cols = list(masterData.columns)
    for key in reversed(matchKeys):
        if key in cols:
            cols.insert(0, cols.pop(cols.index(key)))
    masterData = masterData[cols]

    # 6. Save final report
    print("95% - Almost finished! Writing final Excel file...")
    finalPath = os.path.join(outputPath, 'CyberArk_Combined_Master_Report.xlsx')
    masterData.to_excel(finalPath, index=False)
    
    print(f"100% - Done! Master report is ready at: {finalPath}")

# --- Set your folders and run ---
sourceDir = r'c:\Users\n369560\Downloads\Reports'
outputDir = r'c:\Users\n369560\Downloads\CombinedReport'

startCyberArkReportMerge(sourceDir, outputDir)
