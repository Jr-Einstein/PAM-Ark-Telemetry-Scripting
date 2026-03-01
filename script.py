import pandas as pd # Loading pandas to handle all the heavy spreadsheet work
import os # Using os to talk to my computer's folders
import time # Bringing in time to see how fast my script runs

def runMyReportMerger(inputFolder, outputFolder):
    startTime = time.time() # Capture the exact moment we start the process
    
    print("Looking for report files...")
    # I am grabbing every file that ends in .xlsx or .csv so I don't miss anything
    allFiles = [f for f in os.listdir(inputFolder) if f.endswith(('.xlsx', '.csv'))]
    
    # Safety check: if the folder is empty, I need to stop here
    if not allFiles:
        print("No files found!")
        return

    # These are the 3 "Anchor" columns that connect all my different reports
    keyColumns = ['SAFENUMBER', 'SAFENAME', 'SAFEURLID']
    
    # I want to start with the 'member' file because it has the most rows (the base)
    baseFile = next((f for f in allFiles if 'member' in f.lower()), allFiles[0])
    
    print(f"Starting with base: {baseFile}")
    # Building the full path so Python knows exactly where the base file lives
    pathToBase = os.path.join(inputFolder, baseFile)
    
    # Reading the first file - I added a check here for both CSV and Excel formats
    if baseFile.endswith('.csv'):
        mainDf = pd.read_csv(pathToBase)
    else:
        mainDf = pd.read_excel(pathToBase)

    # Now I'm looping through every OTHER file in that folder
    for fileName in allFiles:
        # I skip the base file because I've already loaded it above
        if fileName == baseFile:
            continue
            
        print(f"Adding data from {fileName}...")
        # Get the location of the next file to merge
        currentFilePath = os.path.join(inputFolder, fileName)
        
        # Load the new file (handles csv/excel) into a temporary workspace
        tempDf = pd.read_csv(currentFilePath) if fileName.endswith('.csv') else pd.read_excel(currentFilePath)
        
        # This is the most important part: I am mapping the data side-by-side
        # 'how=left' ensures I don't lose any members even if they aren't in the other sheet
        mainDf = pd.merge(mainDf, tempDf, on=keyColumns, how='left')

    # If I merge multiple files, some columns might repeat - this line deletes the duplicates
    mainDf = mainDf.loc[:, ~mainDf.columns.duplicated()]

    # I want the final report to be clean, so I am putting the 3 ID columns at the very start
    allColumns = list(mainDf.columns)
    for k in reversed(keyColumns):
        if k in allColumns:
            # I'm popping the ID column out and moving it to the front of the list
            allColumns.insert(0, allColumns.pop(allColumns.index(k)))
    
    # Re-ordering the entire spreadsheet based on my new column list
    mainDf = mainDf[allColumns]

    print("Saving the final report...")
    # Setting the name for the final combined master file
    finalSavePath = os.path.join(outputFolder, 'CyberArkFinalReport.xlsx')
    # Converting my digital data into an actual Excel file on my drive
    mainDf.to_excel(finalSavePath, index=False)
    
    # Check the clock again to see when we finished
    endTime = time.time()
    # Subtracting start from end to get the total duration
    totalTime = round(endTime - startTime, 2)
    
    print(f"Done! Report saved in {totalTime} seconds.")

# Setting the source folder where my daily reports land
myReports = r'c:\Users\n369560\Downloads\Reports'
# Setting the destination folder where I want the final combined file
myOutput = r'c:\Users\n369560\Downloads\CombinedReport'

# Actually starting the script with my specific folders
runMyReportMerger(myReports, myOutput)
