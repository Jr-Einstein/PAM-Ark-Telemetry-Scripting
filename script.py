import pandas as pd
import os

def createFinalMasterReport(sourceFolder, outputFolder):
    # 1. Get a list of all Excel files in your Reports folder
    allFiles = [f for f in os.listdir(sourceFolder) if f.endswith('.xlsx')]
    
    if not allFiles:
        print("No Excel files found in the source folder.")
        return

    # 2. Define the three columns that link every sheet together
    matchKeys = ['SAFENUMBER', 'SAFENAME', 'SAFEURLID']

    # 3. Start with the first file as our base (usually safe_members2)
    # We sort to ensure a consistent starting point
    allFiles.sort() 
    finalData = pd.read_excel(os.path.join(sourceFolder, allFiles[0]))

    # 4. Loop through all remaining files and merge them one by one
    for fileName in allFiles[1:]:
        currentSheet = pd.read_excel(os.path.join(sourceFolder, fileName))
        
        # We use 'left' to keep all rows from our base and attach new columns
        finalData = pd.merge(finalData, currentSheet, on=matchKeys, how='left')

    # 5. Clean up any duplicate columns that might occur during merging
    # (Optional: removes columns ending in _y or _x if they are redundant)
    finalData = finalData.loc[:, ~finalData.columns.duplicated()]

    # 6. Save the final compiled result
    outputPath = os.path.join(outputFolder, 'MasterCompiledReport.xlsx')
    finalData.to_excel(outputPath, index=False)
    
    print(f"Success! Combined {len(allFiles)} files into one report.")

# --- Set your paths and run ---
source = r'c:\Users\n369560\Downloads\Reports'
destination = r'c:\Users\n369560\Downloads\CombinedReport'

createFinalMasterReport(source, destination)
