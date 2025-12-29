import os
import shutil
from datetime import datetime
import pandas as pd

EXCEL_FILE = 'Materiaalcontrole.xlsx'
KLASSEN = pd.ExcelFile(EXCEL_FILE).sheet_names
TODAY = pd.Timestamp.today().strftime('%Y-%m-%d')

### Step 1 - Create a dated copy of Materiaalcontrole.xlsx ###

new_filename = f'Materiaalcontrole_{TODAY}.xlsx'

if not os.path.exists(EXCEL_FILE):
    raise Exception(f"File {EXCEL_FILE} does not exist.")

if os.path.exists(new_filename):
    raise Exception(f"File {new_filename} already exists. You probably ran this script earlier today. Interrupting to avoid overwriting.")

print(f"Creating a dated copy {new_filename}")
shutil.copy2(EXCEL_FILE, new_filename)


### Step 2 - Archive sanctions in each sheet ###

print(f"Archiving sanctions in {EXCEL_FILE}")

# Save the updates to the Excel sheet, keeping the other sheets intact
with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    for sheet in KLASSEN:
        print(f"  - Converting sheet: {sheet}")
        students = pd.read_excel(EXCEL_FILE, sheet_name=sheet)

        # Make sure the 'Sancties Archief' column exists
        if "Sancties Archief" not in students.columns:
            print("    - Initialised column 'Sancties Archief'")
            students['Sancties Archief'] = 0
        
        # Update 'Sancties Archief' by summing all date columns
        students['Sancties Archief'] += students['Sancties']
        students['Sancties'] = 0
        print(f"    - Column 'Sancties' transferred to 'Sancties Archief'")

        # Remove all date columns
        date_columns = set(students.columns) - {'Voornaam', 'Sancties', 'Sancties Archief', 'Nota'}
        students.drop(columns=date_columns, inplace=True)
        print(f"    - Removed date columns: {date_columns}")

        # Write the updated DataFrame to the Excel sheet
        students.to_excel(writer, sheet_name=sheet, index=False)
        print(f"    - Updates written to {EXCEL_FILE}")
