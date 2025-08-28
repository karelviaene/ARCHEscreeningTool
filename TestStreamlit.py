print("Running ED screener")
print("Importing packages")

#### Code to get C&L url based on a CAS number ####
# Original code found on this URL: https://stackoverflow.com/questions/60698025/how-to-get-resulting-url-from-search
import pandas as pd
import time
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO
import os
import random
import requests
from bs4 import BeautifulSoup

# Function to load in a file
def select_file():
    # Create a hidden root window
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Make sure the window appears on top (important for macOS)
    root.call('wm', 'attributes', '.', '-topmost', True)

    # Open the file selection dialog
    file_path = filedialog.askopenfilename(
        title="Select an input file",
        filetypes=[("XLSX files", "*.xlsx"), ("All files", "*.*")],
        initialdir="~"
    )
    # Return the selected file path or None
    return file_path

# Function to select target directory
def select_folder():
    # Create a hidden root window
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Make sure the window appears on top (important for macOS)
    root.call('wm', 'attributes', '.', '-topmost', True)

    # Open the file selection dialog
    folder_path = filedialog.askdirectory(
        title="Select folder",
    )
    # Return the selected folder path or None
    return folder_path

# Function to check if file was downloaded today
today = datetime.now().date()
def file_downloaded_today(file_path):
    # Check if the file exists
    if os.path.exists(file_path):
        # Get the file's modification time
        file_mod_time = datetime.fromtimestamp(os.path.getmtime(file_path)).date()
        # Check if the file was modified today
        return file_mod_time == today
    return False


### SET UP ###
# CASall = ['100-51-6']
print("Loading xlsx file")
file_path = select_file()
CASallpd = pd.read_excel(file_path)
if "CAS" in CASallpd.columns:
    CASall = CASallpd['CAS'].dropna().tolist()
else:
    print("The 'CAS' column was not found in the Excel file.")

N_CAS = len(CASall)       # TO BE UPDATED LATER TO LENGTH OF LIST OF CAS NUMBERS
clp_info = [{"id": i+1} for i in range(N_CAS)]  # Create list of dictionaries with length number of CAS numbers
now = datetime.now()
for i, entry in enumerate(clp_info):  # Add CAS and date to all entries
    entry["Input"] = CASall[i]
    entry["Date collected"] = now.strftime("%d/%m/%Y %H:%M:%S")

# List of names to add as keys
key_names = [
     "Input","CAS","EC","Name ECHA-CHEM","ECHA-CHEM checked","REACH tonnage band","On C&L?","Entries C&L","C&L URL","C&L Type","Joint Entries","Classification - Hazard classes",
     "Classification - Hazard statements","Classification - Organs/ExposureRoute",
     "Labeling - Hazard statements", "Labeling - Supplementary Hazard statements",
     "Labeling - Organs/ExposureRoute","Specific concentration limits","M-factors","C&L notes",
     "ED PPP: Yes/No","ED PPP: Status","ED PPP: Conclusion HH","ED PPP: Conclusion non-TO",
     "ED PPP: EFSA conclusion link",
     "BPR: Yes/No","BPR: ED HH","BPR: ED ENV",
     "ED Assessment List: Yes/No","ED Assessment List: Outcome",
     "ED Assessment List: Status","ED Assessment List: Authority","ED Assessment List: Last updated",
     "SVHC: Yes/No","SVHC: Reason","SVHC: Date Inclusion","SVHC: Decision",
     "Food additive: Yes/No","Food additive: E number","Food flavourings: Yes/No","Food flavourings: FL",
     "SVHC intent: Yes/No","SVHC intent: Status","SVHC intent: Scope","SVHC intent: Last updated",
     "PACT: Yes/No","PACT: SEv","PACT: SEv link","PACT: DEv","PACT: DEv link","PACT: ED","PACT: ED link",
     "PACT: ARN","PACT: ARN link","PACT: PBT","PACT: PBT link","PACT: CLH","PACT: CLH link","PACT: SVHC",
     "PACT: SVHC link",
     "CoRAP: Yes/No","CoRAP: Initial grounds of Concern","CoRAP: Status","CoRAP: Latest update"
     ]
# Add empty key-value pairs using dictionary unpacking
clp_info = [{**entry, **{key: "-" for key in key_names}} for entry in clp_info]
# Also set up a json for saving further info
clp_json = {}

# Create folders and set up to current directory
#currentdir = select_folder()   # Ask for user input
currentdir = os.getcwd()
folderdatabases = os.path.join(currentdir, "databases")
if not os.path.exists(folderdatabases):
    os.makedirs(folderdatabases)

# Get the current date time as string
formatted_time = datetime.now().strftime("%Y-%m-%d %H-%M")  # Customize format as needed
datetoday = datetime.now().strftime("%Y-%m-%d")
row = -1         # For the loop

############################################################################################

### LOAD DATABASES ####

#### Get latest Excel ED PPP and save it

# Step 1: Specify the target website and output folder
efsaPPP_url = "https://www.efsa.europa.eu/en/applications/pesticides"  # Replace with the actual website URL
PPP_ED_string = "overview-endocrine-disrupting-assessment-pesticide-active-substances"
    # Step 2: Scrape the website to find the Excel file
responseEFSA = requests.get(efsaPPP_url)
if responseEFSA.status_code == 200:
    soupEFSA = BeautifulSoup(responseEFSA.text, "html.parser")
    # Find all links containing the search string and ending on xls / xlsx
    matching_linksEFSA = [
        linkEFSA.get("href") for linkEFSA in soupEFSA.find_all("a", href=True)
        if PPP_ED_string in linkEFSA.get("href") and linkEFSA.get("href").endswith((".xls", ".xlsx"))
    ]
    if matching_linksEFSA:
        print(f"Found {len(matching_linksEFSA)} Excel file containing ED PPP on https://www.efsa.europa.eu/en/applications/pesticides.")
    # Step 3: Save the Excel to memory for further use & download to folder databases
        for link in matching_linksEFSA:
            file_url = requests.compat.urljoin(efsaPPP_url, link)
            ED_PPP = requests.get(file_url)     # To be used later in the script
            # Step 0: Check if file already was downloaded today
            file_name = "PPP ED list " + datetoday + ".xlsx"
            file_path_PPP_ED = os.path.join(folderdatabases, file_name)
            # Only download if file was not downloaded already today
            if file_downloaded_today(file_path_PPP_ED):
                print(f"The PPP ED list was already downloaded today, not updating.")
            else:
                if ED_PPP.status_code == 200:       # Download to databases folder
                    with open(file_path_PPP_ED, "wb") as file:
                        file.write(ED_PPP.content)
                    print(f"Saved PPP ED list: {file_path_PPP_ED}")
                else:
                    print(f"Failed to download PPP ED file: {file_url}")
    else:
        print(f"No Excel files found on the EFSA PPP website containing the string '{PPP_ED_string}'.")
    soupEFSA.decompose()  # Close/destroy the BS tree
else:
    print(f"Failed to access the EFSA PPP website. Status code: {responseEFSA.status_code}")
responseEFSA.close()


##################################################################
##### LOOP OVER ALL CAS NUMBERS IN LIST ######
##################################################################

# Set up for error handling
max_retries = 20
retry_delay = 3  # seconds

# Quality check the CAS numbers
for i in range(len(CASall)):
    CASall[i] = re.sub(r'[^\d-]', '', CASall[i])    # Keep digits and hyphens


# Use while so that you can add CAS/EC in case of multiple entries
i = 0
while i < len(clp_info):
    CAS = CASall[i]

    row = i
    clp_info[row]["Input"] = CAS
    progress = (i + 1) / len(clp_info) * 100
    print("Checking " + CAS + " (" + f"Progress: {progress:.2f}%" +  ")")

    for attempt in range(max_retries):
        time.sleep(random.uniform(1, 3))  # Random delay between 1 and 2 seconds to avoid ECHA webscraping detection

        try:
            chemlink = "https://chem.echa.europa.eu/substance-search?searchText=" + CAS

            ### EXTRACT INFO FROM THE EFSA PPP ED Excel ####
            ################################################
            if ED_PPP.status_code == 200:     # If Excel with ED PPP was found succesfully
                # Step 2: Load the Excel file into memory
                excel_data = BytesIO(ED_PPP.content)
                workbook = openpyxl.load_workbook(excel_data)

                # Step 3: Access the first sheet
                first_sheet = workbook.worksheets[0]

                # Step 4: Search for a specific string in the first sheet
                found_value = None

                for rowExcel in first_sheet.iter_rows(min_row=1, max_row=first_sheet.max_row, values_only=False):
                    for cell in rowExcel:
                        cell_value = str(cell.value).strip()
                        if cell_value != "-" and cell_value in (clp_info[row]["CAS"], clp_info[row]["EC"],clp_info[row]["Input"]):  # Check both CAS or EC
                            print(f"Found '{CAS}' in row {cell.row}, column {cell.column} on PPP ED list.")
                            # Extract the value from column H of the same row
                            PPPstatus = first_sheet[f"H{cell.row}"].value
                            found_value = PPPstatus
                            PPP_HH = first_sheet[f"I{cell.row}"].value
                            PPP_nonTO = first_sheet[f"J{cell.row}"].value
                            PPP_link = first_sheet[f"N{cell.row}"].value
                            clp_info[row]["ED PPP: Yes/No"] = "Yes"
                            clp_info[row]["ED PPP: Status"] = PPPstatus
                            clp_info[row]["ED PPP: Conclusion HH"] = PPP_HH
                            clp_info[row]["ED PPP: Conclusion non-TO"] = PPP_nonTO
                            clp_info[row]["ED PPP: EFSA conclusion link"] = PPP_link
                            break
                    if found_value:
                        break

                if not found_value:
                    clp_info[row]["ED PPP: Yes/No"] = "No"
            else:
                print(f"Failed to download the PPP ED file. Status code: {ED_PPP.status_code}")

            print("PPP list done")

            break  # Success, exit the loop

        except OSError as e:
            print(f"Error encountered, attempt {attempt + 1} of {max_retries}. Retrying after delay...")
            time.sleep(retry_delay)

        else:
            print("All retry attempts failed due to errors.")

    i += 1  # Update in the while loop

####### END LOOP W####################################
########################################################

# Write away as Excel file
df = pd.DataFrame(clp_info)
# Sort by the first column (usually Column A in Excel)
df = df.sort_values(by=df.columns[0], ascending=True)

exportpath = os.path.join(currentdir,"output")
if not os.path.exists(exportpath):
    os.makedirs(exportpath)
exportfile = os.path.join(exportpath, "EDscreener export " + formatted_time +".xlsx")
print("Saving results to " + exportfile)
df.to_excel(exportfile, index=False)
