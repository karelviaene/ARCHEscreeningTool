import streamlit as st
import pandas as pd
import re
import requests
from bs4 import BeautifulSoup
from io import BytesIO, StringIO
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import logging
import zipfile
import random
import time
from datetime import datetime
import json
import copy
st.title("ARCHE screener")

uploaded_file = st.file_uploader("Upload SoC Template", type=["xlsx"])

# In-memory log stream
if "log_stream" not in st.session_state:
    st.session_state.log_stream = StringIO()
    logging.basicConfig(stream=st.session_state.log_stream, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')

def process_data(file):
    logging.info("Started ED screener process")
    # Load file with CAS (or other input strings)
    CASallpd = pd.read_excel(file, engine="openpyxl", sheet_name="COMPOSITION", skiprows=7)   # CAS column only on row 8 in SoC template
    if "CAS" not in CASallpd.columns:
        st.error("Error: 'CAS' column not found.")
        return None
    CASall = CASallpd["CAS"].dropna().tolist()
    CASall = [re.sub(r'[^\d\-]', '', str(cas)) for cas in CASall]   # Make sure there are no issues with CAS eg whitespaces
    N_CAS = len(CASall)
    # Load file again usen openpyxl so that we can save the collected info to the template
    templateSoC = load_workbook(file)

    # Create dictionary to save screening output in
    clp_info = [{"id": i + 1} for i in range(N_CAS)]  # Create list of dictionaries with length number of CAS numbers
    now = datetime.now()
    for i, entry in enumerate(clp_info):  # Add CAS and date to all entries
        entry["Input"] = CASall[i]
        entry["Date collected"] = now.strftime("%d/%m/%Y %H:%M:%S")
    # List of names to add as keys
    key_names = [
        "CAS", "EC", "Name ECHA-CHEM", "ECHA-CHEM checked", "REACH tonnage band", "On C&L?", "Entries C&L",
        "C&L URL", "C&L Type", "Joint Entries", "Classification - Hazard classes",
        "Classification - Hazard statements", "Classification - Organs/ExposureRoute",
        "Labeling - Hazard statements", "Labeling - Supplementary Hazard statements",
        "Labeling - Organs/ExposureRoute", "Specific concentration limits", "M-factors", "C&L notes"
    ]
    # Add empty key-value pairs using dictionary unpacking
    clp_info = [{**entry, **{key: "-" for key in key_names}} for entry in clp_info]

    #### LOAD DATA SOURCES ####
    # C&L info from NextSDS API
    st.write(f"Checking NextSDS API")
    data = [{"casNumber": cas, "ecNumber": ""} for cas in CASall]
    url = "https://api.nextsds.com/echa"
    headers = {
        "accept": "application/json",
        "Authorization": "Bearer b4077cae-b5b0-49a3-9c93-9925740adfe6",
        "Content-Type": "application/json"
    }
    try:
        response = requests.post(url, headers=headers, json=data)
        logging.info(f"API status: {response.status_code}")
        if response.json()["output"]:
            CnL_json = response.json()["output"]
        else:
            st.write(f"JSON issue: {response.json()["error"]}")
    except Exception as e:
        logging.error(f"Error parsing JSON: {e}")
        st.write("Failed to parse response from API.")
        st.write(f"JSON issue: {response.json()["error"]}")

    #### LOOP OVER ALL CAS NUMBERS AND ADD CLASSIFICATION ####
    sheet_comp = templateSoC["COMPOSITION"] # To update COMPOSITION sheet while going through the json
    i = 0
    while i < len(clp_info):
        #### ECHA-CHEM C&L from NEXTSDS-API ####
        # Extract required info from response json
        # Find the entries matching the input string in the json
        matching_entries = [entryJSON for entryJSON in CnL_json if entryJSON.get("casNumber") == clp_info[i]["Input"]]
        if len(matching_entries) > 1:   # If there are multiple hits for input string (eg multiple EC for CAS)
            if clp_info[i]["ECHA-CHEM checked"] == "-": # If this is the first time the input string is checked
                clp_info[i]["ECHA-CHEM checked"] = 0   # Indicate that this is the first of multiple hits
                for n, entry in enumerate(matching_entries[1:], start=1):  # Skip the first match
                    copied_entry = copy.deepcopy(clp_info[i])  # Deep copy the original entry
                    copied_entry.update(entry)  # Merge with the new matching entry
                    copied_entry["ECHA-CHEM checked"] = n
                    clp_info.insert(i + n, copied_entry)  # Insert at position i+n
            entry = matching_entries[clp_info[i]["ECHA-CHEM checked"]] # Use the corresponding entry in the json
        else: entry = matching_entries[0]   # Standard, take the first entry if only one hit
        if entry.get("found") == False:  # If the chemical was NOT found on C&L
            clp_info[i]["On C&L?"] = "No"
        else:  # If the chemical was found on C&L (then there is no "found" entry)
            clp_info[i]["On C&L?"] = "Yes"
            clp_info[i]["CAS"] = entry.get("cas")
            clp_info[i]["EC"] = entry.get("ecNumber")
            clp_info[i]["Name ECHA-CHEM"] = entry.get("name")
            clp_info[i]["REACH tonnage band"] = entry.get("tonnageBand")
            # If harmonised classification, give this
            if entry.get("type") == "harmonised":
                clp_info[i]["C&L Type"] = "Harmonised C&L"
                clp_info[i]["C&L URL"] = "https://chem.echa.europa.eu/" + entry.get("rmlId") + "/harmonised"
                clp_info[i]["Entries C&L"] = entry.get("totalIndustryClassifications")
            else:  # Self-classification by industry
                clp_info[i]["C&L Type"] = "Notified C&L"
                clp_info[i]["C&L URL"] = "https://chem.echa.europa.eu/" + entry.get("rmlId") + "/self-classified"
                clp_info[i]["Entries C&L"] = entry.get("totalIndustryClassifications")
                clp_info[i]["Joint Entries"] = entry.get("industryClassification")[0]["dataSource"]
            clp_info[i]["Classification - Hazard classes"] = entry.get("hazards")["hazardClasses"]
            clp_info[i]["Classification - Hazard statements"] = entry.get("hazards")["statements"]
            clp_info[i]["Classification - Organs/ExposureRoute"] = entry.get("hazards")["targetOrgsAndRoutes"]
            clp_info[i]["Labeling - Hazard statements"] = entry.get("hazards")["statements"]
            # clp_info[i]["Labeling - Supplementary Hazard statements"] = entry.get("labelling")["targetOrgsAndRoutes"]
            # clp_info[i]["Labeling - Organs/ExposureRoute"] = entry.get("labelling")["targetOrgsAndRoutes"]
            clp_info[i]["Specific concentration limits"] = entry.get("hazards")["scl"]
            clp_info[i]["M-factors"] = entry.get("hazards")["mFactors"]
            if entry.get("notes"):
                print(entry.get("notes")[0])
                clp_info[i]["C&L notes"] = entry.get("notes")[0]["note"]["noteCode"] + ": " + \
                                           entry.get("notes")[0]["note"]["noteText"]
        logging.info("Finished Next-SDS API")

        # Finalize the loop per chemical
        logging.info(f"Processed {i+1}/{len(clp_info)}: {clp_info[i]["CAS"]}")
        st.write(f"Processed {i+1}/{len(clp_info)}: {clp_info[i]["CAS"]}")

        i += 1  # Update in the while loop


    output_excel = BytesIO()
    templateSoC.save(output_excel)
    # df.to_excel(output_excel, index=False, engine="openpyxl")
    # output_excel.seek(0)

    ### SAVING ####
    # Save to zip file
    st.session_state.log_stream.seek(0)
    log_bytes = BytesIO(st.session_state.log_stream.read().encode("utf-8"))
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        zip_file.writestr("EDscreener_results.xlsx", output_excel.getvalue())
        zip_file.writestr("EDscreener_log.txt", log_bytes.getvalue())
        json_data = json.dumps(response.json()["output"], indent=2)
        json_bytes = json_data.encode('utf-8')
        zip_file.writestr(f"databases/API json {datetime.now().strftime('%Y-%m-%d %H-%M')}.json", json_bytes)
    zip_buffer.seek(0)

    return zip_buffer

if uploaded_file:
    if st.button("Run Screener"):
        st.info("Processing started...")
        zip_result = process_data(uploaded_file)
        if zip_result:
            st.download_button("Download All Results (ZIP)", zip_result, file_name=f"EDscreener_package_{datetime.now().strftime("%Y-%m-%d %H-%M")}.zip")
            st.success("Processing finished!")
