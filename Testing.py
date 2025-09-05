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

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
file_BPR_ED = st.file_uploader("Upload BPR ED file (xlsx)", type=["xlsx"])
file_food_add = st.file_uploader("Upload Food additives Excel (xlsx)", type=["xlsx"])
file_food_flav = st.file_uploader("Upload Food flavourings Excel (xlsx)", type=["xlsx"])

# In-memory log stream
if "log_stream" not in st.session_state:
    st.session_state.log_stream = StringIO()
    logging.basicConfig(stream=st.session_state.log_stream, level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')

def process_data(file):
    logging.info("Started ED screener process")
    # Load file with CAS (or other input strings)
    CASallpd = pd.read_excel(file, engine="openpyxl")
    if "CAS" not in CASallpd.columns:
        st.error("Error: 'CAS' column not found.")
        return None
    CASall = CASallpd["CAS"].dropna().tolist()
    CASall = [re.sub(r'[^\d\-]', '', str(cas)) for cas in CASall]
    N_CAS = len(CASall)

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
        "Labeling - Organs/ExposureRoute", "Specific concentration limits", "M-factors", "C&L notes",
        "ED PPP: Yes/No", "ED PPP: Status", "ED PPP: Conclusion HH", "ED PPP: Conclusion non-TO",
        "ED PPP: EFSA conclusion link",
        "BPR: Yes/No", "BPR: ED HH", "BPR: ED ENV",
        "ED Assessment List: Yes/No", "ED Assessment List: Outcome",
        "ED Assessment List: Status", "ED Assessment List: Authority", "ED Assessment List: Last updated",
        "SVHC: Yes/No", "SVHC: Reason", "SVHC: Date Inclusion", "SVHC: Decision",
        "Food additive: Yes/No", "Food additive: E number", "Food flavourings: Yes/No", "Food flavourings: FL",
        "SVHC intent: Yes/No", "SVHC intent: Status", "SVHC intent: Scope", "SVHC intent: Last updated",
        "PACT: Yes/No", "PACT: SEv", "PACT: SEv link", "PACT: DEv", "PACT: DEv link", "PACT: ED", "PACT: ED link",
        "PACT: ARN", "PACT: ARN link", "PACT: PBT", "PACT: PBT link", "PACT: CLH", "PACT: CLH link", "PACT: SVHC",
        "PACT: SVHC link",
        "CoRAP: Yes/No", "CoRAP: Initial grounds of Concern", "CoRAP: Status", "CoRAP: Latest update"
    ]
    # Add empty key-value pairs using dictionary unpacking
    clp_info = [{**entry, **{key: "-" for key in key_names}} for entry in clp_info]

    # Set up for requests (webscraping)
    user_agents_list = [
        'Mozilla/5.0 (iPad; CPU OS 12_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.83 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36'
        'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:91.0) Gecko/20100101 Firefox/91.0',
        'Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15A372 Safari/604.1',
        'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 11_2_3) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Safari/605.1.15',
        'Mozilla/5.0 (Linux; Android 10; SM-G973F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.78 Mobile Safari/537.36',
        'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:98.0) Gecko/20100101 Firefox/98.0',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/117.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 13_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Linux; Android 11; Pixel 5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.5735.131 Mobile Safari/537.36',
        'Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1',
        'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko',
        'Mozilla/5.0 (Linux; Android 12; SM-A525F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.5845.92 Mobile Safari/537.36',
        'Mozilla/5.0 (iPad; CPU OS 15_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.5 Mobile/15E148 Safari/604.1',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 12_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.1 Safari/605.1.15',
        'Mozilla/5.0 (Windows NT 10.0; ARM64; rv:110.0) Gecko/20100101 Firefox/110.0'
    ]

    # C&L info from NextSDS API (Using JOBS)
    logging.info("Starting nextSDS API")
    st.write("Checking ECHA-CHEM API")
    api_key = "b4077cae-b5b0-49a3-9c93-9925740adfe6"
    start_url = "https://api.nextsds.com/jobs/start"
    status_url = "https://api.nextsds.com//jobs/retrieve"
    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    data = {
        "taskId": "echa-api",
        "payload": [{"casNumber": cas, "ecNumber": ""} for cas in CASall]
    }

    try:
        # response_start = requests.post(start_url, headers=headers, json=data)
        # logging.info(f"API status: {response_start.status_code}")
        # if response_start.status_code == 200:   # If job submitted successfully
        #     job_id = response_start.json().get("id")
            job_id = "run_cmf5neapz139h2qoc0g9jkqru"
            logging.info("Job submitted successfully: ", job_id)
            st.write("Job submitted successfully to API: ", job_id)

            # job_finished = False
            # job_finished = True

            # while not job_finished:
                # time.sleep(10)
            status_response = requests.get(f"{status_url}/{job_id}", headers=headers)
            if status_response.status_code == 200:
                status_data = status_response.json()
                job_status = status_data.get("status")
                logging.info(f"Job status: {job_status}")
                st.write(f"Job status: {job_status}")
                if job_status not in ["STARTED", "EXECUTING"]:
                    job_finished = True
            elif status_response.status_code in [400, 404]:
                logging.warning(f"Job error ({status_response.status_code}): {job_id}")
                st.write(f"Job error ({status_response.status_code}): {job_id}")
                    # job_finished = True

            # Once the job is finished, collect the response
            output = status_response.json().get("output")
            if output:
                CnL_json = output
            else:
                error_msg = status_response.json().get("error", "Unknown error")
                st.write(f"JSON issue: {error_msg}")

        # elif response_start.status_code == 400:
        #     logging.warning("Job failed to submit to API")
        #     st.write("Job failed to submit to API")


    except Exception as e:
        logging.error(f"Failed to retrieve JSON from API")
        st.write("Failed to retrieve JSON from API")
        # st.write(f"JSON issue: {response_start.json()["error"]}")

    logging.info("NextSDS API successfull")
    st.write("JSON successfully retrieved from API")

    #### LOOP OVER ALL CAS NUMBERS AND SCREEN SOURCES ####

    # st.write(CnL_json)
    i = 0
    while i < len(clp_info):
        st.write(f"Checking chemical: {clp_info[i]["Input"]}")

        #### ECHA-CHEM C&L from NEXTSDS-API ####
        # Find the entries matching the input string in the json
        matching_entries = [entryJSON for entryJSON in CnL_json if
                            entryJSON.get("casNumber") == clp_info[i]["Input"]]
        if len(matching_entries) > 1:  # If there are multiple hits for input string (eg multiple EC for CAS)
            if clp_info[i]["ECHA-CHEM checked"] == "-":  # If this is the first time the input string is checked
                clp_info[i]["ECHA-CHEM checked"] = 0  # Indicate that this is the first of multiple hits
                for n, entry in enumerate(matching_entries[1:], start=1):  # Skip the first match
                    copied_entry = copy.deepcopy(clp_info[i])  # Deep copy the original entry
                    copied_entry.update(entry)  # Merge with the new matching entry
                    copied_entry["ECHA-CHEM checked"] = n
                    clp_info.insert(i + n, copied_entry)  # Insert at position i+n
            entry = matching_entries[clp_info[i]["ECHA-CHEM checked"]]  # Use the corresponding entry in the json
        else:
            entry = matching_entries[0]  # Standard, take the first entry if only one hit
        # Extract required info from response json
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
            labelling = entry.get("labelling", [])
            hazard_codes = [
                item["hazardStatement"]["hazardStatementCode"]
                for item in labelling
                if "hazardStatement" in item and "hazardStatementCode" in item["hazardStatement"]
            ]
            clp_info[i]["Labeling - Hazard statements"] = ", ".join(hazard_codes)
            # clp_info[i]["Labeling - Supplementary Hazard statements"] = entry.get("labelling")["targetOrgsAndRoutes"]
            # clp_info[i]["Labeling - Organs/ExposureRoute"] = entry.get("labelling")["targetOrgsAndRoutes"]
            clp_info[i]["Specific concentration limits"] = entry.get("hazards")["scl"]
            mfactor = entry.get("mFactor", {})
            if mfactor:
                items = mfactor.get("items", [])
                if isinstance(items, list) and items:
                    mfactor_strings = []
                    for item in items:
                        acute = item.get("mfactorAcute", "-")
                        chronic = item.get("mfactorChronic", "-")
                        mfactor_strings.append(f"Acute: {acute}; Chronic: {chronic}")
                    clp_info[i]["M-factors"] = " | ".join(mfactor_strings)
                elif isinstance(mfactor, dict):
                    clp_info[i]["M-factors"] = "Acute: " + str(mfactor.get("mfactorAcute", "-")) + \
                                               "; Chronic: " + str(mfactor.get("mfactorChronic", "-"))
            if entry.get("notes"):
                clp_info[i]["C&L notes"] = entry.get("notes")[0]["note"]["noteCode"] + ": " + \
                                           entry.get("notes")[0]["note"]["noteText"]
        logging.info("Finished Next-SDS API")
        i += 1  # Update in the while loop

    st.write(clp_info)

if uploaded_file:
    if st.button("Run Screener"):
        st.info("Processing started...")
        zip_result = process_data(uploaded_file)
        if zip_result:
            st.download_button("Download All Results (ZIP)", zip_result, file_name=f"EDscreener_package_{datetime.now().strftime("%Y-%m-%d %H-%M")}.zip")
            st.success("Processing finished!")

        else:
            # If processing failed, offer log download
            st.session_state.log_stream.seek(0)
            log_text = st.session_state.log_stream.read()
            st.download_button("Download Log Only", log_text, file_name="EDscreener_log.txt")
            st.error("Processing failed. You can download the log for details.")
