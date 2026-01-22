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
import tkinter as tk
from tkinter import filedialog

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
        filetypes=[("XLSX files", "*.xlsx"),("TXT files","*.txt"), ("All files", "*.*")],
        initialdir="~"
    )
    # Return the selected file path or None
    return file_path

def chunk_list(lst, chunk_size):
    for i in range(0, len(lst), chunk_size):
        yield lst[i:i + chunk_size]

print("Loading xlsx file")
file_path = select_file()
CASallpd = pd.read_excel(file_path)
print(CASallpd)
if "CAS" in CASallpd.columns:
    CASall = CASallpd['CAS'].dropna().tolist()
else:
    print("The 'CAS' column was not found in the Excel file.")

# Clean and deduplicate CAS values
CASall = CASallpd["CAS"].dropna().tolist()  # selected_col is chosen in the app
CASall = [re.sub(r'[^\d\-]', '', str(cas)) for cas in CASall]
order = pd.unique(CASall)  # Save the order for reordering later on (needs to be unique values)
CASall = list(set(CASall))  # Keep only unique values
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

# NextSDS info
# Load API key
api_file = select_file()
if api_file is not None:
    with open(api_file, "r", encoding="utf-8") as f:
        api_key = f.read().strip()
    # now use api_key

start_url = "https://api.nextsds.com/jobs/start"
status_url = "https://api.nextsds.com/jobs/retrieve"
headers = {
    "accept": "application/json",
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json"
}

# Step 1: Submit all jobs
jobs = []
for idx, cas_chunk in enumerate(chunk_list(CASall, 150)):
    data = {
        "taskId": "echa-api",
        "payload": cas_chunk
    }
    try:
        response = requests.post(start_url, headers=headers, json=data)
        print(data)
        if response.status_code == 200:
            job_id = response.json().get("id")
            jobs.append({"id": job_id, "index": idx + 1, "done": False, "output": None})
            print(f"Chunk {idx + 1}: Job submitted successfully: {job_id}")
        else:
            print(f"Chunk {idx + 1}: Failed to submit job")
            logging.info(f"Chunk {idx + 1}: Failed to submit job")
    except Exception as e:
        print(f"Chunk {idx + 1}: Exception during job submission: {str(e)}")

# Step 2: Monitor all jobs in one loop
while not all(job["done"] for job in jobs):
    time.sleep(5)
    for job in jobs:
        if job["done"]:
            continue
        try:
            status_response = requests.get(f"{status_url}/{job['id']}", headers=headers)
            if status_response.status_code == 200:
                status_data = status_response.json()
                job_status = status_data.get("status")
                print(f"Chunk {job['index']}: Job status: {job_status}")
                # 🔥 Show detailed error info if FAILED
                if job_status == "FAILED":
                    print(f"Chunk {job['index']} - FAIL DETAILS:")
                    print(json.dumps(status_data, indent=2))  # full response
                if job_status not in ["STARTED", "EXECUTING","DEQUEUED"]:
                    job["done"] = True
                    job["output"] = status_data.get("output", [])
            elif status_response.status_code in [400, 404]:
                print(f"Chunk {job['index']}: Job error ({status_response.status_code})")
                try:
                    print(status_response.json())
                except:
                    print(status_response.text)
                job["done"] = True
        except Exception as e:
            print(f"Chunk {job['index']}: Exception during status check: {str(e)}")

# Step 3: Combine all outputs
CnL_json = []
for job in jobs:
    if job["output"]:
        CnL_json.extend(job["output"])

print("All JSON chunks successfully retrieved and combined")

# Save JSON to file
json_data = json.dumps(CnL_json, indent=2)
filename = "NextSDS API output.json"   # choose your filename

with open(filename, "w", encoding="utf-8") as f:
    f.write(json_data)

print("Saved JSON")
