import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
import requests
from bs4 import BeautifulSoup
from io import BytesIO
import openpyxl
import logging

def setup_logging(folder_path):
    log_folder = os.path.join(folder_path, "log")
    os.makedirs(log_folder, exist_ok=True)
    log_filename = f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
    log_path = os.path.join(log_folder, log_filename)
    logging.basicConfig(filename=log_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    return log_path

def file_downloaded_today(path):
    if os.path.exists(path):
        mod_time = datetime.fromtimestamp(os.path.getmtime(path)).date()
        return mod_time == datetime.now().date()
    return False

def process_data(uploaded_file, folder_path):
    log_path = setup_logging(folder_path)
    logging.info("Started ED screener process")

    CASallpd = pd.read_excel(uploaded_file)
    if "CAS" not in CASallpd.columns:
        st.error("Error: 'CAS' column not found.")
        return None

    CASall = CASallpd["CAS"].dropna().tolist()
    CASall = [re.sub(r'[^\d\-]', '', str(cas)) for cas in CASall]
    N_CAS = len(CASall)

    clp_info = [{"id": i+1, "Input": CASall[i]} for i in range(N_CAS)]
    key_names = [
        "Input", "ED PPP: Yes/No", "ED PPP: Status", "ED PPP: Conclusion HH",
        "ED PPP: Conclusion non-TO", "ED PPP: EFSA conclusion link"
    ]
    for entry in clp_info:
        for key in key_names:
            entry[key] = "-"

    databases_folder = os.path.join(folder_path, "databases")
    os.makedirs(databases_folder, exist_ok=True)

    efsaPPP_url = "https://www.efsa.europa.eu/en/applications/pesticides"
    PPP_ED_string = "overview-endocrine-disrupting-assessment-pesticide-active-substances"
    responseEFSA = requests.get(efsaPPP_url)
    ED_PPP = None

    if responseEFSA.status_code == 200:
        soupEFSA = BeautifulSoup(responseEFSA.text, "html.parser")
        matching_links = [link.get("href") for link in soupEFSA.find_all("a", href=True)
                          if PPP_ED_string in link.get("href") and link.get("href").endswith((".xls", ".xlsx"))]
        if matching_links:
            file_url = requests.compat.urljoin(efsaPPP_url, matching_links[0])
            ED_PPP = requests.get(file_url)

    datetoday = datetime.now().strftime("%Y-%m-%d")
    file_name = "PPP ED list " + datetoday + ".xlsx"
    file_path_PPP_ED = os.path.join(databases_folder, file_name)

    if not file_downloaded_today(file_path_PPP_ED) and ED_PPP and ED_PPP.status_code == 200:
        with open(file_path_PPP_ED, "wb") as file:
            file.write(ED_PPP.content)
        logging.info(f"Saved PPP ED list: {file_path_PPP_ED}")

    for i, entry in enumerate(clp_info):
        if ED_PPP and ED_PPP.status_code == 200:
            excel_data = BytesIO(ED_PPP.content)
            workbook = openpyxl.load_workbook(excel_data)
            sheet = workbook.worksheets[0]
            found = False
            for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
                for cell in row:
                    val = str(cell.value).strip()
                    if val == entry["Input"]:
                        entry["ED PPP: Yes/No"] = "Yes"
                        entry["ED PPP: Status"] = sheet[f"H{cell.row}"].value
                        entry["ED PPP: Conclusion HH"] = sheet[f"I{cell.row}"].value
                        entry["ED PPP: Conclusion non-TO"] = sheet[f"J{cell.row}"].value
                        entry["ED PPP: EFSA conclusion link"] = sheet[f"N{cell.row}"].value
                        found = True
                        break
                if found:
                    break
        st.write(f"Processed {i+1}/{N_CAS}")
        logging.info(f"Processed {i+1}/{N_CAS}")

    df = pd.DataFrame(clp_info)
    now = datetime.now().strftime("%Y-%m-%d %H-%M")
    output_file = os.path.join(folder_path, f"output_EDscreener_export_{now}.xlsx")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)
    logging.info(f"Saved to {output_file}")
    return output_file

# Streamlit UI
st.title("ED Screener Tool")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
folder_path = st.text_input("Enter output folder path", value=os.getcwd())

if uploaded_file and folder_path:
    if st.button("Run Screener"):
        st.info("Processing started...")
        result_path = process_data(uploaded_file, folder_path)
        if result_path:
            with open(result_path, "rb") as f:
                st.download_button("Download Results", f, file_name=os.path.basename(result_path))
            st.success("Processing finished!")