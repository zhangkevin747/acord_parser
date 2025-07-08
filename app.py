import streamlit as st
import os
import re
import pdfplumber
import pandas as pd
from io import BytesIO
from typing import List
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# --- Parser Functions (same as yours) ---
def extract_lines_from_pdf(pdf_file) -> List[str]:
    lines = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split('\n'))
    return [line.strip() for line in lines if line.strip()]

def full_block_vehicle_parser(lines: List[str]) -> pd.DataFrame:
    vehicles = []
    last_make = "Unknown"
    for i, line in enumerate(lines):
        if "MAKE:" in line:
            make_match = re.search(r'MAKE:\s*([A-Za-z0-9]+)', line)
            if make_match:
                last_make = make_match.group(1)

        match = re.search(r'(?P<veh_no>\d+)\s+(?P<year>\d{4})\s+MODEL:\s*(?P<model>.*?)\s+V\.I\.N\.\:\s*(?P<vin>[A-Z0-9]+)', line)
        if match:
            current_vehicle = {
                'Vehicle No': int(match.group("veh_no")),
                'Year': int(match.group("year")),
                'Model': match.group("model").strip(),
                'VIN': match.group("vin").strip(),
                'Make': last_make,
                'Street': '',
                'City': '',
                'State': '',
                'Zip': '',
                'Cost New': ''
            }
            for j in range(1, 6):
                if i + j >= len(lines):
                    break
                next_line = lines[i + j]
                if next_line.startswith("ADDRESS"):
                    addr_match = re.search(r'ADDRESS\s+(.*)\s+([A-Z]{2})\s+(\d{5})', next_line)
                    if addr_match:
                        current_vehicle['Street'] = addr_match.group(1).strip()
                        current_vehicle['State'] = addr_match.group(2)
                        current_vehicle['Zip'] = addr_match.group(3)
                if "$" in next_line:
                    cost_match = re.search(r'\$\s*([\d,]+)', next_line)
                    if cost_match:
                        current_vehicle['Cost New'] = cost_match.group(1).replace(",", "")
            vehicles.append(current_vehicle)
    return pd.DataFrame(vehicles)

def garage_location_parser(lines: List[str]) -> pd.DataFrame:
    garages = []
    for i, line in enumerate(lines):
        if line.startswith("LOC #"):
            try:
                street = lines[i + 1].strip()
                loc_line = lines[i + 2] if i + 2 < len(lines) else ""
                loc_num_match = re.match(r'^(\d+)\b', loc_line)
                loc_num = int(loc_num_match.group(1)) if loc_num_match else None
                city_state_line = lines[i + 4] if i + 4 < len(lines) else ""
                city_match = re.search(r'CITY:([A-Za-z\s]+?)\s+STATE:', city_state_line)
                state_match = re.search(r'STATE:([A-Z]{2})', city_state_line)
                zip_line = lines[i + 5] if i + 5 < len(lines) else ""
                zip_match = re.search(r'\b(\d{5})\b', zip_line)
                if loc_num and "OWNER OCCUPIED" not in street:
                    garages.append({
                        "Location Number": loc_num,
                        "Street": street,
                        "City": city_match.group(1).strip() if city_match else "",
                        "State": state_match.group(1).strip() if state_match else "",
                        "Zip": zip_match.group(1) if zip_match else ""
                    })
            except:
                pass
    return pd.DataFrame(garages)

def generate_excel(vehicle_df, garage_df):
    excel_headers = [
        "Vehicle No", "Year", "Make", "Model", "VIN", "Garage Location #",
        "CSL Limit", "MedPay Limit", "Physical Damage Deductible",
        "Size GVW, GCW or Seating Capacity", "Original Cost new", "Rating Type",
        "Business Use", "Radius of Operations", "Classification Group",
        "Truckers Classification", "Specialized Delivery Classification",
        "Food Delivery Classification", "Waste Disposal Classification",
        "Farmers Classification", "Demp and Transit Mix Classification",
        "Contractors Classification", "Not Otherwise Specified Classification",
        "Dumping Capability", "Operator Experience", "Use"
    ]
    output_df = pd.DataFrame(columns=excel_headers)
    output_df["Vehicle No"] = vehicle_df.get("Vehicle No", "")
    output_df["Year"] = vehicle_df.get("Year", "")
    output_df["Make"] = vehicle_df.get("Make", "")
    output_df["Model"] = vehicle_df.get("Model", "")
    output_df["VIN"] = vehicle_df.get("VIN", "")
    output_df["Original Cost new"] = vehicle_df.get("Cost New", "")

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Vehicles"
    for row in dataframe_to_rows(output_df, index=False, header=True):
        ws1.append(row)

    ws2 = wb.create_sheet(title="Garages")
    for row in dataframe_to_rows(garage_df, index=False, header=True):
        ws2.append(row)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

import streamlit as st

# --- Page Config ---
st.set_page_config(
    page_title="ACORD Parser",
    layout="centered"
)

# --- Logo ---
from PIL import Image
logo = Image.open("PLMR_BIG.png")  # Ensure logo.png is in the same folder or adjust the path
st.image(logo, use_container_width=True)

# --- Title ---
st.markdown("<h2 style='text-align: center;'>ACORD Form Parser</h2>", unsafe_allow_html=True)
st.markdown("---")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload an ACORD 129 PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("Processing the uploaded document..."):
        lines = extract_lines_from_pdf(uploaded_file)
        vehicle_df = full_block_vehicle_parser(lines)
        garage_df = garage_location_parser(lines)
        excel_file = generate_excel(vehicle_df, garage_df)

    st.success("Extraction complete. You can now download the parsed Excel file.")

    # --- Download Button ---
    st.download_button(
        label="Download Excel File",
        data=excel_file,
        file_name="acord_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- Optional Data Preview ---
    with st.expander("View Parsed Vehicle Table"):
        st.dataframe(vehicle_df)

    with st.expander("View Parsed Garage Table"):
        st.dataframe(garage_df)
