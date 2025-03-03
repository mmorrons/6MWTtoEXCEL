import os
import re
import xml.etree.ElementTree as ET
import openpyxl
import streamlit as st
import io

def process_file(file_path_or_buffer):
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
    # Parse the XML; streamlit file uploader returns a BytesIO object.
    tree = ET.parse(file_path_or_buffer)
    root = tree.getroot()
    rows = root.findall(".//ss:Row", ns)

    # --- Extract subject's surname and name ---
    surname = ""
    name = ""
    for row in rows:
        cells = row.findall("ss:Cell", ns)
        for i, cell in enumerate(cells):
            data = cell.find("ss:Data", ns)
            if data is not None and data.text:
                txt = data.text.strip().upper()
                if txt == "COGNOME" and i+1 < len(cells):
                    next_cell = cells[i+1].find("ss:Data", ns)
                    if next_cell is not None and next_cell.text:
                        surname = next_cell.text.strip()
                elif txt == "NOME" and i+1 < len(cells):
                    next_cell = cells[i+1].find("ss:Data", ns)
                    if next_cell is not None and next_cell.text:
                        name = next_cell.text.strip()

    # --- Determine the Marker column ---
    # Search for the header row containing "MARKER" (case-insensitive).
    marker_index = None
    for row in rows:
        cells = row.findall("ss:Cell", ns)
        for i, cell in enumerate(cells):
            data = cell.find("ss:Data", ns)
            if data is not None and data.text and data.text.strip().upper() == "MARKER":
                marker_index = i
                break
        if marker_index is not None:
            break
    if marker_index is None:
        raise ValueError("Marker column not found.")

    # --- Collect boundary markers based on new logic ---
    # Expect a "START" marker (assigned 0.0 m) and then six numeric markers.
    boundaries = []
    found_start = False
    for idx, row in enumerate(rows):
        cells = row.findall("ss:Cell", ns)
        if marker_index < len(cells):
            data = cells[marker_index].find("ss:Data", ns)
            marker_text = data.text.strip() if (data is not None and data.text) else ""
            # Get time from two cells before the marker (if available)
            time_val = ""
            if marker_index - 2 >= 0 and marker_index - 2 < len(cells):
                time_cell = cells[marker_index - 2].find("ss:Data", ns)
                if time_cell is not None and time_cell.text:
                    time_val = time_cell.text.strip()
            if not found_start:
                if marker_text.upper() == "START":
                    boundaries.append({"row_index": idx, "time": time_val, "meter": 0.0})
                    found_start = True
            else:
                if marker_text != "":
                    try:
                        meter_val = float(marker_text.replace(',', '.'))
                        boundaries.append({"row_index": idx, "time": time_val, "meter": meter_val})
                    except Exception as e:
                        continue
            if len(boundaries) == 7:
                break
    if len(boundaries) < 7:
        raise ValueError("Not enough boundaries found (expected START plus six numeric markers).")

    # --- Compute per-minute averages for physiological metrics ---
    # The columns immediately after the Marker column contain the 17 physiological metrics:
    # V'O2, V'O2/kg, V'E, RER, FC, CHO, FAT, EE, EECHO, EEFAT, METS, V'CO2, V'E/V'CO2, BF, EE/BSA, EE/kg, EE/kg/magra
    num_metrics = 17
    per_minute_avgs = []
    for interval in range(6):
        start_idx = boundaries[interval]["row_index"]
        end_idx = boundaries[interval+1]["row_index"]
        sum_metrics = [0.0] * num_metrics
        count_metrics = [0] * num_metrics
        for r in range(start_idx, end_idx + 1):
            row = rows[r]
            cells = row.findall("ss:Cell", ns)
            for j in range(marker_index+1, marker_index+1+num_metrics):
                if j < len(cells):
                    cell_val = cells[j].find("ss:Data", ns)
                    if cell_val is not None and cell_val.text:
                        raw = cell_val.text.strip().replace(',', '.')
                        try:
                            val = float(raw)
                        except:
                            val = None
                        if val is not None:
                            sum_metrics[j - (marker_index+1)] += val
                            count_metrics[j - (marker_index+1)] += 1
        avg_metrics = []
        for k in range(num_metrics):
            avg_metrics.append(sum_metrics[k] / count_metrics[k] if count_metrics[k] > 0 else None)
        per_minute_avgs.append(avg_metrics)

    # --- Build per-metric lists and aggregate values ---
    variables = {}
    var_names = ["V'O2", "V'O2/kg", "V'E", "RER", "FC", "CHO", "FAT",
                 "EE", "EECHO", "EEFAT", "METS", "V'CO2", "V'E/V'CO2", "BF", "EE/BSA", "EE/kg", "EE/kg/magra"]
    for idx, var in enumerate(var_names):
        var_data = {}
        for m in range(1, 7):
            var_data[f"{var}_{m}"] = per_minute_avgs[m-1][idx]
        # Aggregates: compute averages for groups of minutes.
        def avg_list(vals):
            valid = [v for v in vals if v is not None]
            return sum(valid)/len(valid) if valid else None
        vals = [per_minute_avgs[m][idx] for m in range(6)]
        var_data["third1"] = avg_list(vals[0:2])
        var_data["third2"] = avg_list(vals[2:4])
        var_data["third3"] = avg_list(vals[4:6])
        var_data["half1"] = avg_list(vals[0:3])
        var_data["half2"] = avg_list(vals[3:6])
        var_data["total"] = avg_list(vals)
        variables[var] = var_data

    subject_data = {
        "surname": surname,
        "name": name,
        "boundaries": [b["time"] for b in boundaries],
        "meters": [b["meter"] for b in boundaries],
        "variables": variables
    }
    return subject_data

def main():
    st.title("6MWT Data Extraction")
    st.write("Upload your XML file(s) to extract and process the data.")

    uploaded_files = st.file_uploader("Choose XML file(s)", type=["xml"], accept_multiple_files=True)
    if uploaded_files:
        all_subjects = []
        for uploaded_file in uploaded_files:
            try:
                subject_data = process_file(uploaded_file)
                all_subjects.append(subject_data)
                st.success(f"Processed file: {uploaded_file.name}")
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {e}")
        if all_subjects:
            st.write("Extracted Data:")
            st.json(all_subjects)

            # --- Let the user choose the Excel file name ---
            excel_filename = st.text_input("Enter the Excel file name (include .xlsx):", value="6MWT_Data.xlsx")

            # --- Generate and Download Excel File ---
            from openpyxl import Workbook
            from io import BytesIO
            wb = Workbook()
            ws = wb.active
            ws.title = "6MWT Data"

            # Build header:
            headers = [f"Edge_{i}" for i in range(1, 8)]
            headers.extend(["Surname", "Name", "ID"])
            for var in ["V'O2", "V'O2/kg", "V'E", "RER", "FC", "CHO", "FAT",
                        "EE", "EECHO", "EEFAT", "METS", "V'CO2", "V'E/V'CO2", "BF", "EE/BSA", "EE/kg", "EE/kg/magra"]:
                for m in range(1, 7):
                    headers.append(f"{var}_{m}")
                for suffix in ["third1", "third2", "third3", "half1", "half2", "total"]:
                    headers.append(f"{var}_{suffix}")
            ws.append(headers)

            for subj in all_subjects:
                row = []
                row.extend(subj["boundaries"])
                row.extend([subj["surname"], subj["name"], ""])
                for var in ["V'O2", "V'O2/kg", "V'E", "RER", "FC", "CHO", "FAT",
                            "EE", "EECHO", "EEFAT", "METS", "V'CO2", "V'E/V'CO2", "BF", "EE/BSA", "EE/kg", "EE/kg/magra"]:
                    var_data = subj["variables"].get(var, {})
                    for m in range(1, 7):
                        row.append(var_data.get(f"{var}_{m}"))
                    for suffix in ["third1", "third2", "third3", "half1", "half2", "total"]:
                        row.append(var_data.get(f"{var}_{suffix}"))
                ws.append(row)

            excel_io = BytesIO()
            wb.save(excel_io)
            excel_io.seek(0)
            st.download_button(label="Download Excel File", data=excel_io, file_name=excel_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
