import os
import re
import xml.etree.ElementTree as ET
import openpyxl
import streamlit as st
import io

# --- Your existing functions (modified if needed) ---

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

    # --- Find marker column: locate the cell with "START" ---
    marker_index = None
    for row in rows:
        cells = row.findall("ss:Cell", ns)
        for i, cell in enumerate(cells):
            data = cell.find("ss:Data", ns)
            if data is not None and data.text and data.text.strip().upper() == "START":
                marker_index = i
                break
        if marker_index is not None:
            break
    if marker_index is None:
        raise ValueError("START marker not found.")

    # --- Collect boundary markers ---
    boundaries = []
    for idx, row in enumerate(rows):
        cells = row.findall("ss:Cell", ns)
        if marker_index < len(cells):
            marker_cell = cells[marker_index].find("ss:Data", ns)
            if marker_cell is not None and marker_cell.text and marker_cell.text.strip():
                marker_text = marker_cell.text.strip().upper()
                if marker_text == "START" or marker_text.startswith("MINUTO") or marker_text == "STOP":
                    time_val = ""
                    if marker_index - 2 >= 0 and marker_index - 2 < len(cells):
                        time_cell = cells[marker_index - 2].find("ss:Data", ns)
                        if time_cell is not None and time_cell.text:
                            time_val = time_cell.text.strip()
                    boundaries.append({"row_index": idx, "time": time_val})
    if len(boundaries) < 2:
        raise ValueError("Not enough boundary markers found.")

    if len(boundaries) < 7:
        raise ValueError("Expected at least 7 boundary markers (START, MINUTO 1...MINUTO 6).")

    # --- Compute per-minute averages ---
    num_metrics = 18
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

    # --- Build per-metric lists ---
    variables = []
    for k in range(num_metrics):
        vals = []
        for m in range(6):
            vals.append(per_minute_avgs[m][k])
        variables.append(vals)

    def avg_list(vals):
        valid = [v for v in vals if v is not None]
        return sum(valid)/len(valid) if valid else None

    aggregates = []
    for vals in variables:
        aggregates.append({
            "third1": avg_list(vals[0:2]),
            "third2": avg_list(vals[2:4]),
            "third3": avg_list(vals[4:6]),
            "half1": avg_list(vals[0:3]),
            "half2": avg_list(vals[3:6]),
            "total": avg_list(vals[0:6])
        })

    var_names = ["V'O2", "V'CO2", "V'O2/kg", "VCO2kg", "V'E", "RER", "FC", "CHO", "FAT",
                 "EE", "EECHO", "EEFAT", "METS", "Borg", "BF", "WR", "V'E/V'O2", "V'E/V'CO2"]
    boundaries_times = [boundaries[i]["time"] for i in range(7)]
    subject_data = {"surname": surname, "name": name, "boundaries": boundaries_times, "variables": {}}
    for idx, var in enumerate(var_names):
        var_data = {}
        for m in range(1, 7):
            var_data[f"{var}_{m}"] = variables[idx][m-1]
        for suffix in ["third1", "third2", "third3", "half1", "half2", "total"]:
            var_data[f"{var}_{suffix}"] = aggregates[idx][suffix]
        subject_data["variables"][var] = var_data

    return subject_data

# --- Streamlit Interface ---

def main():
    st.title("6MWT Data Extraction")
    st.write("Upload your XML file(s) to extract and process the data.")

    uploaded_files = st.file_uploader("Choose XML file(s)", type=["xml"], accept_multiple_files=True)

    if uploaded_files:
        all_subjects = []
        for uploaded_file in uploaded_files:
            try:
                # uploaded_file is a BytesIO object
                subject_data = process_file(uploaded_file)
                all_subjects.append(subject_data)
                st.success(f"Processed file: {uploaded_file.name}")
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {e}")

        if all_subjects:
            # Optionally, show one subject's data or provide a download button for the Excel file.
            st.write("Extracted Data:")
            st.json(all_subjects)

            # --- Generate and Download Excel File ---
            import openpyxl
            from openpyxl import Workbook
            from io import BytesIO

            wb = Workbook()
            ws = wb.active
            ws.title = "6MWT Data"

            # Build header:
            headers = [f"Edge_{i}" for i in range(1, 8)]
            headers.extend(["Surname", "Name"])
            var_names = ["V'O2", "V'CO2", "V'O2/kg", "VCO2kg", "V'E", "RER", "FC", "CHO", "FAT",
                         "EE", "EECHO", "EEFAT", "METS", "Borg", "BF", "WR", "V'E/V'O2", "V'E/V'CO2"]
            for var in var_names:
                for m in range(1, 7):
                    headers.append(f"{var}_{m}")
                for suffix in ["third1", "third2", "third3", "half1", "half2", "total"]:
                    headers.append(f"{var}_{suffix}")
            ws.append(headers)

            for subj in all_subjects:
                row = []
                row.extend(subj["boundaries"])
                row.extend([subj["surname"], subj["name"]])
                for var in var_names:
                    var_data = subj["variables"].get(var, {})
                    for m in range(1, 7):
                        row.append(var_data.get(f"{var}_{m}"))
                    for suffix in ["third1", "third2", "third3", "half1", "half2", "total"]:
                        row.append(var_data.get(f"{var}_{suffix}"))
                ws.append(row)

            # Save workbook to a BytesIO stream
            excel_io = BytesIO()
            wb.save(excel_io)
            excel_io.seek(0)

            st.download_button(label="Download Excel File", data=excel_io, file_name="martinez.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
