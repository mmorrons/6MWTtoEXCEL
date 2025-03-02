import os
import re
import xml.etree.ElementTree as ET
import openpyxl
import streamlit as st
import io

# --- Updated functions ---

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

    # --- Collect boundary markers and meter readings ---
    # We expect boundaries: START (assumed meter=0) and then 6 numeric values.
    boundaries = []
    for idx, row in enumerate(rows):
        cells = row.findall("ss:Cell", ns)
        if marker_index < len(cells):
            marker_cell = cells[marker_index].find("ss:Data", ns)
            if marker_cell is not None and marker_cell.text and marker_cell.text.strip():
                marker_text = marker_cell.text.strip().upper()
                # Get the time from 2 cells before the marker (as before)
                time_val = ""
                if marker_index - 2 >= 0 and marker_index - 2 < len(cells):
                    time_cell = cells[marker_index - 2].find("ss:Data", ns)
                    if time_cell is not None and time_cell.text:
                        time_val = time_cell.text.strip()
                meter_val = None
                if marker_text == "START":
                    meter_val = 0.0
                else:
                    # Try to convert the marker text to a float.
                    try:
                        meter_val = float(marker_text.replace(',', '.'))
                    except:
                        # If marker_text is STOP or non-numeric, skip it.
                        if marker_text == "STOP":
                            continue
                        else:
                            continue
                boundaries.append({"row_index": idx, "time": time_val, "meter": meter_val})
    if len(boundaries) < 7:
        raise ValueError("Not enough boundary markers found (expected at least 7).")
    
    # Take only the first 7 boundaries (START and 6 minutes)
    boundaries = boundaries[:7]
    boundaries_times = [b["time"] for b in boundaries]
    meters = [b["meter"] for b in boundaries]

    # --- Process physiological metrics (if needed later) ---
    # Now we have 17 metrics (columns 4 to 20, index 3 to 19)
    num_metrics = 17
    # Update var_names to match the new XML columns
    var_names = ["V'O2", "V'O2/kg", "V'E", "RER", "FC", "CHO", "FAT",
                 "EE", "EECHO", "EEFAT", "METS", "V'CO2", "V'E/V'CO2", "BF", "EE/BSA", "EE/kg", "EE/kg/magra"]

    # Compute per-minute averages (over the 6 intervals)
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

    # (The rest of the physiological variable aggregation remains available if needed.)
    aggregates = []
    for col in range(num_metrics):
        vals = [per_minute_avgs[m][col] for m in range(6)]
        valid = [v for v in vals if v is not None]
        avg_val = sum(valid)/len(valid) if valid else None
        aggregates.append({
            "third1": sum(vals[0:2])/2 if len(vals[0:2])==2 else None,
            "third2": sum(vals[2:4])/2 if len(vals[2:4])==2 else None,
            "third3": sum(vals[4:6])/2 if len(vals[4:6])==2 else None,
            "half1": sum(vals[0:3])/3 if len(vals[0:3])==3 else None,
            "half2": sum(vals[3:6])/3 if len(vals[3:6])==3 else None,
            "total": avg_val
        })

    subject_data = {
        "surname": surname,
        "name": name,
        "boundaries": boundaries_times,
        "meters": meters,  # list of meter values from boundaries (length 7)
        "variables": {}
    }
    for idx, var in enumerate(var_names):
        var_data = {}
        for m in range(1, 7):
            var_data[f"{var}_{m}"] = per_minute_avgs[m-1][idx] if m-1 < len(per_minute_avgs) else None
        for suffix in ["third1", "third2", "third3", "half1", "half2", "total"]:
            var_data[f"{var}_{suffix}"] = aggregates[idx][suffix]
        subject_data["variables"][var] = var_data

    return subject_data

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
            st.write("Extracted Data:")
            st.json(all_subjects)

            # --- Generate and Download Excel File ---
            from openpyxl import Workbook
            from io import BytesIO

            wb = Workbook()
            ws = wb.active
            ws.title = "6MWT Data"

            # Build header:
            # First, boundaries (time) columns (7 edges)
            headers = [f"Edge_{i}" for i in range(1, 8)]
            # Then Surname, Name and an empty ID column
            headers.extend(["Surname", "Name", "ID"])
            # Then the distance and speed columns:
            distance_cols = [
                "SixMWT_m_1", "SixMWT_m_2", "SixMWT_m_3", "SixMWT_m_4", "SixMWT_m_5", "SixMWT_m_6",
                "SixMWT_m_first2", "SixMWT_m_second2", "SixMWT_m_third2",
                "SixMWT_m_first3", "SixMWT_m_last3", "SixMWT_m_tot",
                "SixMWT_km_h_tot", "TwoMWT_m_tot", "TwoMWT_km_h_tot",
                "DWI_6-2", "DWI_6-3", "DWI_6-1"
            ]
            headers.extend(distance_cols)
            ws.append(headers)

            for subj in all_subjects:
                row = []
                # Boundaries times:
                row.extend(subj["boundaries"])
                # Surname, Name, and empty ID:
                row.extend([subj["surname"], subj["name"], ""])
                # Compute distance/speed metrics using the meter boundaries.
                # Expecting exactly 7 meter values: m0, m1, ..., m6
                m = subj["meters"]
                if len(m) < 7:
                    # Fill with Nones if not enough boundaries
                    m = m + [None]*(7-len(m))
                m0, m1, m2, m3, m4, m5, m6 = m[:7]
                # Compute per-minute distances (handle None if any)
                def diff(a, b):
                    return (b - a) if (a is not None and b is not None) else None

                SixMWT_m_1 = diff(m0, m1)
                SixMWT_m_2 = diff(m1, m2)
                SixMWT_m_3 = diff(m2, m3)
                SixMWT_m_4 = diff(m3, m4)
                SixMWT_m_5 = diff(m4, m5)
                SixMWT_m_6 = diff(m5, m6)
                SixMWT_m_first2 = diff(m0, m2)
                SixMWT_m_second2 = diff(m2, m4)
                SixMWT_m_third2 = diff(m4, m6)
                SixMWT_m_first3 = diff(m0, m3)
                SixMWT_m_last3 = diff(m3, m6)
                SixMWT_m_tot = diff(m0, m6)
                # Speed calculations:
                # Total 6 minutes: time = 6/60 = 0.1 hour, so km/h = (distance in km)/0.1 = distance/1000/0.1 = distance/100
                SixMWT_km_h_tot = (SixMWT_m_tot / 100) if SixMWT_m_tot is not None else None
                # Two minutes: time = 2/60 hour, so km/h = (distance/1000) / (2/60) = (distance*30)/1000
                TwoMWT_m_tot = SixMWT_m_first2  # distance from m0 to m2
                TwoMWT_km_h_tot = (TwoMWT_m_tot * 30 / 1000) if TwoMWT_m_tot is not None else None

                # Percentage differences (if denominators are nonzero)
                def perc_diff(x, y):
                    # returns ((y - x) / x)*100, comparing minute y to minute x
                    if x not in (None, 0) and y is not None:
                        return ((y - x) / x) * 100
                    else:
                        return None

                DWI_6_2 = perc_diff(SixMWT_m_2, SixMWT_m_6)
                DWI_6_3 = perc_diff(SixMWT_m_3, SixMWT_m_6)
                DWI_6_1 = perc_diff(SixMWT_m_1, SixMWT_m_6)

                computed = [
                    SixMWT_m_1, SixMWT_m_2, SixMWT_m_3, SixMWT_m_4, SixMWT_m_5, SixMWT_m_6,
                    SixMWT_m_first2, SixMWT_m_second2, SixMWT_m_third2,
                    SixMWT_m_first3, SixMWT_m_last3, SixMWT_m_tot,
                    SixMWT_km_h_tot, TwoMWT_m_tot, TwoMWT_km_h_tot,
                    DWI_6_2, DWI_6_3, DWI_6_1
                ]
                row.extend(computed)
                ws.append(row)

            # Save workbook to a BytesIO stream
            excel_io = BytesIO()
            wb.save(excel_io)
            excel_io.seek(0)

            st.download_button(label="Download Excel File", data=excel_io, file_name="martinez.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
