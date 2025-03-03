import os
import re
import xml.etree.ElementTree as ET
import openpyxl
import streamlit as st
import io

def process_file(file_path_or_buffer):
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
    # Parse the XML (file uploader returns a BytesIO object)
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
    # We expect to find a "START" marker (assigned 0.0 m) and then six rows with numeric markers.
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
                    except Exception:
                        continue
            if len(boundaries) == 7:
                break
    if len(boundaries) < 7:
        raise ValueError("Not enough boundaries found (expected START plus six numeric markers).")

    # --- Compute per-minute averages for physiological metrics ---
    # It is assumed that immediately after the Marker column, there are 17 metrics:
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
    # For each variable, we compute aggregates as:
    # "third1": average of minutes 1-2,
    # "third2": average of minutes 3-4,
    # "third3": average of minutes 5-6,
    # "half1": average of minutes 1-3,
    # "half2": average of minutes 4-6,
    # "total": average of minutes 1-6.
    variables = {}
    var_names = ["V'O2", "V'O2/kg", "V'E", "RER", "FC", "CHO", "FAT",
                 "EE", "EECHO", "EEFAT", "METS", "V'CO2", "V'E/V'CO2", "BF", "EE/BSA", "EE/kg", "EE/kg/magra"]
    for idx, var in enumerate(var_names):
        var_data = {}
        for m in range(1, 7):
            var_data[f"{var}_{m}"] = per_minute_avgs[m-1][idx]
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

            # Build header for fixed columns:
            headers = [f"Edge_{i}" for i in range(1, 8)]
            headers.extend(["Surname", "Name", "ID"])
            # Distance and speed columns:
            dist_cols = [
                "SixMWT_m_1", "SixMWT_m_2", "SixMWT_m_3", "SixMWT_m_4", "SixMWT_m_5", "SixMWT_m_6",
                "SixMWT_m_first2", "SixMWT_m_second2", "SixMWT_m_third2",
                "SixMWT_m_first3", "SixMWT_m_last3", "SixMWT_m_tot",
                "SixMWT_km_h_tot", "TwoMWT_m_tot", "TwoMWT_km_h_tot",
                "DWI_6-2", "DWI_6-3", "DWI_6-1"
            ]
            headers.extend(dist_cols)
            # Now, build headers for physiological aggregates in the hardcoded order.
            # Define groups and corresponding aggregate keys:
            groups = [
                ("SixMWT_First3", "half1"),
                ("SixMWT_Last3", "half2"),
                ("SIXmwt_Tot", "total"),
                ("SIXmwt_First2", "third1"),
                ("SIXmwt_Second2", "third2"),
                ("SIXmwt_Last2", "third3")
            ]
            # For each group, output 18 columns in this order:
            # V'O2, V'O2/kg, V'E, RER, FC, CHO, FAT, EE, EECHO, EEFAT, METS, V'CO2, V'E/V'CO2, V'E/V'O2, BF, EE/BSA, EE/kg, EE/kg_magra
            phys_vars_order = ["V'O2", "V'O2/kg", "V'E", "RER", "FC", "CHO", "FAT",
                                "EE", "EECHO", "EEFAT", "METS", "V'CO2", "V'E/V'CO2", "V'E/V'O2", "BF", "EE/BSA", "EE/kg", "EE/kg_magra"]
            for grp, _ in groups:
                for var in phys_vars_order:
                    headers.append(f"{grp}_{var}")
            ws.append(headers)

            # For each subject, build the row:
            for subj in all_subjects:
                row = []
                row.extend(subj["boundaries"])
                row.extend([subj["surname"], subj["name"], ""])
                # --- Calculate 6MWT distance and speed values from boundaries ---
                m = subj["meters"]
                if len(m) == 7:
                    m0, m1, m2, m3, m4, m5, m6 = m
                    SixMWT_m_1 = m1 - m0
                    SixMWT_m_2 = m2 - m1
                    SixMWT_m_3 = m3 - m2
                    SixMWT_m_4 = m4 - m3
                    SixMWT_m_5 = m5 - m4
                    SixMWT_m_6 = m6 - m5
                    SixMWT_m_first2 = m2 - m0
                    SixMWT_m_second2 = m4 - m2
                    SixMWT_m_third2 = m6 - m4
                    SixMWT_m_first3 = m3 - m0
                    SixMWT_m_last3 = m6 - m3
                    SixMWT_m_tot = m6 - m0
                    SixMWT_km_h_tot = SixMWT_m_tot / 100 if SixMWT_m_tot is not None else None
                    TwoMWT_m_tot = SixMWT_m_first2
                    TwoMWT_km_h_tot = (TwoMWT_m_tot * 30 / 1000) if TwoMWT_m_tot is not None else None
                    DWI_6_2 = ((SixMWT_m_6 - SixMWT_m_2) / SixMWT_m_2 * 100) if SixMWT_m_2 not in (None, 0) else None
                    DWI_6_3 = ((SixMWT_m_6 - SixMWT_m_3) / SixMWT_m_3 * 100) if SixMWT_m_3 not in (None, 0) else None
                    DWI_6_1 = ((SixMWT_m_6 - SixMWT_m_1) / SixMWT_m_1 * 100) if SixMWT_m_1 not in (None, 0) else None
                else:
                    SixMWT_m_1 = SixMWT_m_2 = SixMWT_m_3 = SixMWT_m_4 = SixMWT_m_5 = SixMWT_m_6 = None
                    SixMWT_m_first2 = SixMWT_m_second2 = SixMWT_m_third2 = None
                    SixMWT_m_first3 = SixMWT_m_last3 = SixMWT_m_tot = None
                    SixMWT_km_h_tot = TwoMWT_m_tot = TwoMWT_km_h_tot = None
                    DWI_6_2 = DWI_6_3 = DWI_6_1 = None

                computed = [
                    SixMWT_m_1, SixMWT_m_2, SixMWT_m_3, SixMWT_m_4, SixMWT_m_5, SixMWT_m_6,
                    SixMWT_m_first2, SixMWT_m_second2, SixMWT_m_third2,
                    SixMWT_m_first3, SixMWT_m_last3, SixMWT_m_tot,
                    SixMWT_km_h_tot, TwoMWT_m_tot, TwoMWT_km_h_tot,
                    DWI_6_2, DWI_6_3, DWI_6_1
                ]
                row.extend(computed)

                # --- For physiological variables, use the new groups order ---
                # Our aggregates are stored in subj["variables"][var] under keys:
                # "third1", "third2", "third3", "half1", "half2", "total"
                phys_vars = ["V'O2", "V'O2/kg", "V'E", "RER", "FC", "CHO", "FAT",
                             "EE", "EECHO", "EEFAT", "METS", "V'CO2", "V'E/V'CO2", "BF", "EE/BSA", "EE/kg", "EE/kg/magra"]
                # Mapping groups to the aggregate keys:
                group_key = {
                    "SixMWT_First3": "half1",
                    "SixMWT_Last3": "half2",
                    "SIXmwt_Tot": "total",
                    "SIXmwt_First2": "third1",
                    "SIXmwt_Second2": "third2",
                    "SIXmwt_Last2": "third3"
                }
                for grp, agg_key in groups:
                    for var in phys_vars:
                        # For all variables except V'E/V'O2, value comes directly:
                        if var != "V'E/V'O2":
                            val = subj["variables"].get(var, {}).get(agg_key)
                        else:
                            # For V'E/V'O2, compute aggregated V'E / aggregated V'O2
                            v_e = subj["variables"].get("V'E", {}).get(agg_key)
                            v_o2 = subj["variables"].get("V'O2", {}).get(agg_key)
                            try:
                                val = v_e / v_o2 if (v_o2 not in (None, 0)) else None
                            except Exception:
                                val = None
                        # However, note that our original var_names list did not include V'E/V'O2.
                        # To allow this column, we compute it on the fly.
                        # Since the requested order includes V'E/V'O2 after V'E/V'CO2 and before BF,
                        # we insert it manually.
                        # So, we loop over the fixed order below.
                        # Here we instead create an ordered list:
                        pass  # We'll build the ordered header and then assign values accordingly.
                    # Instead, we’ll loop over a fixed order list defined below.
                # Build the ordered list for each group:
                # Ordered variables per group:
                order = ["V'O2", "V'O2/kg", "V'E", "RER", "FC", "CHO", "FAT",
                         "EE", "EECHO", "EEFAT", "METS", "V'CO2", "V'E/V'CO2", "V'E/V'O2", "BF", "EE/BSA", "EE/kg", "EE/kg/magra"]
                for grp, agg_key in groups:
                    for var in order:
                        if var == "V'E/V'O2":
                            v_e = subj["variables"].get("V'E", {}).get(agg_key)
                            v_o2 = subj["variables"].get("V'O2", {}).get(agg_key)
                            try:
                                comp_val = v_e / v_o2 if (v_o2 not in (None, 0)) else None
                            except Exception:
                                comp_val = None
                        else:
                            comp_val = subj["variables"].get(var, {}).get(agg_key)
                        row.append(comp_val)
                ws.append(row)

            excel_io = BytesIO()
            wb.save(excel_io)
            excel_io.seek(0)
            st.download_button(label="Download Excel File", data=excel_io, file_name=excel_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
