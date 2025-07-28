# import streamlit as st
# import pandas as pd
# from openpyxl import load_workbook
# from io import BytesIO
# from datetime import datetime
# import re

# # Known plant names to identify plant from sheet
# KNOWN_PLANTS = {
#     "Bielsko Biala", "Birmingham", "Blatna", "Einbeck", "Forsheda",
#     "Olofstrom", "Rotenburg", "Celaya", "Dickson", "Goshen",
#     "Kalamazoo", "Saltillo", "Valley City", "Wellington"
# }


# def detect_metric_columns(sheet):
#     metric_cols = []
#     headers = {}

#     for search_row in range(1, min(6, sheet.max_row + 1)):
#         for col in range(3, min(sheet.max_column + 1, 20)):
#             cell = sheet.cell(row=search_row, column=col).value
#             if cell and isinstance(cell, str) and len(cell.strip()) > 1:
#                 header_clean = ' '.join(str(cell).strip().split()).lower()
#                 headers[col] = header_clean
#                 metric_cols.append(col)

#     if not metric_cols:
#         metric_cols = [3, 4, 5, 6]
#         headers = {3: "column_c", 4: "column_d", 5: "column_e", 6: "column_f"}

#     return metric_cols, headers, None


# def detect_categories(sheet):
#     categories = []
#     category_map = {
#         'S': 'Sales Price', 'M': 'Material', 'I': 'Investment',
#         'T': 'Tooling', 'C': 'Cycle Times', 'H': 'Headcount'
#     }

#     for col in range(1, min(4, sheet.max_column + 1)):
#         for row in range(1, min(sheet.max_row + 1, 50)):
#             val = sheet.cell(row=row, column=col).value
#             if not val:
#                 continue
#             text = str(val).strip().upper()
#             if len(text) == 1 and text in category_map:
#                 if not any(c['letter'] == text for c in categories):
#                     categories.append({
#                         'row': row, 'column': col,
#                         'letter': text, 'name': category_map[text]
#                     })
#     categories.sort(key=lambda x: x['row'])
#     return categories


# def find_subcategory_column(sheet, categories):
#     if not categories:
#         return 3
#     category_col = categories[0]['column']
#     candidates = [category_col + 1, category_col + 2, 3, 2]
#     best_col = category_col + 1
#     max_text_cells = 0

#     for col in candidates:
#         text_cells = 0
#         for row in range(1, min(30, sheet.max_row + 1)):
#             cell = sheet.cell(row=row, column=col).value
#             if cell and isinstance(cell, str) and len(cell.strip()) >= 2:
#                 text_cells += 1
#         if text_cells > max_text_cells:
#             max_text_cells = text_cells
#             best_col = col
#     return best_col


# def detect_plant(sheet):
#     for row in range(1, sheet.max_row + 1):
#         for col in range(1, sheet.max_column + 1):
#             val = sheet.cell(row=row, column=col).value
#             if val and isinstance(val, str):
#                 if any(plant.lower() in val.lower() for plant in KNOWN_PLANTS):
#                     return val.strip(), row
#     return None, None


# def detect_part_name(sheet, categories):
#     if not categories:
#         return None
#     first_row = categories[0]['row']
#     for row in range(first_row - 1, 0, -1):
#         val = sheet.cell(row=row, column=2).value
#         if val and isinstance(val, str) and len(val.strip()) > 3:
#             return val.strip()
#     return None


# def extract_date(text):
#     if not isinstance(text, str):
#         return None
#     matches = re.findall(r"\b\d{1,2}[/-]\d{2,4}\b", text)
#     for match in matches:
#         try:
#             if re.match(r"\d{1,2}/\d{4}$", match):
#                 return datetime.strptime(match, "%m/%Y").strftime("%Y-%m-%d")
#             elif re.match(r"\d{1,2}/\d{2}$", match):
#                 return datetime.strptime(match, "%m/%y").strftime("%Y-%m-%d")
#         except:
#             continue
#     return None


# def extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name=None, part_name=None):
#     extracted = []

#     if not categories:
#         st.warning("No categories found")
#         return []

#     col_date_map = {}
#     for col in metric_cols:
#         for row in range(1, 6):
#             cell_val = sheet.cell(row=row, column=col).value
#             if isinstance(cell_val, str):
#                 date_found = extract_date(cell_val)
#                 if date_found:
#                     col_date_map[col] = date_found
#                     break

#     for i in range(len(categories)):
#         current = categories[i]
#         start_row = current['row']
#         end_row = categories[i + 1]['row'] - 1 if i + 1 < len(categories) else sheet.max_row

#         for row in range(start_row, end_row + 1):
#             subcat_cell = sheet.cell(row=row, column=subcategory_col).value
#             if not subcat_cell:
#                 continue
#             subcat = str(subcat_cell).strip()

#             for col in metric_cols:
#                 val = sheet.cell(row=row, column=col).value
#                 if not isinstance(val, (int, float)):
#                     continue

#                 raw_header = headers.get(col, f"Column_{chr(64 + col)}").strip().lower()
#                 if "cm%" in raw_header:
#                     continue

#                 # Normalize metric to first word (e.g., "Actual Performance" → "Actual")
#                 metric = raw_header.split()[0].capitalize()

#                 date_str = col_date_map.get(col)

#                 entry = {
#                     'Category': current['name'],
#                     'Subcategory': subcat,
#                     'Date': date_str,
#                     'Metric': metric,
#                     'Value': float(val)
#                 }
#                 if plant_name:
#                     entry['Plant'] = plant_name
#                 if part_name:
#                     entry['Part Name'] = part_name

#                 extracted.append(entry)

#     return extracted


# # Streamlit UI
# st.title("📊 SMITCH Excel Extractor")
# uploaded_files = st.file_uploader("Upload Excel files", type=["xlsm", "xlsx"], accept_multiple_files=True)

# if uploaded_files:
#     for file in uploaded_files:
#         st.subheader(file.name)
#         wb = load_workbook(file, data_only=True)
#         sheet = wb.active

#         metric_cols, headers, _ = detect_metric_columns(sheet)
#         categories = detect_categories(sheet)
#         subcategory_col = find_subcategory_column(sheet, categories)
#         plant_name, _ = detect_plant(sheet)
#         part_name = detect_part_name(sheet, categories)

#         df = pd.DataFrame(extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name, part_name))

#         if not df.empty:
#             st.success(f" Extracted {len(df)} rows")
#             st.dataframe(df.head(10))

#             buffer = BytesIO()
#             with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
#                 df.to_excel(writer, index=False)
#             buffer.seek(0)

#             st.download_button(
#                 label=" Download Extracted Excel",
#                 data=buffer,
#                 file_name=f"{file.name.split('.')[0]}_extracted.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )
#         else:
#             st.warning("No valid data extracted.")
# else:
#     st.info(" Upload Excel files to begin.")

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime
import re
from dateutil import parser


KNOWN_PLANTS = {
    "Bielsko Biala", "Birmingham", "Blatna", "Einbeck", "Forsheda",
    "Olofstrom", "Rotenburg", "Celaya", "Dickson", "Goshen",
    "Kalamazoo", "Saltillo", "Valley City", "Wellington"
}

# def detect_metric_columns(sheet, stop_at_keywords=None):
#     if stop_at_keywords is None:
#         stop_at_keywords = [
#             "demon-strated rate at 100%", "demonstrated rate at 100%",
#             "demon-strated rate", "demonstrated rate",
#         ]

#     metric_cols = []
#     headers = {}
#     stop_column_found = None

#     try:
#         for search_row in range(1, min(6, sheet.max_row + 1)):
#             temp_cols = []
#             temp_headers = {}
#             temp_stop_col = None

#             for col in range(3, min(sheet.max_column + 1, 20)):
#                 try:
#                     cell = sheet.cell(row=search_row, column=col).value
#                     if cell:
#                         header_clean = ' '.join(str(cell).split())
#                         temp_headers[col] = header_clean
#                         temp_cols.append(col)

#                         header_lower = header_clean.lower()
#                         for stop_keyword in stop_at_keywords:
#                             if stop_keyword in header_lower:
#                                 temp_stop_col = stop_keyword
#                                 break

#                         if temp_stop_col:
#                             break
#                 except:
#                     continue

#             if temp_stop_col or len(temp_headers) > len(headers):
#                 headers = temp_headers
#                 metric_cols = temp_cols
#                 if temp_stop_col:
#                     stop_column_found = temp_stop_col
#                     break

#         if not metric_cols:
#             metric_cols = list(range(3, min(8, sheet.max_column + 1)))
#             for col in metric_cols:
#                 headers[col] = f"Column_{chr(64 + col)}"

#     except:
#         metric_cols = [3, 4, 5, 6]
#         headers = {3: "Column_C", 4: "Column_D", 5: "Column_E", 6: "Column_F"}

#     return metric_cols, headers, stop_column_found
def detect_metric_columns(sheet, stop_at_keywords=None):
    if stop_at_keywords is None:
        stop_at_keywords = [
            "demon-strated rate at 100%", "demonstrated rate at 100%",
            "demon-strated rate", "demonstrated rate",
        ]

    metric_cols = []
    headers = {}
    stop_column_found = None

    try:
        for search_row in range(1, min(6, sheet.max_row + 1)):
            temp_cols = []
            temp_headers = {}
            temp_stop_col = None

            for col in range(3, min(sheet.max_column + 1, 20)):
                try:
                    cell = sheet.cell(row=search_row, column=col).value
                    if cell and isinstance(cell, str) and len(cell.strip()) > 1:
                        header_clean = ' '.join(str(cell).strip().split()).lower()
                        temp_headers[col] = header_clean
                        temp_cols.append(col)

                        for stop_keyword in stop_at_keywords:
                            if stop_keyword in header_clean:
                                temp_stop_col = stop_keyword
                                break
                        if temp_stop_col:
                            break
                except:
                    continue

            if temp_stop_col or len(temp_headers) > len(headers):
                headers = temp_headers
                metric_cols = temp_cols
                if temp_stop_col:
                    stop_column_found = temp_stop_col
                    break

        if not metric_cols:
            metric_cols = list(range(3, min(8, sheet.max_column + 1)))
            for col in metric_cols:
                headers[col] = f"column_{chr(64 + col)}".lower()

    except:
        metric_cols = [3, 4, 5, 6]
        headers = {3: "column_c", 4: "column_d", 5: "column_e", 6: "column_f"}

    return metric_cols, headers, stop_column_found


def detect_categories(sheet):
    categories = []
    category_map = {
        'S': 'Sales Price', 'M': 'Material', 'I': 'Investment',
        'T': 'Tooling', 'C': 'Cycle Times', 'H': 'Headcount'
    }

    try:
        for col in range(1, min(4, sheet.max_column + 1)):
            for row in range(1, min(sheet.max_row + 1, 50)):
                try:
                    val = sheet.cell(row=row, column=col).value
                    if not val:
                        continue
                    text = str(val).strip()
                    lines = text.split('\n') if '\n' in text else [text]
                    for line in lines:
                        line_clean = line.strip().upper()
                        if len(line_clean) == 1 and line_clean in category_map:
                            if not any(c['letter'] == line_clean for c in categories):
                                categories.append({
                                    'row': row, 'column': col,
                                    'letter': line_clean,
                                    'name': category_map[line_clean]
                                })
                            break
                except:
                    continue
        categories.sort(key=lambda x: x['row'])
    except Exception as e:
        st.error(f"Error detecting categories: {e}")
        categories = []

    return categories

def detect_plant(sheet):
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            val = sheet.cell(row=row, column=col).value
            if val and isinstance(val, str):
                text = val.strip()
                for plant in KNOWN_PLANTS:
                    if plant.lower() in text.lower():
                        return plant, row
    return None, None

def detect_part_name(sheet, categories):
    try:
        if not categories:
            return None
        first_category_row = categories[0]['row']
        for row in range(first_category_row - 1, 0, -1):  # search upward
            val = sheet.cell(row=row, column=2).value  
            if val and isinstance(val, str):
                val = val.strip()
                # Return first non-empty text that isn't a single letter
                if len(val) > 3 and val.upper() not in {'S', 'M', 'I', 'T', 'C', 'H'}:
                    return val
    except:
        pass
    return None
      

def extract_date(text):
    if not isinstance(text, str):
        return None

    matches = re.findall(r"\b\d{1,2}[/-]\d{2,4}(?:[/-]\d{2,4})?\b", text)
    for match in matches:
        try:
            if re.match(r"\d{1,2}/\d{4}$", match):  # MM/YYYY
                dt = datetime.strptime(match, "%m/%Y")
            elif re.match(r"\d{1,2}/\d{2}$", match):  # MM/YY
                dt = datetime.strptime(match, "%m/%y")
            elif re.match(r"\d{1,2}/\d{1,2}/\d{4}$", match):  # MM/DD/YYYY
                dt = datetime.strptime(match, "%m/%d/%Y")
            else:
                continue
            return dt.strftime("%Y-%m-%d")
        except:
            continue
    return None



def find_subcategory_column(sheet, categories):
    try:
        if not categories:
            return 3
        category_col = categories[0]['column']
        candidates = [category_col + 1, category_col + 2, 3, 2]
        best_col = category_col + 1
        max_text_cells = 0

        for col in candidates:
            if col < 1 or col > sheet.max_column:
                continue
            text_cells = 0
            for row in range(1, min(30, sheet.max_row + 1)):
                cell = sheet.cell(row=row, column=col).value
                if cell and isinstance(cell, str) and len(cell.strip()) >= 2:
                    text_cells += 1
            if text_cells > max_text_cells:
                max_text_cells = text_cells
                best_col = col
        return best_col
    except:
        return 3

# def extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name=None, part_name=None):
#     extracted = []

#     if not categories:
#         st.warning("No categories found")
#         return []

#     METRIC_NORMALIZATION = {
#         "quoted cost model": "Quoted",
#         "quoted": "Quoted",
#         "plex standard": "Plex",
#         "plex": "Plex",
#         "actual performance": "Actual",
#         "actual": "Actual",
#         "forecasted cost": "Forecasted",
#         "forecasted": "Forecasted",
#         "demonstrated rate": "Demonstrated",
#         "demon-strated": "Demonstrated",
#     }

#     col_date_map = {}
#     for col in metric_cols:
#         date_found = None
#         for row in range(1, 6):
#             cell_val = sheet.cell(row=row, column=col).value
#             if isinstance(cell_val, str):
#                 possible_date = extract_date(cell_val)
#                 if possible_date:
#                     date_found = possible_date
#                     break
#         col_date_map[col] = date_found

#     for i in range(len(categories)):
#         current = categories[i]
#         start_row = current['row']
#         end_row = categories[i + 1]['row'] - 1 if i + 1 < len(categories) else min(start_row + 25, sheet.max_row)

#         for row in range(start_row, end_row + 1):
#             subcat_cell = sheet.cell(row=row, column=subcategory_col).value
#             if not subcat_cell:
#                 continue
#             subcat = str(subcat_cell).strip()

#             for col in metric_cols:
#                 val = sheet.cell(row=row, column=col).value
#                 if not isinstance(val, (int, float)):
#                     continue
#             raw_header = headers.get(col, f"Column_{chr(64 + col)}").strip().lower().split('\n')[0]
#             if "cm%" in raw_header:
#                 continue
#             metric = raw_header.split()[0].capitalize() if raw_header else f"Col_{col}"

#                 # raw_header = headers.get(col, f"Column_{chr(64 + col)}").strip().lower().split('\n')[0]
#                 # if "cm%" in raw_header:
#                 #     continue
#                 # matched_key = next((k for k in METRIC_NORMALIZATION if k in raw_header), None)
 

#                 # if matched_key:
#                 #     metric = METRIC_NORMALIZATION[matched_key]
#                 # else:
#                 #     metric = raw_header.split()[0].capitalize() if raw_header else f"Col_{col}"

#                 date_str = col_date_map.get(col)

#                 entry = {
#                     'Category': current['name'],
#                     'Subcategory': subcat,
#                     'Date': date_str,
#                     'Metric': metric,
#                     'Value': float(val)
#                 }
#                 if plant_name:
#                     entry['Plant'] = plant_name
#                 if part_name:
#                     entry['Part Name'] = part_name
                

#                 extracted.append(entry)

#     return extracted

def extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name=None, part_name=None):
    extracted = []

    if not categories:
        st.warning("No categories found")
        return []

    METRIC_NORMALIZATION = {
        "quoted cost model": "Quoted",
        "quoted": "Quoted",
        "plex standard": "Plex",
        "plex": "Plex",
        "actual performance": "Actual",
        "actual": "Actual",
        "forecasted cost": "Forecasted",
        "forecasted": "Forecasted",
        "demonstrated rate": "Demonstrated",
        "demon-strated": "Demonstrated",
    }

    # Preprocess: Extract date from headers for each metric column
    col_date_map = {}
    for col in metric_cols:
        date_found = None
        for row in range(1, 6):
            cell_val = sheet.cell(row=row, column=col).value
            if isinstance(cell_val, str):
                possible_date = extract_date(cell_val)
                if possible_date:
                    date_found = possible_date
                    break
        col_date_map[col] = date_found

    # Iterate through category rows
    for i in range(len(categories)):
        current = categories[i]
        start_row = current['row']
        end_row = categories[i + 1]['row'] - 1 if i + 1 < len(categories) else min(start_row + 25, sheet.max_row)

        for row in range(start_row, end_row + 1):
            subcat_cell = sheet.cell(row=row, column=subcategory_col).value
            if not subcat_cell:
                continue
            subcat = str(subcat_cell).strip()

            for col in metric_cols:
                val = sheet.cell(row=row, column=col).value
                if not isinstance(val, (int, float)):
                    continue

                # Normalize header
                raw_header = headers.get(col, f"column_{chr(64 + col)}").strip().lower().split('\n')[0]
                if "cm%" in raw_header:
                    continue

                # Clean header for matching
                cleaned_header = re.sub(r'[^a-z\s$\/→-]', '', raw_header.lower()).strip()

                # Match against normalization dict
                matched_key = next((k for k in METRIC_NORMALIZATION if k in cleaned_header), None)

                if matched_key:
                    metric = METRIC_NORMALIZATION[matched_key]
                elif "quoted jph" in cleaned_header:
                    metric = "Quoted_JPH"
                elif "quoted $" in cleaned_header or "quoted $ / piece" in cleaned_header:
                    metric = "Quoted_$"
                elif "actual jph" in cleaned_header:
                    metric = "Actual_JPH"
                elif "actual $" in cleaned_header or "actual $ / piece" in cleaned_header:
                    metric = "Actual_$"
                elif "plex std" in cleaned_header and "jph" in cleaned_header:
                    metric = "Plex_JPH"
                elif "plex std" in cleaned_header and ("$" in cleaned_header or "piece" in cleaned_header):
                    metric = "Plex_$"
                else:
                    metric = raw_header.split()[0].capitalize() if raw_header else f"Col_{col}"

                date_str = col_date_map.get(col)

                entry = {
                    'Category': current['name'],
                    'Subcategory': subcat,
                    'Date': date_str,
                    'Metric': metric,
                    'Value': float(val)
                }
                if plant_name:
                    entry['Plant'] = plant_name
                if part_name:
                    entry['Part Name'] = part_name

                extracted.append(entry)

    return extracted



st.title("📊 SMITCH Excel Extractor")
st.write("Upload SMITCH Excel files to extract structured data")

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsm", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.write(f"Processing {len(uploaded_files)} file(s)...")
    for file in uploaded_files:
        st.subheader(f"{file.name}")
        try:
            wb = load_workbook(file, data_only=True)
            ws = wb.active
            st.write(f"File loaded: {ws.max_row} rows × {ws.max_column} columns")
            with st.spinner("Detecting file structure..."):
                metric_columns, headers, stop_column_found = detect_metric_columns(ws)
                category_rows = detect_categories(ws)
                subcategory_col = find_subcategory_column(ws, category_rows)
                plant_name, plant_row = detect_plant(ws)
                smitch_row = category_rows[0]['row'] if category_rows else ws.max_row
                part_name = detect_part_name(ws, category_rows)

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Categories", len(category_rows))
            with col2:
                st.metric("Metric Columns", len(metric_columns))
            with col3:
                st.metric("Subcategory Col", chr(64 + subcategory_col))
            with col4:
                st.metric("Stop Column", stop_column_found.title() if stop_column_found else "Auto-detected")

            with st.spinner("Extracting data..."):
                data = extract_smitch_data(ws, category_rows, metric_columns, headers, subcategory_col, plant_name, part_name)

            if data:
                df = pd.DataFrame(data)
                st.success(f"Extracted {len(df)} records")

                st.write("**Categories found:**")
                for cat, count in df['Category'].value_counts().items():
                    st.write(f"• {cat}: {count} records")

                st.write("**Data preview:**")
                st.dataframe(df.head(10))

                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Extracted')
                output.seek(0)

                st.download_button(
                    label=" Download Excel",
                    data=output,
                    file_name=f"{file.name.split('.')[0]}_extracted.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.warning(" No data extracted from this file")
        except Exception as e:
            st.error(f" Failed to process {file.name}")
            st.error(f"Error: {str(e)}")
else:
    st.info(" Upload Excel files to get started")


