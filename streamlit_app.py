
# import streamlit as st
# import pandas as pd
# from openpyxl import load_workbook
# from io import BytesIO
# import re

# # ================================
# # CORE LOGIC - ROBUST AND SIMPLE
# # ================================

# KNOWN_PLANTS = {
#     "Bielsko Biala", "Birmingham", "Blatna", "Einbeck", "Forsheda",
#     "Olofstrom", "Rotenburg", "Celaya", "Dickson", "Goshen",
#     "Kalamazoo", "Saltillo", "Valley City", "Wellington"
# }

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
#                     if cell and isinstance(cell, str) and len(cell.strip()) > 1:
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



# def detect_categories(sheet):
#     categories = []
#     category_map = {
#         'S': 'Sales Price', 'M': 'Material', 'I': 'Investment',
#         'T': 'Tooling', 'C': 'Cycle Times', 'H': 'Headcount'
#     }

#     try:
#         for col in range(1, min(4, sheet.max_column + 1)):
#             for row in range(1, min(sheet.max_row + 1, 50)):
#                 try:
#                     val = sheet.cell(row=row, column=col).value
#                     if not val:
#                         continue
#                     text = str(val).strip()
#                     lines = text.split('\n') if '\n' in text else [text]
#                     for line in lines:
#                         line_clean = line.strip().upper()
#                         if len(line_clean) == 1 and line_clean in category_map:
#                             if not any(c['letter'] == line_clean for c in categories):
#                                 categories.append({
#                                     'row': row, 'column': col,
#                                     'letter': line_clean,
#                                     'name': category_map[line_clean]
#                                 })
#                             break
#                 except:
#                     continue
#         categories.sort(key=lambda x: x['row'])
#     except Exception as e:
#         st.error(f"Error detecting categories: {e}")
#         categories = []

#     return categories

# def detect_plant(sheet):
#     for row in range(1, sheet.max_row + 1):
#         for col in range(1, sheet.max_column + 1):
#             val = sheet.cell(row=row, column=col).value
#             if val and isinstance(val, str):
#                 text = val.strip()
#                 for plant in KNOWN_PLANTS:
#                     if plant.lower() in text.lower():
#                         return plant
#     return None

# def find_subcategory_column(sheet, categories):
#     try:
#         if not categories:
#             return 3
#         category_col = categories[0]['column']
#         candidates = [category_col + 1, category_col + 2, 3, 2]
#         best_col = category_col + 1
#         max_text_cells = 0

#         for col in candidates:
#             if col < 1 or col > sheet.max_column:
#                 continue
#             text_cells = 0
#             for row in range(1, min(30, sheet.max_row + 1)):
#                 cell = sheet.cell(row=row, column=col).value
#                 if cell and isinstance(cell, str) and len(cell.strip()) >= 2:
#                     text_cells += 1
#             if text_cells > max_text_cells:
#                 max_text_cells = text_cells
#                 best_col = col
#         return best_col
#     except:
#         return 3

# def extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name=None):
#     extracted = []
#     if not categories:
#         st.warning("No categories found")
#         return []
#     for i in range(len(categories)):
#         current = categories[i]
#         start_row = current['row']
#         end_row = categories[i + 1]['row'] - 1 if i + 1 < len(categories) else min(start_row + 25, sheet.max_row)
#         for row in range(start_row, end_row + 1):
#             subcat_cell = sheet.cell(row=row, column=subcategory_col).value
#             if not subcat_cell:
#                 continue
#             subcat = str(subcat_cell).strip()
#             if len(subcat) < 2:
#                 continue
#             for col in metric_cols:
#                 val = sheet.cell(row=row, column=col).value
#                 if isinstance(val, (int, float)) and val is not None:
#                     header = headers.get(col, f"Column_{chr(64 + col)}")
#                     if isinstance(header, str) and '\n' in header:
#                         header = header.split('\n')[0]
#                     header = str(header)[:30]
#                     entry = {
#                         'Category': current['name'],
#                         'Subcategory': subcat,
#                         'Metric': header,
#                         'Value': float(val)
#                     }
#                     if plant_name:
#                         entry['Plant'] = plant_name
#                     extracted.append(entry)
#     return extracted

# # ================================
# # STREAMLIT APP
# # ================================

# st.title("📊 SMITCH Excel Extractor")
# st.write("Upload SMITCH Excel files to extract structured data")

# uploaded_files = st.file_uploader("Choose Excel files", type=["xlsm", "xlsx"], accept_multiple_files=True)

# if uploaded_files:
#     st.write(f"Processing {len(uploaded_files)} file(s)...")
#     for file in uploaded_files:
#         st.subheader(f"📂 {file.name}")
#         try:
#             wb = load_workbook(file, data_only=True)
#             ws = wb.active
#             st.write(f"File loaded: {ws.max_row} rows × {ws.max_column} columns")
#             with st.spinner("Detecting file structure..."):
#                 metric_columns, headers, stop_column_found = detect_metric_columns(ws)
#                 category_rows = detect_categories(ws)
#                 subcategory_col = find_subcategory_column(ws, category_rows)
#                 plant_name = detect_plant(ws)

#             col1, col2, col3, col4 = st.columns(4)
#             with col1:
#                 st.metric("Categories", len(category_rows))
#             with col2:
#                 st.metric("Metric Columns", len(metric_columns))
#             with col3:
#                 st.metric("Subcategory Col", chr(64 + subcategory_col))
#             with col4:
#                 st.metric("Stop Column", stop_column_found.title() if stop_column_found else "Auto-detected")

#             with st.spinner("Extracting data..."):
#                 data = extract_smitch_data(ws, category_rows, metric_columns, headers, subcategory_col, plant_name)

#             if data:
#                 df = pd.DataFrame(data)
#                 st.success(f"Extracted {len(df)} records")

#                 st.write("**Categories found:**")
#                 for cat, count in df['Category'].value_counts().items():
#                     st.write(f"• {cat}: {count} records")

#                 st.write("**Data preview:**")
#                 st.dataframe(df.head(10))

#                 output = BytesIO()
#                 with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
#                     df.to_excel(writer, index=False, sheet_name='Extracted')
#                 output.seek(0)

#                 st.download_button(
#                     label=" Download Excel",
#                     data=output,
#                     file_name=f"{file.name.split('.')[0]}_extracted.xlsx",
#                     mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
#                 )
#             else:
#                 st.warning(" No data extracted from this file")
#         except Exception as e:
#             st.error(f" Failed to process {file.name}")
#             st.error(f"Error: {str(e)}")
# else:
#     st.info("👆 Upload Excel files to get started")

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re
from datetime import datetime

# ================================
# CONSTANTS
# ================================

KNOWN_PLANTS = {
    "Bielsko Biala", "Birmingham", "Blatna", "Einbeck", "Forsheda",
    "Olofstrom", "Rotenburg", "Celaya", "Dickson", "Goshen",
    "Kalamazoo", "Saltillo", "Valley City", "Wellington"
}

# ================================
# DETECTION FUNCTIONS
# ================================

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
            for col in range(3, min(sheet.max_column + 1, 30)):
                try:
                    cell = sheet.cell(row=search_row, column=col).value
                    if cell and isinstance(cell, str) and len(cell.strip()) > 1:
                        header_clean = ' '.join(str(cell).split())
                        temp_headers[col] = header_clean
                        temp_cols.append(col)
                        if any(stop.lower() in header_clean.lower() for stop in stop_at_keywords):
                            break
                except:
                    continue
            if len(temp_headers) > len(headers):
                headers = temp_headers
                metric_cols = temp_cols
    except:
        metric_cols = [3, 4, 5, 6]
        headers = {col: f"Column_{chr(64 + col)}" for col in metric_cols}
    return metric_cols, headers, None

def detect_categories(sheet):
    categories = []
    category_map = {
        'S': 'Sales Price', 'M': 'Material', 'I': 'Investment',
        'T': 'Tooling', 'C': 'Cycle Times', 'H': 'Headcount'
    }
    for col in range(1, min(4, sheet.max_column + 1)):
        for row in range(1, min(sheet.max_row + 1, 50)):
            val = sheet.cell(row=row, column=col).value
            if val:
                line = str(val).strip().upper()
                if len(line) == 1 and line in category_map:
                    if not any(c['letter'] == line for c in categories):
                        categories.append({
                            'row': row, 'column': col, 'letter': line, 'name': category_map[line]
                        })
    categories.sort(key=lambda x: x['row'])
    return categories

def detect_plant(sheet):
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            val = sheet.cell(row=row, column=col).value
            if val and isinstance(val, str):
                for plant in KNOWN_PLANTS:
                    if plant.lower() in val.lower():
                        return plant
    return None

def detect_part_name(sheet, plant_row, first_category_row):
    for row in range(plant_row + 1, first_category_row):
        val = sheet.cell(row=row, column=2).value  # Column B
        if val and isinstance(val, str):
            if re.search(r"[A-Za-z]+\s*\d", val):  # e.g., "Bracket 123A"
                return val.strip()
    return None

def find_subcategory_column(sheet, categories):
    category_col = categories[0]['column'] if categories else 3
    candidates = [category_col + 1, category_col + 2, 3, 2]
    best_col = category_col + 1
    max_text_cells = 0
    for col in candidates:
        text_cells = 0
        for row in range(1, min(30, sheet.max_row + 1)):
            val = sheet.cell(row=row, column=col).value
            if val and isinstance(val, str) and len(val.strip()) > 1:
                text_cells += 1
        if text_cells > max_text_cells:
            best_col = col
            max_text_cells = text_cells
    return best_col

def extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name=None, part_name=None):
    extracted = []
    for i in range(len(categories)):
        current = categories[i]
        start_row = current['row']
        end_row = categories[i + 1]['row'] - 1 if i + 1 < len(categories) else sheet.max_row
        for row in range(start_row, end_row + 1):
            subcat_cell = sheet.cell(row=row, column=subcategory_col).value
            if not subcat_cell or len(str(subcat_cell).strip()) < 2:
                continue
            subcat = str(subcat_cell).strip()

            # Try to extract a date (e.g., "Actual 12/2025")
            date_match = re.search(r'(\d{1,2}/\d{2,4})', subcat)
            try:
                date_val = datetime.strptime(date_match.group(1), "%m/%Y").date() if date_match else None
            except:
                try:
                    date_val = datetime.strptime(date_match.group(1), "%m/%y").date() if date_match else None
                except:
                    date_val = None

            for col in metric_cols:
                val = sheet.cell(row=row, column=col).value
                if isinstance(val, (int, float)) and val is not None:
                    header = headers.get(col, f"Column_{chr(64 + col)}")
                    if '\n' in header:
                        header = header.split('\n')[0]
                    record = {
                        "Category": current['name'],
                        "Subcategory": subcat,
                        "Metric": header,
                        "Value": float(val)
                    }
                    if plant_name:
                        record["Plant"] = plant_name
                    if part_name:
                        record["Part Name"] = part_name
                    if date_val:
                        record["Date"] = date_val
                    extracted.append(record)
    return extracted

# ================================
# STREAMLIT APP
# ================================

st.title("📊 SMITCH Excel Extractor")
st.write("Upload SMITCH Excel files to extract structured data")

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsm", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        st.subheader(f"📂 {file.name}")
        try:
            wb = load_workbook(file, data_only=True)
            ws = wb.active
            with st.spinner("Analyzing file..."):
                metric_columns, headers, _ = detect_metric_columns(ws)
                categories = detect_categories(ws)
                subcat_col = find_subcategory_column(ws, categories)
                plant_name = detect_plant(ws)
                plant_row = next((r for r in range(1, ws.max_row) for c in range(1, ws.max_column)
                                 if ws.cell(row=r, column=c).value and
                                 isinstance(ws.cell(row=r, column=c).value, str) and
                                 any(p.lower() in ws.cell(row=r, column=c).value.lower() for p in KNOWN_PLANTS)), None)
                first_category_row = categories[0]['row'] if categories else ws.max_row
                part_name = detect_part_name(ws, plant_row or 1, first_category_row)
                data = extract_smitch_data(ws, categories, metric_columns, headers, subcat_col, plant_name, part_name)

            if data:
                df = pd.DataFrame(data)
                st.success(f"Extracted {len(df)} records from {file.name}")
                st.dataframe(df.head(10))
                output = BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False, sheet_name="Extracted")
                st.download_button(
                    label="📥 Download Excel",
                    data=output.getvalue(),
                    file_name=f"{file.name.split('.')[0]}_extracted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No data extracted from this file.")

        except Exception as e:
            st.error(f"❌ Failed to process {file.name}")
            st.exception(e)
else:
    st.info("👆 Upload one or more SMITCH Excel files to begin.")



                       



   
