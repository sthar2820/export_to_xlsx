import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime
import re

# Known plant names to identify plant from sheet
KNOWN_PLANTS = {
    "Bielsko Biala", "Birmingham", "Blatna", "Einbeck", "Forsheda",
    "Olofstrom", "Rotenburg", "Celaya", "Dickson", "Goshen",
    "Kalamazoo", "Saltillo", "Valley City", "Wellington"
}


def detect_metric_columns(sheet):
    metric_cols = []
    headers = {}

    for search_row in range(1, min(6, sheet.max_row + 1)):
        for col in range(3, min(sheet.max_column + 1, 20)):
            cell = sheet.cell(row=search_row, column=col).value
            if cell and isinstance(cell, str) and len(cell.strip()) > 1:
                header_clean = ' '.join(str(cell).strip().split()).lower()
                headers[col] = header_clean
                metric_cols.append(col)

    if not metric_cols:
        metric_cols = [3, 4, 5, 6]
        headers = {3: "column_c", 4: "column_d", 5: "column_e", 6: "column_f"}

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
            if not val:
                continue
            text = str(val).strip().upper()
            if len(text) == 1 and text in category_map:
                if not any(c['letter'] == text for c in categories):
                    categories.append({
                        'row': row, 'column': col,
                        'letter': text, 'name': category_map[text]
                    })
    categories.sort(key=lambda x: x['row'])
    return categories


def find_subcategory_column(sheet, categories):
    if not categories:
        return 3
    category_col = categories[0]['column']
    candidates = [category_col + 1, category_col + 2, 3, 2]
    best_col = category_col + 1
    max_text_cells = 0

    for col in candidates:
        text_cells = 0
        for row in range(1, min(30, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=col).value
            if cell and isinstance(cell, str) and len(cell.strip()) >= 2:
                text_cells += 1
        if text_cells > max_text_cells:
            max_text_cells = text_cells
            best_col = col
    return best_col


def detect_plant(sheet):
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            val = sheet.cell(row=row, column=col).value
            if val and isinstance(val, str):
                if any(plant.lower() in val.lower() for plant in KNOWN_PLANTS):
                    return val.strip(), row
    return None, None


def detect_part_name(sheet, categories):
    if not categories:
        return None
    first_row = categories[0]['row']
    for row in range(first_row - 1, 0, -1):
        val = sheet.cell(row=row, column=2).value
        if val and isinstance(val, str) and len(val.strip()) > 3:
            return val.strip()
    return None


def extract_date(text):
    if not isinstance(text, str):
        return None
    matches = re.findall(r"\b\d{1,2}[/-]\d{2,4}\b", text)
    for match in matches:
        try:
            if re.match(r"\d{1,2}/\d{4}$", match):
                return datetime.strptime(match, "%m/%Y").strftime("%Y-%m-%d")
            elif re.match(r"\d{1,2}/\d{2}$", match):
                return datetime.strptime(match, "%m/%y").strftime("%Y-%m-%d")
        except:
            continue
    return None


def extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name=None, part_name=None):
    extracted = []

    if not categories:
        st.warning("No categories found")
        return []

    col_date_map = {}
    for col in metric_cols:
        for row in range(1, 6):
            cell_val = sheet.cell(row=row, column=col).value
            if isinstance(cell_val, str):
                date_found = extract_date(cell_val)
                if date_found:
                    col_date_map[col] = date_found
                    break

    for i in range(len(categories)):
        current = categories[i]
        start_row = current['row']
        end_row = categories[i + 1]['row'] - 1 if i + 1 < len(categories) else sheet.max_row

        for row in range(start_row, end_row + 1):
            subcat_cell = sheet.cell(row=row, column=subcategory_col).value
            if not subcat_cell:
                continue
            subcat = str(subcat_cell).strip()

            for col in metric_cols:
                val = sheet.cell(row=row, column=col).value
                if not isinstance(val, (int, float)):
                    continue

                raw_header = headers.get(col, f"Column_{chr(64 + col)}").strip().lower()
                if "cm%" in raw_header:
                    continue

                # Normalize metric to first word (e.g., "Actual Performance" → "Actual")
                metric = raw_header.split()[0].capitalize()

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


# Streamlit UI
st.title("📊 SMITCH Excel Extractor")
uploaded_files = st.file_uploader("Upload Excel files", type=["xlsm", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        st.subheader(file.name)
        wb = load_workbook(file, data_only=True)
        sheet = wb.active

        metric_cols, headers, _ = detect_metric_columns(sheet)
        categories = detect_categories(sheet)
        subcategory_col = find_subcategory_column(sheet, categories)
        plant_name, _ = detect_plant(sheet)
        part_name = detect_part_name(sheet, categories)

        df = pd.DataFrame(extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name, part_name))

        if not df.empty:
            st.success(f" Extracted {len(df)} rows")
            st.dataframe(df.head(10))

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False)
            buffer.seek(0)

            st.download_button(
                label=" Download Extracted Excel",
                data=buffer,
                file_name=f"{file.name.split('.')[0]}_extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No valid data extracted.")
else:
    st.info(" Upload Excel files to begin.")
