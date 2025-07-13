import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

def detect_metric_columns(sheet, header_row=3, stop_at="baseline"):
    metric_cols = []
    for col in range(3, sheet.max_column + 1):
        cell = sheet.cell(row=header_row, column=col).value
        if cell and isinstance(cell, str):
            if stop_at.lower() in cell.lower():
                metric_cols.append(col)
                break
            metric_cols.append(col)
    return metric_cols

def detect_categories(sheet):
    categories = []
    category_map = {
        'S': 'Sales Price',
        'M': 'Material',
        'I': 'Investment',
        'T': 'Tooling',
        'C': 'Cycle Times',
        'H': 'Headcount'
    }
    for row in range(1, sheet.max_row + 1):
        val = sheet.cell(row=row, column=2).value
        if val:
            parts = str(val).split('\n')
            for p in parts:
                p_clean = p.strip().upper()
                if p_clean in category_map:
                    categories.append({
                        'row': row,
                        'letter': p_clean,
                        'name': category_map[p_clean]
                    })
                    break
    return categories

def extract_smitch_data(sheet, categories, metric_cols):
    extracted = []
    for i in range(len(categories)):
        current = categories[i]
        start_row = current['row']
        end_row = categories[i + 1]['row'] - 1 if i + 1 < len(categories) else sheet.max_row
        for row in range(start_row, end_row + 1):
            subcat = sheet.cell(row=row, column=3).value
            if not subcat or len(str(subcat).strip()) < 2:
                continue
            subcat = str(subcat).strip()
            for col in metric_cols:
                val = sheet.cell(row=row, column=col).value
                if isinstance(val, (int, float)):
                    header = sheet.cell(row=3, column=col).value
                    header_clean = str(header).split('\n')[0].strip()[:30] if header else f"Column_{col}"
                    extracted.append({
                        'Category': current['name'],
                        'Subcategory': subcat,
                        'Metric': header_clean,
                        'Value': float(val)
                    })
    return extracted

st.title("📊 SMITCH Excel Extractor App")

uploaded_files = st.file_uploader("Upload SMITCH Excel Files", type=["xlsm", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        st.subheader(f"📂 Processing: {file.name}")
        try:
            wb = load_workbook(file, data_only=True)
            ws = wb.active

            metric_columns = detect_metric_columns(ws)
            category_rows = detect_categories(ws)
            data = extract_smitch_data(ws, category_rows, metric_columns)
            df = pd.DataFrame(data)[['Category', 'Subcategory', 'Metric', 'Value']]

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Extracted')
            output.seek(0)

            st.download_button(
                label="📥 Download Extracted Excel",
                data=output,
                file_name=f"{file.name.split('.')[0]}_extracted.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            st.error(f"❌ Failed to process {file.name}: {e}")
