import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import zipfile
from datetime import datetime

# ================================
# PAGE CONFIGURATION
# ================================

st.set_page_config(
    page_title="SMITCH Excel Extractor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================================
# IMPROVED DETECTION FUNCTIONS
# ================================

def detect_metric_columns(sheet, header_row=3, stop_at="baseline"):
    metric_cols = []
    for col in range(4, sheet.max_column + 1):  # Start from Column D
        val = sheet.cell(row=header_row, column=col).value
        if val and isinstance(val, str):
            metric_cols.append(col)
            if stop_at.lower() in val.lower():
                break  # Stop including after 'Baseline'
    return metric_cols

def detect_categories(sheet):
    """
    Enhanced category detection with better multi-line handling
    """
    categories = []
    category_map = {
        'S': 'Sales Price',
        'M': 'Material',
        'I': 'Investment',
        'T': 'Tooling',
        'C': 'Cycle Times',
        'H': 'Headcount'
    }
    
    # Check multiple columns for categories (not just column 2)
    for col in range(1, min(4, sheet.max_column + 1)):
        for row in range(1, sheet.max_row + 1):
            val = sheet.cell(row=row, column=col).value
            if val:
                # Handle multi-line cells
                text = str(val).strip()
                lines = text.split('\n') if '\n' in text else [text]
                
                for line in lines:
                    line_clean = line.strip().upper()
                    if line_clean in category_map:
                        # Check if we already found this category
                        existing = [c for c in categories if c['letter'] == line_clean]
                        if not existing:
                            categories.append({
                                'row': row,
                                'column': col,
                                'letter': line_clean,
                                'name': category_map[line_clean]
                            })
                        break
    
    # Sort by row number
    categories.sort(key=lambda x: x['row'])
    return categories

def find_subcategory_column(sheet, categories):
    """
    Auto-detect which column contains subcategories
    """
    if not categories:
        return 3  # Default fallback
    
    # Check columns near the category column
    category_col = categories[0]['column']
    candidate_cols = [category_col + 1, category_col + 2, category_col - 1]
    
    best_col = category_col + 1  # Default
    max_subcats = 0
    
    for col in candidate_cols:
        if col < 1 or col > sheet.max_column:
            continue
            
        subcat_count = 0
        for row in range(1, min(30, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=col).value
            if cell and isinstance(cell, str) and len(cell.strip()) >= 2:
                # Check if it looks like a subcategory
                text = cell.strip().lower()
                if any(word in text for word in ['total', 'bracket', 'bushing', 'spacer', 'welding', 'recliner']):
                    subcat_count += 1
        
        if subcat_count > max_subcats:
            max_subcats = subcat_count
            best_col = col
    
    return best_col

def extract_smitch_data(sheet, categories, metric_cols, headers):
    """
    Robust data extractor: uses fixed subcategory column (3) and improves boundary logic.
    """
    extracted = []
    
    for i in range(len(categories)):
        current = categories[i]
        start_row = current['row']
        end_row = categories[i + 1]['row'] - 1 if i + 1 < len(categories) else sheet.max_row

        for row in range(start_row, end_row + 1):
            subcat_cell = sheet.cell(row=row, column=3).value  # COLUMN C — fixed
            if not subcat_cell or len(str(subcat_cell).strip()) < 2:
                continue
            
            subcat = str(subcat_cell).strip()

            for col in metric_cols:
                val = sheet.cell(row=row, column=col).value
                if isinstance(val, (int, float)):
                    header = headers.get(col, sheet.cell(row=3, column=col).value)
                    if header:
                        header_clean = str(header).split('\n')[0].strip()[:30]
                    else:
                        header_clean = f"Column_{col}"

                    extracted.append({
                        'Category': current['name'],
                        'Subcategory': subcat,
                        'Metric': header_clean,
                        'Value': float(val),
                    })
    
    return extracted


# ================================
# STREAMLIT UI
# ================================

st.title("📊 SMITCH Excel Extractor App")
st.markdown("### Extract and standardize SMITCH manufacturing data from Excel files")

# Sidebar for configuration
with st.sidebar:
    st.header("⚙️ Configuration")
    
    stop_at_baseline = st.checkbox("Stop at 'Baseline' column", value=True)
    custom_stop_word = st.text_input("Custom stop word:", value="baseline" if stop_at_baseline else "")
    
    include_debug_info = st.checkbox("Include debug columns (Row, Column, Excel_Cell)", value=False)
    
    st.header("📋 Instructions")
    st.markdown("""
    1. **Upload** one or more SMITCH Excel files
    2. **Review** the extraction results  
    3. **Download** individual files or bulk ZIP
    4. **Check** the summary statistics
    """)

# File upload section
uploaded_files = st.file_uploader(
    "Upload SMITCH Excel Files", 
    type=["xlsm", "xlsx"], 
    accept_multiple_files=True,
    help="Select one or more SMITCH Excel files to process"
)

if uploaded_files:
    st.success(f"📁 {len(uploaded_files)} file(s) uploaded successfully!")
    
    # Processing options
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader("🔄 Processing Results")
    with col2:
        if len(uploaded_files) > 1:
            bulk_download = st.button("📦 Download All as ZIP", key="bulk_download")
        else:
            bulk_download = False
    
    # Store results for bulk download
    all_results = {}
    total_records = 0
    processing_summary = []
    
    # Process each file
    for idx, file in enumerate(uploaded_files):
        with st.expander(f"📂 {file.name}", expanded=len(uploaded_files) <= 3):
            try:
                # Load workbook
                with st.spinner(f"Processing {file.name}..."):
                    wb = load_workbook(file, data_only=True)
                    ws = wb.active
                
                # Auto-detection
                metric_columns, headers = detect_metric_columns(ws, stop_at=custom_stop_word if custom_stop_word else "baseline")
                category_rows = detect_categories(ws)
                subcategory_col = find_subcategory_column(ws, category_rows)
                
                # Show detection results
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Categories Found", len(category_rows))
                with col2:
                    st.metric("Metric Columns", len(metric_columns))
                with col3:
                    st.metric("Subcategory Column", chr(64 + subcategory_col))
                
                # Extract data
                data = extract_smitch_data(ws, category_rows, metric_columns, headers, subcategory_col)
                
                if data:
                    # Create DataFrame
                    df = pd.DataFrame(data)
                    
                    # Select columns based on user preference
                    if include_debug_info:
                        display_columns = ['Category', 'Subcategory', 'Metric', 'Value', 'Row', 'Column', 'Excel_Cell']
                    else:
                        display_columns = ['Category', 'Subcategory', 'Metric', 'Value']
                    
                    df_display = df[display_columns]
                    
                    # Show summary
                    st.success(f"✅ Extracted {len(df)} records")
                    
                    # Summary statistics
                    summary_col1, summary_col2 = st.columns(2)
                    with summary_col1:
                        st.write("**Categories:**")
                        category_summary = df['Category'].value_counts()
                        for cat, count in category_summary.items():
                            st.write(f"• {cat}: {count} records")
                    
                    with summary_col2:
                        st.write("**Metrics:**")
                        metric_summary = df['Metric'].value_counts()
                        for metric, count in metric_summary.head(5).items():
                            st.write(f"• {metric}: {count} records")
                    
                    # Show data preview
                    st.write("**Data Preview:**")
                    st.dataframe(df_display.head(10), use_container_width=True)
                    
                    # Individual file download
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_display.to_excel(writer, index=False, sheet_name='Extracted_Data')
                        
                        # Add summary sheet
                        summary_data = {
                            'Metric': ['Total Records', 'Categories', 'Subcategories', 'Metrics'],
                            'Value': [len(df), df['Category'].nunique(), df['Subcategory'].nunique(), df['Metric'].nunique()]
                        }
                        pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name='Summary')
                    
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Download Excel",
                        data=output,
                        file_name=f"{file.name.split('.')[0]}_extracted.xlsx",
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=f"download_{idx}"
                    )
                    
                    # Store for bulk download
                    all_results[file.name] = df_display
                    total_records += len(df)
                    
                    processing_summary.append({
                        'File': file.name,
                        'Status': 'Success',
                        'Records': len(df),
                        'Categories': df['Category'].nunique(),
                        'Subcategories': df['Subcategory'].nunique()
                    })
                
                else:
                    st.error("❌ No data extracted - check file structure")
                    processing_summary.append({
                        'File': file.name,
                        'Status': 'Failed - No Data',
                        'Records': 0,
                        'Categories': 0,
                        'Subcategories': 0
                    })
                
            except Exception as e:
                st.error(f"❌ Failed to process {file.name}: {str(e)}")
                processing_summary.append({
                    'File': file.name,
                    'Status': f'Error: {str(e)[:50]}...',
                    'Records': 0,
                    'Categories': 0,
                    'Subcategories': 0
                })
    
    # Bulk download functionality
    if bulk_download and all_results:
        with st.spinner("Creating ZIP file..."):
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for filename, df in all_results.items():
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Extracted_Data')
                    output.seek(0)
                    
                    excel_filename = f"{filename.split('.')[0]}_extracted.xlsx"
                    zip_file.writestr(excel_filename, output.getvalue())
                
                # Add summary file
                summary_df = pd.DataFrame(processing_summary)
                summary_output = BytesIO()
                with pd.ExcelWriter(summary_output, engine='xlsxwriter') as writer:
                    summary_df.to_excel(writer, index=False, sheet_name='Processing_Summary')
                summary_output.seek(0)
                zip_file.writestr('Processing_Summary.xlsx', summary_output.getvalue())
            
            zip_buffer.seek(0)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            st.download_button(
                label="📦 Download ZIP with All Files",
                data=zip_buffer,
                file_name=f"SMITCH_Extracted_{timestamp}.zip",
                mime="application/zip"
            )
    
    # Final summary
    if processing_summary:
        st.subheader("📊 Processing Summary")
        summary_df = pd.DataFrame(processing_summary)
        st.dataframe(summary_df, use_container_width=True)
        
        # Overall stats
        successful_files = len([s for s in processing_summary if s['Status'] == 'Success'])
        st.info(f"📈 Successfully processed {successful_files}/{len(uploaded_files)} files with {total_records} total records")

else:
    # Instructions when no files uploaded
    st.info("👆 Upload SMITCH Excel files to get started")
    
    with st.expander("ℹ️ What this app does"):
        st.markdown("""
        This app automatically extracts and standardizes SMITCH manufacturing data:
        
        **🔍 Auto-Detection:**
        - Finds SMITCH categories (S, M, I, T, C, H) in any column
        - Detects subcategory column automatically  
        - Stops at 'Baseline' column or custom word
        
        **📊 Data Structure:**
        - Category: Sales Price, Material, Investment, etc.
        - Subcategory: Total, Recliner Bushing Adap, etc.
        - Metric: Column headers (Total Score, PLEX Standard, etc.)
        - Value: Numeric data
        
        **💾 Output Options:**
        - Individual Excel files per input
        - Bulk ZIP download for multiple files
        - Processing summary report
        """)
