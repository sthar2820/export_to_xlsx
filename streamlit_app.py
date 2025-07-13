
# import streamlit as st
# import pandas as pd
# from openpyxl import load_workbook
# from io import BytesIO

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
#                         header_clean = cell.strip()
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

# ================================
# CORE LOGIC - ROBUST AND SIMPLE
# ================================

def detect_plant_and_part_name(sheet):
    """
    Detect plant name and part name from known plant list
    """
    # Known list of 14 plants (add your actual plant names here)
    known_plants = [
        "BIRMINGHAM", "DETROIT", "CHICAGO", "ATLANTA", "DALLAS", "PHOENIX", 
        "LOS ANGELES", "SEATTLE", "DENVER", "BOSTON", "MIAMI", "HOUSTON",
        "KANSAS CITY", "PORTLAND"  # Add your actual plant names
    ]
    
    plant_name = None
    part_name = None
    plant_row = None
    
    try:
        # Search for plant name in first 20 rows and first 5 columns
        for row in range(1, min(21, sheet.max_row + 1)):
            for col in range(1, min(6, sheet.max_column + 1)):
                cell_value = sheet.cell(row=row, column=col).value
                if cell_value and isinstance(cell_value, str):
                    cell_text = cell_value.strip().upper()
                    
                    # Check if this matches any known plant
                    for plant in known_plants:
                        if plant.upper() in cell_text:
                            plant_name = plant
                            plant_row = row
                            plant_col = col
                            break
                    
                    if plant_name:
                        break
            if plant_name:
                break
        
        # If plant found, look for part name in the same column between plant and first SMITCH category
        if plant_name and plant_row:
            # Find first SMITCH category row in the same column
            smitch_letters = ['S', 'M', 'I', 'T', 'C', 'H']
            first_smitch_row = None
            
            for row in range(plant_row + 1, min(plant_row + 15, sheet.max_row + 1)):
                cell_value = sheet.cell(row=row, column=plant_col).value
                if cell_value:
                    cell_str = str(cell_value).strip()
                    # Check for single letter or multi-line SMITCH categories
                    if (len(cell_str) == 1 and cell_str.upper() in smitch_letters) or \
                       any(line.strip().upper() in smitch_letters for line in cell_str.split('\n')):
                        first_smitch_row = row
                        break
            
            # Search for part name between plant and first SMITCH category
            search_end = first_smitch_row if first_smitch_row else min(plant_row + 10, sheet.max_row + 1)
            
            for row in range(plant_row + 1, search_end):
                cell_value = sheet.cell(row=row, column=plant_col).value
                if cell_value and isinstance(cell_value, str):
                    candidate = cell_value.strip()
                    
                    # Check if this looks like a part name
                    if (len(candidate) > 3 and  # Not too short
                        not candidate.isdigit() and  # Not just numbers
                        not candidate.upper() in known_plants and  # Not another plant name
                        any(c.isalpha() for c in candidate)):  # Contains letters
                        part_name = candidate
                        break
        
    except Exception as e:
        pass  # Silent fallback
    
    return plant_name, part_name

def detect_metric_columns(sheet, stop_at_keywords=None):
    """
    Robust metric column detection with flexible stop keywords
    """
    if stop_at_keywords is None:
        stop_at_keywords = [
            "demon-strated rate at 100%", 
            "demonstrated rate at 100%",
            "demon-strated rate", 
            "demonstrated rate",
            "baseline", 
            "actual", 
            "final", 
            "target", 
            "current"
        ]
    
    metric_cols = []
    headers = {}
    stop_column_found = None
    
    try:
        # Search rows 1-5 for headers
        for search_row in range(1, min(6, sheet.max_row + 1)):
            temp_cols = []
            temp_headers = {}
            temp_stop_col = None
            
            for col in range(3, min(sheet.max_column + 1, 20)):  # Reasonable limit
                try:
                    cell = sheet.cell(row=search_row, column=col).value
                    if cell and isinstance(cell, str) and len(cell.strip()) > 1:
                        header_clean = cell.strip()
                        temp_headers[col] = header_clean
                        temp_cols.append(col)
                        
                        # Check for any stop keyword
                        header_lower = header_clean.lower()
                        for stop_keyword in stop_at_keywords:
                            if stop_keyword in header_lower:
                                temp_stop_col = stop_keyword
                                break
                        
                        # If we found a stop keyword, include this column and break
                        if temp_stop_col:
                            break
                except:
                    continue
            
            # Use the row with most headers or the one with stop column
            if temp_stop_col or len(temp_headers) > len(headers):
                headers = temp_headers
                metric_cols = temp_cols
                if temp_stop_col:
                    stop_column_found = temp_stop_col
                    break  # Prefer the row with stop column
        
        # Fallback if no headers found
        if not metric_cols:
            metric_cols = list(range(3, min(8, sheet.max_column + 1)))  # Default C-G
            for col in metric_cols:
                headers[col] = f"Column_{chr(64 + col)}"
        
    except Exception as e:
        # Emergency fallback
        metric_cols = [3, 4, 5, 6]
        headers = {3: "Column_C", 4: "Column_D", 5: "Column_E", 6: "Column_F"}
    
    return metric_cols, headers, stop_column_found

def detect_categories(sheet):
    """
    Robust category detection with fallback
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
    
    try:
        # Check first 3 columns for categories
        for col in range(1, min(4, sheet.max_column + 1)):
            for row in range(1, min(sheet.max_row + 1, 50)):  # Limit search
                try:
                    val = sheet.cell(row=row, column=col).value
                    if not val:
                        continue
                        
                    # Handle different cell formats
                    text = str(val).strip()
                    
                    # Split on newlines if present
                    lines = text.split('\n') if '\n' in text else [text]
                    
                    for line in lines:
                        line_clean = line.strip().upper()
                        if len(line_clean) == 1 and line_clean in category_map:
                            # Check if already found
                            if not any(c['letter'] == line_clean for c in categories):
                                categories.append({
                                    'row': row,
                                    'column': col,
                                    'letter': line_clean,
                                    'name': category_map[line_clean]
                                })
                            break
                except:
                    continue
        
        # Sort by row number
        categories.sort(key=lambda x: x['row'])
        
    except Exception as e:
        st.error(f"Error detecting categories: {e}")
        categories = []
    
    return categories

def find_subcategory_column(sheet, categories):
    """
    Find subcategory column with robust fallback
    """
    try:
        if not categories:
            return 3  # Default fallback
        
        # Start with the most common pattern
        category_col = categories[0]['column']
        candidates = [category_col + 1, category_col + 2, 3, 2]  # Common patterns
        
        best_col = category_col + 1  # Default
        max_text_cells = 0
        
        for col in candidates:
            if col < 1 or col > sheet.max_column:
                continue
            
            text_cells = 0
            try:
                for row in range(1, min(30, sheet.max_row + 1)):
                    cell = sheet.cell(row=row, column=col).value
                    if cell and isinstance(cell, str) and len(cell.strip()) >= 2:
                        text_cells += 1
                
                if text_cells > max_text_cells:
                    max_text_cells = text_cells
                    best_col = col
            except:
                continue
        
        return best_col
        
    except Exception as e:
        return 3  # Safe fallback

def extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name=None, part_name=None):
    """
    Robust data extraction with comprehensive error handling
    """
    extracted = []
    
    try:
        if not categories:
            st.warning("No categories found")
            return []
        
        for i in range(len(categories)):
            try:
                current = categories[i]
                start_row = current['row']
                
                # Calculate end row safely
                if i + 1 < len(categories):
                    end_row = categories[i + 1]['row'] - 1
                else:
                    end_row = min(start_row + 25, sheet.max_row)
                
                # Extract data within boundaries
                for row in range(start_row, end_row + 1):
                    try:
                        # Get subcategory
                        subcat_cell = sheet.cell(row=row, column=subcategory_col).value
                        if not subcat_cell:
                            continue
                        
                        subcat = str(subcat_cell).strip()
                        if len(subcat) < 2:
                            continue
                        
                        # Extract numeric data
                        for col in metric_cols:
                            try:
                                val = sheet.cell(row=row, column=col).value
                                if isinstance(val, (int, float)) and val is not None:
                                    # Get header safely
                                    header = headers.get(col, f"Column_{chr(64 + col)}")
                                    if isinstance(header, str) and '\n' in header:
                                        header = header.split('\n')[0]
                                    header = str(header)[:30]  # Truncate long headers
                                    
                                    data_point = {
                                        'Category': current['name'],
                                        'Subcategory': subcat,
                                        'Metric': header,
                                        'Value': float(val)
                                    }
                                    
                                    # Add plant name and part name if available
                                    if plant_name:
                                        data_point['Plant Name'] = plant_name
                                    if part_name:
                                        data_point['Part Name'] = part_name
                                    
                                    extracted.append(data_point)
                            except Exception as col_error:
                                continue  # Skip problematic cells
                                
                    except Exception as row_error:
                        continue  # Skip problematic rows
                        
            except Exception as cat_error:
                continue  # Skip problematic categories
        
    except Exception as e:
        st.error(f"Error during extraction: {e}")
    
    return extracted

# ================================
# STREAMLIT APP - SIMPLE AND ROBUST
# ================================

st.title("📊 SMITCH Excel Extractor")
st.write("Upload SMITCH Excel files to extract structured data")

uploaded_files = st.file_uploader(
    "Choose Excel files", 
    type=["xlsm", "xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"Processing {len(uploaded_files)} file(s)...")
    
    for file in uploaded_files:
        st.subheader(f"📂 {file.name}")
        
        try:
            # Load file with error handling
            wb = load_workbook(file, data_only=True)
            ws = wb.active
            
            st.write(f"✅ File loaded: {ws.max_row} rows × {ws.max_column} columns")
            
            # Step 1: Detect plant and part name
            with st.spinner("Detecting plant and part name..."):
                plant_name, part_name = detect_plant_and_part_name(ws)
            
            # Step 2: Detect structure
            with st.spinner("Detecting file structure..."):
                metric_columns, headers, stop_column_found = detect_metric_columns(ws)
                category_rows = detect_categories(ws)
                subcategory_col = find_subcategory_column(ws, category_rows)
            
            # Show detection results
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("Categories", len(category_rows))
            with col2:
                st.metric("Metric Columns", len(metric_columns))
            with col3:
                st.metric("Subcategory Col", chr(64 + subcategory_col))
            with col4:
                if stop_column_found:
                    st.metric("Stop Column", stop_column_found.title())
                else:
                    st.metric("Stop Column", "Auto-detected")
            with col5:
                if plant_name:
                    st.metric("Plant", plant_name)
                else:
                    st.metric("Plant", "Not found")
            
            # Show plant and part info
            if plant_name or part_name:
                st.info(f"🏭 Plant: {plant_name or 'Not detected'} | 🔧 Part: {part_name or 'Not detected'}")
            
            # Step 3: Extract data
            with st.spinner("Extracting data..."):
                data = extract_smitch_data(ws, category_rows, metric_columns, headers, subcategory_col, plant_name, part_name)
            
            if data:
                df = pd.DataFrame(data)
                st.success(f"✅ Extracted {len(df)} records")
                
                # Show summary
                st.write("**Categories found:**")
                for cat, count in df['Category'].value_counts().items():
                    st.write(f"• {cat}: {count} records")
                
                # Show preview
                st.write("**Data preview:**")
                st.dataframe(df.head(10))
                
                # Download button
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Extracted')
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Excel",
                    data=output,
                    file_name=f"{file.name.split('.')[0]}_extracted.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.warning("⚠️ No data extracted from this file")
                
                # Debug info
                with st.expander("Debug Information"):
                    st.write(f"Plant name detected: {plant_name or 'None'}")
                    st.write(f"Part name detected: {part_name or 'None'}")
                    st.write(f"Categories detected: {len(category_rows)}")
                    if category_rows:
                        for cat in category_rows:
                            st.write(f"• {cat['name']} at row {cat['row']}")
                    
                    st.write(f"Metric columns: {metric_columns}")
                    st.write(f"Headers detected: {list(headers.values())}")
                    st.write(f"Stop column found: {stop_column_found if stop_column_found else 'None'}")
                    st.write(f"Subcategory column: {subcategory_col}")
                    
                    # Show sample data from subcategory column
                    st.write("Sample from subcategory column:")
                    for row in range(1, min(10, ws.max_row + 1)):
                        cell = ws.cell(row=row, column=subcategory_col).value
                        if cell:
                            st.write(f"Row {row}: {cell}")
        
        except Exception as e:
            st.error(f"❌ Failed to process {file.name}")
            st.error(f"Error: {str(e)}")
            
            # Debug information
            with st.expander("Error Details"):
                st.code(str(e))
                st.write("This might be due to:")
                st.write("• Unexpected file structure")
                st.write("• Corrupted file")
                st.write("• Different SMITCH template version")
        
        st.divider()

else:
    st.info("👆 Upload Excel files to get started")
    
    with st.expander("How it works"):
        st.write("""
        1. **Auto-detects** plant name from known plant list
        2. **Finds** part name between plant and first SMITCH category
        3. **Auto-detects** SMITCH categories (S, M, I, T, C, H)
        4. **Finds** subcategory and data columns
        5. **Smart stop detection** for final columns (in priority order):
           - "Demon-strated Rate at 100%" / "Demonstrated Rate at 100%"
           - "Demon-strated Rate" / "Demonstrated Rate"
           - "Baseline"
           - "Actual" 
           - "Final"
           - "Target"
           - "Current"
        6. **Extracts** structured data with plant and part info
        7. **Exports** to Excel format
        
        **Output format:**
        - Plant Name: Auto-detected from known plants
        - Part Name: Found between plant and SMITCH categories
        - Category: Sales Price, Material, etc.
        - Subcategory: Total, Recliner Bushing, etc.  
        - Metric: Column headers
        - Value: Numeric data
        
        **Plant Detection:**
        - Searches first 20 rows for known plant names
        - Supports 14+ plant locations
        
        **Part Name Detection:**
        - Looks in same column as plant name
        - Searches between plant row and first SMITCH category
        - Finds first meaningful text (not numbers or plant names)
        """)
            

                       



   
