import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime
import re
from dateutil import parser
import json
from pathlib import Path
from collections import defaultdict
from rapidfuzz import process, fuzz
import traceback

# Load or initialize normalization map
MAP_FILE = "normalization_map.json"
map_path = Path(MAP_FILE)
if map_path.exists():
    with open(map_path, "r") as f:
        normalization_map = json.load(f)
else:
    normalization_map = {}

reverse_map = defaultdict(list)
for raw, norm in normalization_map.items():
    reverse_map[norm].append(raw)

def normalize_dynamic(label, threshold=85):
    """Enhanced normalization specifically for SMITCH subcategories"""
    if not isinstance(label, str) or not label.strip():
        return label
    
    original_label = label
    
    # Direct hit in existing map
    if label in normalization_map:
        return normalization_map[label]
    
    # SMITCH-specific cleaning rules
    clean_label = label
    
    # Remove directional arrows and line breaks commonly found in SMITCH files
    clean_label = re.sub(r'→.*?↓', '', clean_label)  # Remove "→\nSub comps ↓"
    clean_label = re.sub(r'[\r\n]+', ' ', clean_label)  # Replace line breaks with spaces
    clean_label = re.sub(r'\s+', ' ', clean_label)  # Normalize whitespace
    clean_label = clean_label.strip()
    
    # SMITCH manufacturing subcategory mappings
    smitch_mappings = {
        # Common SMITCH terms
        "total": "Total",
        "sub comps": "Sub Components",
        "assy": "Assembly",
        "bushing": "Bushing",
        "welding": "Welding", 
        "weld": "Welding",
        "stud": "Stud",
        "bolt": "Bolt",
        "stop": "Stop",
        "guide": "Guide",
        "recliner": "Recliner",
        "adap": "Adapter",
        "adapter": "Adapter",
        
        # Rate and percentage terms
        "rate at": "Rate At",
        "cm %": "CM Percentage",
        "demonstrated rate": "Demonstrated Rate",
        "actual performance": "Actual Performance",
        
        # Cost terms
        "quoted cost": "Quoted Cost",
        "plex standard": "PLEX Standard",
        "loss per piece": "Loss Per Piece",
        "total loss": "Total Loss"
    }
    
    # Apply SMITCH-specific mappings
    clean_lower = clean_label.lower()
    for key, canonical in smitch_mappings.items():
        if key in clean_lower:
            result = canonical
            normalization_map[original_label] = result
            reverse_map[result].append(original_label)
            save_normalization_map()
            return result
    
    # Try fuzzy match against existing normalized values
    choices = list(set(reverse_map.keys()))
    if choices:
        best_match, score = process.extractOne(clean_label, choices, scorer=fuzz.ratio)
        if score >= threshold:
            normalization_map[original_label] = best_match
            reverse_map[best_match].append(original_label)
            save_normalization_map()
            return best_match
    
    # Create new canonical form using the cleaned label
    canonical = clean_label.title().strip()
    normalization_map[original_label] = canonical
    reverse_map[canonical].append(original_label)
    save_normalization_map()
    return canonical

def save_normalization_map():
    """Save normalization map with error handling"""
    try:
        with open(MAP_FILE, "w") as f:
            json.dump(normalization_map, f, indent=2)
    except Exception as e:
        st.warning(f"Could not save normalization map: {e}")

KNOWN_PLANTS = {
    "Bielsko Biala", "Birmingham", "Blatna", "Einbeck", "Forsheda",
    "Olofstrom", "Rotenburg", "Celaya", "Dickson", "Goshen",
    "Kalamazoo", "Saltillo", "Valley City", "Wellington"
}

def detect_metric_columns(sheet, stop_at_keywords=None):
    """
    Enhanced metric column detection that works better with SMITCH files
    """
    if stop_at_keywords is None:
        stop_at_keywords = [
            "demon-strated rate at 100%", "demonstrated rate at 100%",
            "demon-strated rate", "demonstrated rate",
            "smitch score", "total points"
        ]

    metric_cols = []
    headers = {}
    stop_column_found = None

    try:
        # Focus on row 3 first (where SMITCH headers typically are)
        priority_rows = [3, 2, 1, 4, 5]
        
        for search_row in priority_rows:
            if search_row > sheet.max_row:
                continue
                
            temp_cols = []
            temp_headers = {}
            temp_stop_col = None

            # Look at columns D onwards (columns 4+) for SMITCH metric headers  
            for col in range(4, min(sheet.max_column + 1, 15)):
                try:
                    cell = sheet.cell(row=search_row, column=col).value
                    
                    # Only process cells with meaningful string content
                    if cell and isinstance(cell, str) and len(cell.strip()) > 2:
                        # Clean header by removing line breaks and extra spaces
                        header_clean = re.sub(r'[\r\n]+', ' ', str(cell))
                        header_clean = re.sub(r'\s+', ' ', header_clean).strip()
                        
                        temp_headers[col] = header_clean
                        temp_cols.append(col)

                        # Check for stop keywords
                        header_lower = header_clean.lower()
                        for stop_keyword in stop_at_keywords:
                            if stop_keyword in header_lower:
                                temp_stop_col = stop_keyword
                                break

                        if temp_stop_col:
                            break
                            
                except Exception:
                    continue

            # Keep the best row found so far (prioritize row 3)
            if temp_stop_col or len(temp_headers) > len(headers) or search_row == 3:
                headers = temp_headers
                metric_cols = temp_cols
                if temp_stop_col:
                    stop_column_found = temp_stop_col
                    break

        # Enhanced fallback for SMITCH files
        if not metric_cols:
            # SMITCH files typically have metrics in columns D-K (4-11)
            metric_cols = list(range(4, min(12, sheet.max_column + 1)))
            for col in metric_cols:
                # Try to get header from row 3, then row 2, then default
                header = None
                for row in [3, 2, 1]:
                    try:
                        cell_val = sheet.cell(row=row, column=col).value
                        if cell_val and isinstance(cell_val, str) and len(cell_val.strip()) > 2:
                            header = re.sub(r'[\r\n]+', ' ', str(cell_val)).strip()
                            break
                    except:
                        continue
                
                headers[col] = header or f"Column_{chr(64 + col)}"

    except Exception as e:
        st.warning(f"Error in metric column detection: {e}")
        # Emergency fallback specifically for SMITCH
        metric_cols = [4, 5, 6, 7, 8, 9, 10]  # D through J
        headers = {
            4: "QUOTED COST MODEL", 5: "PLEX STANDARD", 6: "ACTUAL PERFORMANCE",
            7: "Demonstrated Rate", 8: "Delta Quote Actual", 9: "Delta Quote Plex",
            10: "Delta Plex Actual"
        }

    return metric_cols, headers, stop_column_found

def detect_categories(sheet):
    """Enhanced category detection for SMITCH files"""
    categories = []
    category_map = {
        'S': 'Sales Price', 'M': 'Material', 'I': 'Investment',
        'T': 'Tooling', 'C': 'Cycle Times', 'H': 'Headcount'
    }

    try:
        # Focus on column B (column 2) where SMITCH categories typically are
        for row in range(1, min(sheet.max_row + 1, 50)):
            try:
                val = sheet.cell(row=row, column=2).value  # Column B
                if not val:
                    continue
                    
                text = str(val).strip()
                
                # Handle multi-line cells (common in SMITCH)
                lines = text.split('\n') if '\n' in text else [text]
                
                for line in lines:
                    line_clean = line.strip().upper()
                    if len(line_clean) == 1 and line_clean in category_map:
                        # Avoid duplicates
                        if not any(c['letter'] == line_clean for c in categories):
                            categories.append({
                                'row': row, 'column': 2,
                                'letter': line_clean,
                                'name': category_map[line_clean]
                            })
                        break
            except Exception:
                continue
                
        categories.sort(key=lambda x: x['row'])
        
    except Exception as e:
        st.error(f"Error detecting categories: {e}")
        categories = []

    return categories

def detect_plant(sheet):
    """Enhanced plant detection for SMITCH files"""
    try:
        # In SMITCH files, plant is typically in row 2, column B
        plant_cell = sheet.cell(row=2, column=2).value
        if plant_cell and isinstance(plant_cell, str):
            plant_text = plant_cell.strip()
            # Direct match first
            for plant in KNOWN_PLANTS:
                if plant.lower() == plant_text.lower():
                    return plant, 2
            
            # Fuzzy match as backup
            best_match, score = process.extractOne(plant_text, KNOWN_PLANTS, scorer=fuzz.ratio)
            if score >= 70:  # Lower threshold since SMITCH plant names are usually exact
                return best_match, 2
        
        # Fallback: search more broadly
        for row in range(1, min(sheet.max_row + 1, 10)):
            for col in range(1, min(sheet.max_column + 1, 5)):
                val = sheet.cell(row=row, column=col).value
                if val and isinstance(val, str):
                    text = val.strip()
                    for plant in KNOWN_PLANTS:
                        if plant.lower() in text.lower():
                            return plant, row
                            
    except Exception as e:
        st.warning(f"Error detecting plant: {e}")
        
    return None, None

def detect_part_name(sheet, categories):
    """Enhanced part name detection for SMITCH files"""
    try:
        # In SMITCH files, part name is typically in row 3, column B
        part_cell = sheet.cell(row=3, column=2).value
        if part_cell and isinstance(part_cell, str):
            part_text = part_cell.strip()
            # Check if it looks like a part number (alphanumeric, reasonable length)
            if len(part_text) >= 3 and part_text.upper() not in {'S', 'M', 'I', 'T', 'C', 'H'}:
                return part_text
        
        # Also check cell E1 and D1 where part numbers sometimes appear
        for row, col in [(1, 5), (1, 4)]:  # E1, D1
            try:
                val = sheet.cell(row=row, column=col).value
                if val and isinstance(val, str):
                    val = val.strip()
                    if len(val) >= 3 and re.match(r'^[A-Z0-9]+$', val):
                        return val
            except:
                continue
                
        # Fallback to original logic
        if categories:
            first_category_row = categories[0]['row']
            for row in range(first_category_row - 1, 0, -1):
                val = sheet.cell(row=row, column=2).value
                if val and isinstance(val, str):
                    val = val.strip()
                    if len(val) > 3 and val.upper() not in {'S', 'M', 'I', 'T', 'C', 'H'}:
                        return val
                        
    except Exception as e:
        st.warning(f"Error detecting part name: {e}")
        
    return None

def extract_date(text):
    """Enhanced date extraction for SMITCH files"""
    if not isinstance(text, str):
        return None

    # Common SMITCH date patterns
    patterns = [
        (r"\b\d{1,2}[/-]\d{1,2}[/-]\d{4}\b", ["%m/%d/%Y", "%d/%m/%Y"]),
        (r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b", ["%Y/%m/%d", "%Y/%d/%m"]),
        (r"\b\d{1,2}[/-]\d{4}\b", ["%m/%Y"]),
        (r"\b\d{1,2}[/-]\d{2}\b", ["%m/%y"]),
    ]
    
    for pattern, formats in patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            for fmt in formats:
                try:
                    dt = datetime.strptime(match, fmt)
                    return dt.strftime("%Y-%m-%d")
                except ValueError:
                    continue
                    
    return None

def find_subcategory_column(sheet, categories):
    """Enhanced subcategory column detection for SMITCH files"""
    try:
        # In SMITCH files, subcategories are typically in column B (same as categories)
        # but we want to return column B (2) as the subcategory column
        return 2
        
    except Exception:
        return 2  # Default to column B for SMITCH files

def extract_smitch_data(sheet, categories, metric_cols, headers, subcategory_col, plant_name=None, part_name=None):
    """Enhanced SMITCH data extraction with better subcategory handling"""
    extracted = []
    col_date_map = {}
    
    # Extract dates from column headers (look in rows 1-5)
    for col in metric_cols:
        date_found = None
        for row in range(1, 6):
            try:
                cell_val = sheet.cell(row=row, column=col).value
                if isinstance(cell_val, str):
                    possible_date = extract_date(cell_val)
                    if possible_date:
                        date_found = possible_date
                        break
            except Exception:
                continue
        col_date_map[col] = date_found
    
    if not categories:
        st.warning("No categories found")
        return []

    # Process each category section
    for i in range(len(categories)):
        current = categories[i]
        start_row = current['row']
        end_row = categories[i + 1]['row'] - 1 if i + 1 < len(categories) else min(start_row + 25, sheet.max_row)

        # Process the category row itself first
        for target_row in range(start_row, end_row + 1):
            try:
                subcat_cell = sheet.cell(row=target_row, column=subcategory_col).value
                if not subcat_cell:
                    continue
                
                subcat_text = str(subcat_cell).strip()
                
                # Skip if this is just the category letter
                if len(subcat_text) == 1 and subcat_text.upper() in ['S', 'M', 'I', 'T', 'C', 'H']:
                    continue
                
                # Extract subcategory from multi-line cells
                if '\n' in subcat_text:
                    lines = subcat_text.split('\n')
                    # Take the second line if it exists and isn't a category letter
                    if len(lines) > 1:
                        subcat_text = lines[1].strip()
                    else:
                        subcat_text = lines[0].strip()
                
                # Normalize the subcategory
                subcat = normalize_dynamic(subcat_text)
                
                # Extract metrics from this row
                for col in metric_cols:
                    try:
                        cell_val = sheet.cell(row=target_row, column=col).value
                        if cell_val is None:
                            continue
                            
                        # Handle different cell types
                        if isinstance(cell_val, (int, float)):
                            numeric_value = float(cell_val)
                        else:
                            cell_str = str(cell_val).strip()
                            if not cell_str:
                                continue
                            
                            # Extract numeric value from string
                            numeric_matches = re.findall(r"[-+]?\d*\.?\d+", cell_str)
                            if not numeric_matches:
                                continue
                            numeric_value = float(numeric_matches[0])

                        # Clean and prepare header
                        header = headers.get(col, f"Column_{chr(64 + col)}")
                        header = str(header)
                        
                        # Remove line breaks and clean header
                        header = re.sub(r'[\r\n]+', ' ', header)
                        header = re.sub(r'\s+', ' ', header).strip()
                        
                        # Limit header length
                        if len(header) > 30:
                            header = header[:30] + "..."

                        date_str = col_date_map.get(col)

                        entry = {
                            'Category': current['name'],
                            'Subcategory': subcat,
                            'Date': date_str,
                            'Metric': header,
                            'Value': numeric_value
                        }

                        if plant_name:
                            entry['Plant'] = plant_name
                        if part_name:
                            entry['Part Name'] = part_name

                        extracted.append(entry)
                        
                    except Exception as cell_error:
                        continue
                        
            except Exception as row_error:
                continue

    return extracted

def extract_ebit_loss_block(sheet, plant_name=None, part_name=None):
    """Enhanced EBIT loss extraction for SMITCH files"""
    extracted = []
    key_metrics = {
        "weekly apw": "Weekly APW",
        "annualized loss": "Annualized Loss", 
        "var oh total per piece": "VAR OH Total per Piece",
        "labor total per piece": "Labor Total per Piece",
        "total loss/pc": "Total Loss per Piece",
        "total loss per piece": "Total Loss per Piece"
    }

    try:
        # Search in reasonable area for EBIT data
        for row in range(1, min(sheet.max_row + 1, 60)):
            for col in range(1, min(sheet.max_column + 1, 15)):
                try:
                    cell = sheet.cell(row=row, column=col).value
                    if not isinstance(cell, str):
                        continue
                        
                    cell_lower = cell.lower().strip()

                    for key in key_metrics:
                        if key in cell_lower:
                            # Look for value in adjacent cells (right, below)
                            for delta_row, delta_col in [(0, 1), (1, 0), (0, 2)]:
                                try:
                                    value_cell = sheet.cell(row=row + delta_row, column=col + delta_col).value
                                    if value_cell is not None:
                                        # Clean and convert value
                                        if isinstance(value_cell, (int, float)):
                                            numeric = float(value_cell)
                                        else:
                                            value_str = str(value_cell).replace("£", "").replace("$", "").replace(",", "").strip()
                                            numeric = float(value_str)
                                        
                                        entry = {
                                            "Category": "EBIT LOSS",
                                            "Subcategory": "",
                                            "Date": None,
                                            "Metric": key_metrics[key],
                                            "Value": numeric
                                        }
                                        if plant_name:
                                            entry["Plant"] = plant_name
                                        if part_name:
                                            entry["Part Name"] = part_name
                                        extracted.append(entry)
                                        break  # Found value, move to next key
                                except (ValueError, TypeError):
                                    continue
                            break  # Found key, move to next cell
                except Exception:
                    continue
    except Exception as e:
        st.warning(f"Error extracting EBIT loss data: {e}")
        
    return extracted

# Streamlit App
st.title("🏭 Enhanced SMITCH Excel Extractor")
st.write("Upload SMITCH Excel files to extract structured data with improved normalization")

# Configuration section
with st.expander("⚙️ Configuration & Debug"):
    st.write(f"Normalization Map Entries: {len(normalization_map)}")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Clear Normalization Map"):
            normalization_map.clear()
            reverse_map.clear()
            save_normalization_map()
            st.success("Normalization map cleared!")
    
    with col2:
        if st.button("Show Normalization Map"):
            if normalization_map:
                st.json(normalization_map)
            else:
                st.info("Normalization map is empty")

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsm", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.write(f"Processing {len(uploaded_files)} file(s)...")
    
    for file in uploaded_files:
        st.subheader(f"📄 {file.name}")
        
        try:
            # Load workbook
            wb = load_workbook(file, data_only=True)
            
            # Try to find the main data sheet (look for part number in sheet names)
            main_sheet = None
            part_number_pattern = r'[A-Z]{2}\d+'
            
            for sheet_name in wb.sheetnames:
                if re.search(part_number_pattern, sheet_name):
                    main_sheet = wb[sheet_name]
                    st.info(f"🎯 Using sheet: **{sheet_name}**")
                    break
            
            if not main_sheet:
                main_sheet = wb.active
                st.info(f"📊 Using active sheet: **{main_sheet.title}**")
            
            ws = main_sheet
            st.write(f"File loaded: {ws.max_row} rows × {ws.max_column} columns")
            
            with st.spinner("Detecting file structure..."):
                try:
                    metric_columns, headers, stop_column_found = detect_metric_columns(ws)
                    st.success(f"✅ Detected {len(metric_columns)} metric columns")
                except Exception as e:
                    st.error(f"❌ Error in detect_metric_columns: {e}")
                    st.text(traceback.format_exc())
                    continue

                category_rows = detect_categories(ws)
                subcategory_col = find_subcategory_column(ws, category_rows)
                plant_name, plant_row = detect_plant(ws)
                part_name = detect_part_name(ws, category_rows)

            # Display detection results
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Categories", len(category_rows))
            with col2:
                st.metric("Metric Columns", len(metric_columns))
            with col3:
                st.metric("Subcategory Col", chr(64 + subcategory_col) if subcategory_col <= 26 else f"Col{subcategory_col}")
            with col4:
                st.metric("Stop Column", stop_column_found.title() if stop_column_found else "Auto-detected")

            if plant_name:
                st.info(f"🏭 Plant detected: **{plant_name}**")
            if part_name:
                st.info(f"🔧 Part detected: **{part_name}**")

            # Show detected categories
            if category_rows:
                st.write("**Categories detected:**")
                for cat in category_rows:
                    st.write(f"• Row {cat['row']}: {cat['letter']} - {cat['name']}")

            with st.spinner("Extracting data..."):
                main_data = extract_smitch_data(ws, category_rows, metric_columns, headers, subcategory_col, plant_name, part_name)
                ebit_loss_data = extract_ebit_loss_block(ws, plant_name, part_name)
                data = main_data + ebit_loss_data

            if data:
                df = pd.DataFrame(data)
                st.success(f"✅ Extracted {len(df)} records")

                # Show category breakdown
                st.write("**Categories found:**")
                for cat, count in df['Category'].value_counts().items():
                    st.write(f"• {cat}: {count} records")

                # Show subcategory normalization results
                if 'Subcategory' in df.columns:
                    unique_subcats = df['Subcategory'].nunique()
                    st.write(f"**Normalized subcategories:** {unique_subcats} unique values")
                    
                    # Show the subcategories
                    with st.expander("View Subcategories"):
                        for subcat in sorted(df['Subcategory'].unique()):
                            count = len(df[df['Subcategory'] == subcat])
                            st.write(f"• \"{subcat}\": {count} records")

                # Data preview
                st.write("**Data preview:**")
                st.dataframe(df.head(10), use_container_width=True)

                # Show data quality metrics
                with st.expander("📊 Data Quality Metrics"):
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Records with Dates", len(df[df['Date'].notna()]))
                    with col2:
                        st.metric("Unique Metrics", df['Metric'].nunique())
                    with col3:
                        values = df['Value'].dropna()
                        st.metric("Avg Value", f"{values.mean():.2f}" if len(values) > 0 else "N/A")

                # Download functionality
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Extracted')
                    
                    # Add a summary sheet
                    summary_data = {
                        'Metric': ['Total Records', 'Categories', 'Unique Subcategories', 'Unique Metrics', 'Plant', 'Part Name'],
                        'Value': [len(df), df['Category'].nunique(), df['Subcategory'].nunique(), 
                                df['Metric'].nunique(), plant_name or 'Not detected', part_name or 'Not detected']
                    }
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, index=False, sheet_name='Summary')
                    
                output.seek(0)

                st.download_button(
                    label="📥 Download Enhanced Excel",
                    data=output,
                    file_name=f"{file.name.split('.')[0]}_enhanced_extracted.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.warning("⚠️ No data extracted from this file")

            # Save normalization map after each file
            save_normalization_map()

        except Exception as e:
            st.error(f"❌ Failed to process {file.name}")
            st.error(f"Error: {str(e)}")
            with st.expander("Full Error Details"):
                st.text(traceback.format_exc())
            
else:
    st.info("📤 Upload Excel files to get started")
    
# Show current normalization mappings
if normalization_map:
    with st.expander("🔄 Current Normalization Mappings"):
        st.write("These mappings will be applied to future uploads:")
        for original, normalized in list(normalization_map.items())[:10]:  # Show first 10
            st.write(f"• \"{original}\" → \"{normalized}\"")
        if len(normalization_map) > 10:
            st.write(f"... and {len(normalization_map) - 10} more mappings")
