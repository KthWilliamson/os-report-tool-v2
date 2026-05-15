import streamlit as st
import csv
import openpyxl
import io
import copy
from collections import defaultdict
from datetime import datetime
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Order & Scale | Stoplight Automator", layout="centered")

# --- HEADER SECTION ---
st.title("📊 Stoplight Report Generator")
st.markdown("##### *All data is processed in-memory and is automatically deleted once this tab is closed or refreshed.*")
st.write("Upload your Workamajig CSV and your previous Stoplight Report to begin.")

# --- SIDEBAR SETTINGS ---
with st.sidebar:
    st.header("Project Settings")
    client_input = st.text_input("Client Name:")
    pm_input = st.text_input("AM/PM Name:")

# --- FILE UPLOADERS ---
csv_file = st.file_uploader("1. Drop Workamajig CSV here", type=['csv'])
prev_report = st.file_uploader("2. Drop Previous Report (Excel)", type=['xlsx'])

# --- HELPER FUNCTIONS ---
def get_name_signature(name):
    """Normalizes names to match regardless of word order (e.g., 'Thomas A K' vs 'A K Thomas')"""
    if not name: return None
    return frozenset(str(name).replace('.', ' ').strip().title().split())

if st.button("Process & Sync Report"):
    if csv_file and prev_report:
        # Load workbook (data_only=False to keep formulas)
        wb = openpyxl.load_workbook(prev_report, data_only=False)
        
        # --- STOPLIGHT 3.9.1 CALIBRATION ---
        OMIT_PREFIX = "Yes-"
        START_ROW_OV = 10      
        PROJ_NAME_COL = 2      # Col B
        POP_COL = 3            # Col C
        AWARD_COL = 4          # Col D
        CURR_LABOR_COL = 5     # Col E
        PTD_LABOR_COL = 6      # Col F
        REMAINING_COL = 7      # Col G
        PERCENT_COL = 8        # Col H
        
        # 1. PROCESS TRANSACTIONS (Tab 1)
        ws_trans = wb.worksheets[0]
        trans_header_map = {str(cell.value).strip(): cell.column for cell in ws_trans[1] if cell.value}
        
        # Clear old transactions logic
        if ws_trans.max_row > 1:
            for row in ws_trans.iter_rows(min_row=2, max_row=ws_trans.max_row):
                for cell in row:
                    if not isinstance(cell, MergedCell): cell.value = None

        current_period_totals = defaultdict(float)
        project_date_ranges = {}
        unique_projects_in_csv = {} # Signature: Raw Name

        # Read CSV Data
        decoded_file = csv_file.getvalue().decode('utf-8').splitlines()
        reader = csv.DictReader(decoded_file)
        
        next_trans_row = 2
        for row in reader:
            full_name = row.get('Project Full Name', '').strip()
            if full_name and not full_name.startswith(OMIT_PREFIX):
                sig = get_name_signature(full_name)
                unique_projects_in_csv[sig] = full_name
                
                # Populate Transactions Tab mapping
                for col_name, col_idx in trans_header_map.items():
                    if col_name in row:
                        target_cell = ws_trans.cell(row=next_trans_row, column=col_idx)
                        if not isinstance(target_cell, MergedCell):
                            val = row[col_name]
                            if col_name in ['Quantity', 'Net', 'Gross']:
                                try: val = float(val.replace(',', ''))
                                except: pass
                            target_cell.value = val
                
                # Date and Labor logic
                raw_date = row.get('Expense Date', '')
                if raw_date:
                    try:
                        current_date = datetime.strptime(raw_date, '%m/%d/%Y')
                        if sig not in project_date_ranges:
                            project_date_ranges[sig] = [current_date, current_date]
                        else:
                            project_date_ranges[sig][0] = min(project_date_ranges[sig][0], current_date)
                            project_date_ranges[sig][1] = max(project_date_ranges[sig][1], current_date)
                    except ValueError: pass

                if row.get('Tran Type', '').strip().upper() == 'LABOR':
                    try: current_period_totals[sig] += float(row['Gross'].replace(',', ''))
                    except: pass
            next_trans_row += 1

        # 2. UPDATE ACCOUNT OVERVIEW (Tab 2)
        ws_ov = wb.worksheets[1] 
        
        # Formatting Fills
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        no_fill = PatternFill(fill_type=None)

        # Unmerge
        merged_ranges = list(ws_ov.merged_cells.ranges)
        for m_range in merged_ranges:
            if m_range.min_row >= START_ROW_OV:
                ws_ov.unmerge_cells(str(m_range))

        if client_input: ws_ov["E5"] = client_input
        if pm_input: ws_ov["I5"] = pm_input
        
        # Map existing project signatures to their row index
        excel_project_map = {}
        total_row_idx = None
        for r in range(START_ROW_OV, ws_ov.max_row + 1):
            cell_val = ws_ov.cell(row=r, column=PROJ_NAME_COL).value
            if cell_val and "TOTAL" in str(cell_val).upper():
                total_row_idx = r
                break
            if cell_val:
                excel_project_map[get_name_signature(cell_val)] = r

        total_row_idx = total_row_idx or (ws_ov.max_row + 1)

        # Update Existing & Identify Missing
        found_in_csv = set()
        for sig, excel_row in excel_project_map.items():
            if sig in unique_projects_in_csv:
                found_in_csv.add(sig)
                curr_labor = current_period_totals.get(sig, 0)
                
                # Update Period Labor (Col E)
                ws_ov.cell(row=excel_row, column=CURR_LABOR_COL, value=curr_labor)
                
                # Update PTD (Col F) - Increment existing value
                prev_ptd = ws_ov.cell(row=excel_row, column=PTD_LABOR_COL).value or 0
                try: 
                    if isinstance(prev_ptd, str): prev_ptd = float(prev_ptd.replace('$', '').replace(',', ''))
                except: prev_ptd = 0
                ws_ov.cell(row=excel_row, column=PTD_LABOR_COL, value=float(prev_ptd) + float(curr_labor))
                
                # Update Date Range
                if sig in project_date_ranges:
                    s, e = project_date_ranges[sig]
                    ws_ov.cell(row=excel_row, column=POP_COL, value=f"{s.strftime('%m/%d/%y')} - {e.strftime('%m/%d/%y')}")
                
                # Ensure formula exists in Col G
                ws_ov.cell(row=excel_row, column=REMAINING_COL, value=f"=D{excel_row}-F{excel_row}")
                ws_ov.cell(row=excel_row, column=1).fill = no_fill # Clear any old flags
            else:
                # Project in Excel but NOT in CSV (Flag Red)
                ws_ov.cell(row=excel_row, column=1).fill = red_fill

        # Add New Projects from CSV (Flag Yellow)
        new_projects_sigs = set(unique_projects_in_csv.keys()) - set(excel_project_map.keys())
        for sig in new_projects_sigs:
            ws_ov.insert_rows(total_row_idx)
            new_row = total_row_idx
            
            proj_name = unique_projects_in_csv[sig]
            curr_labor = current_period_totals.get(sig, 0)
            
            ws_ov.cell(row=new_row, column=PROJ_NAME_COL, value=proj_name)
            ws_ov.cell(row=new_row, column=1).fill = yellow_fill # New Project Flag
            ws_ov.cell(row=new_row, column=CURR_LABOR_COL, value=curr_labor)
            ws_ov.cell(row=new_row, column=PTD_LABOR_COL, value=curr_labor)
            
            # Formulas
            ws_ov.cell(row=new_row, column=REMAINING_COL, value=f"=D{new_row}-F{new_row}")
            ws_ov.cell(row=new_row, column=PERCENT_COL, value=f"=F{new_row}/D{new_row}")
            
            if sig in project_date_ranges:
                s, e = project_date_ranges[sig]
                ws_ov.cell(row=new_row, column=POP_COL, value=f"{s.strftime('%m/%d/%y')} - {e.strftime('%m/%d/%y')}")

            # Apply Styles from row above
            for c in range(1, 10):
                source = ws_ov.cell(row=START_ROW_OV, column=c)
                target = ws_ov.cell(row=new_row, column=c)
                if source.has_style:
                    target.font = copy.copy(source.font)
                    target.border = copy.copy(source.border)
                    target.alignment = copy.copy(source.alignment)
                    if c >= AWARD_COL and c <= REMAINING_COL:
                        target.number_format = '"$"#,##0.00'
            
            total_row_idx += 1

        # 3. EXPORT
        output = io.BytesIO()
        wb.save(output)
        st.success(f"Sync Complete! {len(unique_projects_in_csv)} Projects Sync'd.")
        st.download_button(label="💾 Download Final Report", data=output.getvalue(), file_name="Sync_Stoplight_Report.xlsx")
    else:
        st.error("Please provide both the CSV and the Excel report.")

st.markdown("---")
st.caption("🛡️ **Privacy Guard Active:** All session data is ephemeral.")
st.caption("📦 **Version:** 4.0.0")