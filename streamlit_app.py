import streamlit as st
import csv
import openpyxl
import io
import copy
from collections import defaultdict
from datetime import datetime
from openpyxl.cell.cell import MergedCell

st.set_page_config(page_title="Order & Scale | Stoplight Automator", layout="centered")

st.title("📊 Weekly Stoplight Report Generator")
st.write("Upload your Workamajig CSV and your previous Stoplight Report to update history.")

# --- SIDEBAR SETTINGS ---
with st.sidebar:
    st.header("Project Settings")
    client_input = st.text_input("Enter Client Name (Cell E5):")
    pm_input = st.text_input("Enter PM Name (Cell I5):")
    st.info("Updates header only if cells are currently empty.")

# --- FILE UPLOADERS ---
csv_file = st.file_uploader("1. Drop Workamajig CSV here", type=['csv'])
prev_report = st.file_uploader("2. Drop Previous Report (or Template) here", type=['xlsx'])

if st.button("Process & Sync Report"):
    if csv_file and prev_report:
        # Load workbook with formulas preserved
        wb = openpyxl.load_workbook(prev_report, data_only=False)
        
        # --- CONFIGURATION (Based on Stoplight Template) ---
        OMIT_PREFIX = "Yes-"
        START_ROW_OV = 10
        PROJ_NAME_COL = 2    # Col B
        POP_COL = 3          # Col C
        AWARD_COL = 4        # Col D
        CURR_LABOR_COL = 5   # Col E (Period Spend)
        PTD_LABOR_COL = 6    # Col F (Project to Date)
        REMAINING_COL = 7    # Col G (Formula: D-F)
        
        # 1. CLEAN TRANSACTIONS TAB
        ws_trans = wb["Transactions"] if "Transactions" in wb.sheetnames else wb.worksheets[0]
        trans_header_map = {str(cell.value).strip(): cell.column for cell in ws_trans[1] if cell.value}
        
        if ws_trans.max_row > 1:
            for row in ws_trans.iter_rows(min_row=2, max_row=ws_trans.max_row):
                for cell in row:
                    if not isinstance(cell, MergedCell): cell.value = None

        current_period_totals = defaultdict(float)
        project_date_ranges = {}
        unique_projects = set()

        # Read CSV Data
        decoded_file = csv_file.getvalue().decode('utf-8').splitlines()
        reader = csv.DictReader(decoded_file)
        
        next_trans_row = 2
        for row in reader:
            for col_name, col_idx in trans_header_map.items():
                if col_name in row:
                    target_cell = ws_trans.cell(row=next_trans_row, column=col_idx)
                    if not isinstance(target_cell, MergedCell):
                        val = row[col_name]
                        if col_name in ['Quantity', 'Net', 'Gross']:
                            try: val = float(val.replace(',', ''))
                            except: pass
                        target_cell.value = val
            
            full_name = row.get('Project Full Name', '').strip()
            if full_name and not full_name.startswith(OMIT_PREFIX):
                unique_projects.add(full_name)
                raw_date = row.get('Expense Date', '')
                if raw_date:
                    try:
                        current_date = datetime.strptime(raw_date, '%m/%d/%Y')
                        if full_name not in project_date_ranges:
                            project_date_ranges[full_name] = [current_date, current_date]
                        else:
                            project_date_ranges[full_name][0] = min(project_date_ranges[full_name][0], current_date)
                            project_date_ranges[full_name][1] = max(project_date_ranges[full_name][1], current_date)
                    except ValueError: pass

                if row.get('Tran Type', '').strip().upper() == 'LABOR':
                    try: current_period_totals[full_name] += float(row['Gross'].replace(',', ''))
                    except: pass
            next_trans_row += 1

        # 2. UPDATE ACCOUNT OVERVIEW TAB
        ws_ov = wb["Account Overview"] if "Account Overview" in wb.sheetnames else wb.worksheets[2]
        
        # Header Check
        if not ws_ov["E5"].value and client_input: ws_ov["E5"] = client_input
        if not ws_ov["I5"].value and pm_input: ws_ov["I5"] = pm_input
        
        # Find Footer/Total Row
        total_row_idx = None
        for r in range(START_ROW_OV, ws_ov.max_row + 1):
            if "Total" in str(ws_ov.cell(row=r, column=PROJ_NAME_COL).value):
                total_row_idx = r
                break
        total_row_idx = total_row_idx or (ws_ov.max_row + 1)

        # Get existing project map
        existing_rows = {}
        for r in range(START_ROW_OV, total_row_idx):
            name = ws_ov.cell(row=r, column=PROJ_NAME_COL).value
            if name: existing_rows[str(name).strip()] = r

        sorted_projects = sorted(list(unique_projects))
        
        # Ensure we have enough space
        required_rows = len(sorted_projects)
        available_slots = total_row_idx - START_ROW_OV
        if required_rows > available_slots:
            ws_ov.insert_rows(total_row_idx, amount=(required_rows - available_slots))

        # Update and Style
        for i, proj in enumerate(sorted_projects):
            row_idx = START_ROW_OV + i
            
            # --- DATA UPDATE ---
            ws_ov.cell(row=row_idx, column=PROJ_NAME_COL).value = proj
            
            if proj in project_date_ranges:
                s, e = project_date_ranges[proj]
                ws_ov.cell(row=row_idx, column=POP_COL).value = f"{s.strftime('%m/%d/%y')} - {e.strftime('%m/%d/%y')}"

            curr_labor = current_period_totals.get(proj, 0)
            
            # Read and Update PTD
            prev_ptd = ws_ov.cell(row=row_idx, column=PTD_LABOR_COL).value or 0
            try:
                if isinstance(prev_ptd, str):
                    prev_ptd = float(prev_ptd.replace('$', '').replace(',', ''))
            except:
                prev_ptd = 0
            
            ws_ov.cell(row=row_idx, column=CURR_LABOR_COL).value = curr_labor
            ws_ov.cell(row=row_idx, column=PTD_LABOR_COL).value = (float(prev_ptd) + float(curr_labor))
            
            # --- FORMULA REPAIR ---
            # Corrects Column G formula for the specific row: =D - F
            ws_ov.cell(row=row_idx, column=REMAINING_COL).value = f"=D{row_idx}-F{row_idx}"

            # --- STYLE CLONING (The Fix) ---
            # Clones everything from Column B to Column G from Row 10
            for c in range(PROJ_NAME_COL, REMAINING_COL + 1):
                source_cell = ws_ov.cell(row=START_ROW_OV, column=c)
                target_cell = ws_ov.cell(row=row_idx, column=c)
                if source_cell.has_style:
                    target_cell.font = copy.copy(source_cell.font)
                    target_cell.border = copy.copy(source_cell.border)
                    target_cell.fill = copy.copy(source_cell.fill)
                    target_cell.alignment = copy.copy(source_cell.alignment)
                    # Force currency format for financial columns
                    if c >= AWARD_COL:
                        target_cell.number_format = '"$"#,##0.00'

        # 3. EXPORT
        output = io.BytesIO()
        wb.save(output)
        st.success(f"Report Synced: {len(sorted_projects)} Projects Processed.")
        st.download_button(
            label="💾 Download Updated Stoplight Report",
            data=output.getvalue(),
            file_name="Updated_Stoplight_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please provide both the CSV and the Excel report.")

# --- FOOTER ---
st.markdown("---")
st.caption("📦 **Version:** 3.0.0 (Stoplight Engine)")
st.caption("🚀 **Deployed:** May 5, 2026")
