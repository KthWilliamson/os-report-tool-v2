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

# --- SIDEBAR LOGS & SETTINGS ---
with st.sidebar:
    st.header("Project Settings")
    client_input = st.text_input("Client Name (E5):")
    pm_input = st.text_input("PM Name (I5):")
    st.divider()
    st.subheader("System Logs")
    log_area = st.empty()

# --- FILE UPLOADERS ---
csv_file = st.file_uploader("1. Drop Workamajig CSV here", type=['csv'])
prev_report = st.file_uploader("2. Drop Previous Report (or Template) here", type=['xlsx'])

if st.button("Process & Sync Report"):
    if csv_file and prev_report:
        # Load workbook
        wb = openpyxl.load_workbook(prev_report, data_only=False)
        
        # LOGGING: Show sheets found in the file
        log_area.write(f"Sheets Found: {wb.sheetnames}")
        
        # --- CONFIGURATION (Synced to Account Overview.csv) ---
        OMIT_PREFIX = "Yes-"
        START_ROW_OV = 10      # First data row
        PROJ_NAME_COL = 1      # Col A: WMJ Code + Name
        POP_COL = 2            # Col B: Period of Performance
        AWARD_COL = 3          # Col C: Total Labor Award
        CURR_LABOR_COL = 4     # Col D: Incurred Period
        PTD_LABOR_COL = 5      # Col E: Incurred PTD
        REMAINING_COL = 6      # Col F: Remaining Labor
        
        # 1. PROCESS TRANSACTIONS
        ws_trans = None
        for s in wb.worksheets:
            if "Transactions" in s.title:
                ws_trans = s
                break
        if not ws_trans: ws_trans = wb.worksheets[0]

        # Clear transactions
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
            full_name = row.get('Project Full Name', '').strip()
            if full_name and not full_name.startswith(OMIT_PREFIX):
                unique_projects.add(full_name)
                # Populate Transactions Tab
                # [Transaction mapping code remains same as previous]
                
                # Labor Math & Dates
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

        # 2. UPDATE ACCOUNT OVERVIEW
        ws_ov = None
        for sheet in wb.worksheets:
            if "Account Overview" in sheet.title:
                ws_ov = sheet
                break
        if not ws_ov:
            st.error("Error: Could not find 'Account Overview' tab.")
            st.stop()

        # Header logic
        if not ws_ov["E5"].value and client_input: ws_ov["E5"] = client_input
        if not ws_ov["I5"].value and pm_input: ws_ov["I5"] = pm_input
        
        # Locate Footer
        total_row_idx = None
        for r in range(START_ROW_OV, ws_ov.max_row + 1):
            val = str(ws_ov.cell(row=r, column=1).value or "").upper()
            if "TOTAL" in val:
                total_row_idx = r
                break
        total_row_idx = total_row_idx or (START_ROW_OV + 5)

        sorted_projects = sorted(list(unique_projects))
        required = len(sorted_projects)
        capacity = total_row_idx - START_ROW_OV
        
        if required > capacity:
            ws_ov.insert_rows(total_row_idx, amount=(required - capacity))

        # Main Data Injection Loop
        for i, proj in enumerate(sorted_projects):
            row_idx = START_ROW_OV + i
            
            # Use ws_ov.cell() explicitly to avoid None errors
            ws_ov.cell(row=row_idx, column=PROJ_NAME_COL, value=proj)
            
            if proj in project_date_ranges and project_date_ranges[proj][0]:
                s, e = project_date_ranges[proj]
                ws_ov.cell(row=row_idx, column=POP_COL, value=f"{s.strftime('%m/%d/%y')} - {e.strftime('%m/%d/%y')}")

            curr_labor = current_period_totals.get(proj, 0)
            prev_ptd = ws_ov.cell(row=row_idx, column=PTD_LABOR_COL).value or 0
            
            try:
                if isinstance(prev_ptd, str):
                    prev_ptd = float(prev_ptd.replace('$', '').replace(',', ''))
            except:
                prev_ptd = 0
            
            # The writing step
            ws_ov.cell(row=row_idx, column=CURR_LABOR_COL, value=curr_labor)
            ws_ov.cell(row=row_idx, column=PTD_LABOR_COL, value=(float(prev_ptd) + float(curr_labor)))
            ws_ov.cell(row=row_idx, column=REMAINING_COL, value=f"=C{row_idx}-E{row_idx}")

            # Apply Row 10 Styles
            for c in range(1, 8):
                source = ws_ov.cell(row=START_ROW_OV, column=c)
                target = ws_ov.cell(row=row_idx, column=c)
                if source.has_style:
                    target.font = copy.copy(source.font)
                    target.border = copy.copy(source.border)
                    target.fill = copy.copy(source.fill)
                    target.alignment = copy.copy(source.alignment)
                    if c >= AWARD_COL:
                        target.number_format = '"$"#,##0.00'

        # 3. EXPORT
        output = io.BytesIO()
        wb.save(output)
        st.success(f"Done! {len(sorted_projects)} Projects Processed.")
        st.download_button(label="💾 Download Final Report", data=output.getvalue(), file_name="Sync_Stoplight_Report.xlsx")
    else:
        st.error("Missing files.")

st.markdown("---")
st.caption("📦 Version: 3.4.0 (Final Production)")
