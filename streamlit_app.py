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

# --- SIDEBAR SETTINGS ---
with st.sidebar:
    st.header("Project Settings")
    client_input = st.text_input("Client Name (E5):")
    pm_input = st.text_input("PM Name (I5):")

# --- FILE UPLOADERS ---
csv_file = st.file_uploader("1. Drop Workamajig CSV here", type=['csv'])
prev_report = st.file_uploader("2. Drop Previous Report (Excel)", type=['xlsx'])

if st.button("Process & Sync Report"):
    if csv_file and prev_report:
        wb = openpyxl.load_workbook(prev_report, data_only=False)
        
        # --- STOPLIGHT 3.8 CALIBRATION ---
        OMIT_PREFIX = "Yes-"
        START_ROW_OV = 10      
        PROJ_NAME_COL = 2      # Col B
        POP_COL = 3            # Col C
        AWARD_COL = 4          # Col D
        CURR_LABOR_COL = 5     # Col E
        PTD_LABOR_COL = 6      # Col F
        REMAINING_COL = 7      # Col G
        
        # 1. PROCESS TRANSACTIONS (Tab 1)
        ws_trans = wb.worksheets[0]
        trans_header_map = {str(cell.value).strip(): cell.column for cell in ws_trans[1] if cell.value}
        
        if ws_trans.max_row > 1:
            for row in ws_trans.iter_rows(min_row=2, max_row=ws_trans.max_row):
                for cell in row:
                    if not isinstance(cell, MergedCell): cell.value = None

        current_period_totals = defaultdict(float)
        project_date_ranges = {}
        unique_projects = set()

        decoded_file = csv_file.getvalue().decode('utf-8').splitlines()
        reader = csv.DictReader(decoded_file)
        
        for row in reader:
            full_name = row.get('Project Full Name', '').strip()
            if full_name and not full_name.startswith(OMIT_PREFIX):
                unique_projects.add(full_name)
                # [Transaction mapping remains identical]
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

        # 2. UPDATE ACCOUNT OVERVIEW (Tab 2)
        ws_ov = wb.worksheets[1] 
        
        # --- NEW: UNMERGE LOGIC ---
        # We unmerge any cells in the data range (Col A-I) starting from START_ROW_OV
        merged_ranges = list(ws_ov.merged_cells.ranges)
        for m_range in merged_ranges:
            if m_range.min_row >= START_ROW_OV:
                ws_ov.unmerge_cells(str(m_range))

        if not ws_ov["E5"].value and client_input: ws_ov["E5"] = client_input
        if not ws_ov["I5"].value and pm_input: ws_ov["I5"] = pm_input
        
        total_row_idx = None
        for r in range(START_ROW_OV, ws_ov.max_row + 1):
            val = str(ws_ov.cell(row=r, column=PROJ_NAME_COL).value or "").upper()
            if "TOTAL" in val:
                total_row_idx = r
                break
        total_row_idx = total_row_idx or (START_ROW_OV + 5)

        sorted_projects = sorted(list(unique_projects))
        required = len(sorted_projects)
        capacity = total_row_idx - START_ROW_OV
        
        if required > capacity:
            ws_ov.insert_rows(total_row_idx, amount=(required - capacity))

        for i, proj in enumerate(sorted_projects):
            row_idx = START_ROW_OV + i
            ws_ov.cell(row=row_idx, column=PROJ_NAME_COL, value=proj)
            
            if proj in project_date_ranges and project_date_ranges[proj][0]:
                s, e = project_date_ranges[proj]
                ws_ov.cell(row=row_idx, column=POP_COL, value=f"{s.strftime('%m/%d/%y')} - {e.strftime('%m/%d/%y')}")
            else:
                ws_ov.cell(row=row_idx, column=POP_COL, value="TBD")

            curr_labor = current_period_totals.get(proj, 0)
            prev_ptd = ws_ov.cell(row=row_idx, column=PTD_LABOR_COL).value or 0
            try:
                if isinstance(prev_ptd, str):
                    prev_ptd = float(prev_ptd.replace('$', '').replace(',', ''))
            except:
                prev_ptd = 0
            
            ws_ov.cell(row=row_idx, column=CURR_LABOR_COL, value=curr_labor)
            ws_ov.cell(row=row_idx, column=PTD_LABOR_COL, value=(float(prev_ptd) + float(curr_labor)))
            ws_ov.cell(row=row_idx, column=REMAINING_COL).value = f"=D{row_idx}-F{row_idx}"

            # Formatting Sync
            for c in range(1, 10):
                source = ws_ov.cell(row=START_ROW_OV, column=c)
                target = ws_ov.cell(row=row_idx, column=c)
                if source.has_style:
                    target.font = copy.copy(source.font)
                    target.border = copy.copy(source.border)
                    target.fill = copy.copy(source.fill)
                    target.alignment = copy.copy(source.alignment)
                    if c >= AWARD_COL and c <= REMAINING_COL:
                        target.number_format = '"$"#,##0.00'

        # 3. EXPORT
        output = io.BytesIO()
        wb.save(output)
        st.success(f"Processed {len(sorted_projects)} projects with unmerged formatting.")
        st.download_button(label="💾 Download Final Report", data=output.getvalue(), file_name="Sync_Stoplight_Report.xlsx")
