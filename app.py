import streamlit as st
import pandas as pd
import random
import re
import io
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment

# --- STREAMLIT UI SETUP ---
st.set_page_config(page_title="Matrix Distributer", page_icon="📊")
st.title("📊 Sequential Matrix Distribution")
st.markdown("Upload your files below to generate the distribution.")

# --- SETTINGS & STYLES ---
SHEET_PASSWORD = "Gemini2026"
summary_header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
summary_text_font = Font(color="FFFFFF", bold=True)
team_header_fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")
sn_column_fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")
no_border = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))

def clean_val(v):
    return re.sub(r'[^a-zA-Z0-9]', '', str(v)).lower() if v else ""

def format_as_integer_string(v):
    if v is None: return ""
    s = str(v).strip()
    return s[:-2] if s.endswith('.0') else s

# --- FILE UPLOADERS ---
col1, col2 = st.columns(2)
with col1:
    items_file = st.file_uploader("Upload Participants List (Class List)", type=['xlsx'])
with col2:
    template_file = st.file_uploader("Upload Team Names (With Class Label)", type=['xlsx'])

if items_file and template_file:
    if st.button("🚀 Generate Distribution"):
        try:
            # 1. Load Data
            df_items = pd.read_excel(items_file)
            group_headers = sorted([c for c in df_items.columns if "Unnamed" not in str(c)])
            
            wb_tmpl = load_workbook(template_file, data_only=True)
            ws_tmpl = wb_tmpl.worksheets[0]
            first_col_map = {clean_val(ws_tmpl.cell(row=r, column=1).value): r 
                             for r in range(1, ws_tmpl.max_row + 1) if ws_tmpl.cell(row=r, column=1).value}

            # 2. Setup Workbook
            wb_out = Workbook()
            summary_sheet = wb_out.active
            summary_sheet.title = "Distribution Summary"
            
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            summary_sheet.cell(row=1, column=1, value=f"Generated on: {now}").font = Font(italic=True)
            summary_sheet.cell(row=2, column=1, value="Note: Sheets are password protected.").font = Font(size=9, color="FF0000")
            
            headers = ["Group Header", "Status / Row", "Teams Count", "Total Items"]
            for idx, text in enumerate(headers, 1):
                cell = summary_sheet.cell(row=4, column=idx, value=text)
                cell.font, cell.fill, cell.alignment = summary_text_font, summary_header_fill, Alignment(horizontal="center")

            summary_row, sheets_created = 5, 0

            # 3. Process Logic (Unchanged core logic)
            for header in group_headers:
                target = clean_val(header)
                if target not in first_col_map:
                    summary_sheet.cell(row=summary_row, column=1, value=header)
                    summary_sheet.cell(row=summary_row, column=2, value="Skipped: No match")
                    summary_row += 1
                    continue
                
                anchor_r = first_col_map[target]
                items = df_items[header].dropna().apply(format_as_integer_string).tolist()
                random.shuffle(items)

                matrix_blocks, curr_r, team_count = [], anchor_r, 0
                while curr_r <= ws_tmpl.max_row:
                    row_teams = []
                    for c in range(2, ws_tmpl.max_column + 1, 2):
                        s_val, n_val = ws_tmpl.cell(row=curr_r, column=c).value, ws_tmpl.cell(row=curr_r, column=c+1).value
                        if s_val is not None and n_val is not None:
                            row_teams.append({'s': s_val, 'n': str(n_val).strip(), 'col': c})
                            team_count += 1
                    if not row_teams:
                        if matrix_blocks: break
                    else: matrix_blocks.append(row_teams)
                    curr_r += 1

                if not items or not matrix_blocks:
                    summary_sheet.cell(row=summary_row, column=1, value=header)
                    summary_sheet.cell(row=summary_row, column=2, value=f"Row {anchor_r} (Missing Matrix)")
                    summary_row += 1
                    continue

                ws_out = wb_out.create_sheet(title=str(header))
                sheets_created += 1
                ws_out.protection.set_password(SHEET_PASSWORD)

                summary_sheet.cell(row=summary_row, column=1, value=header)
                summary_sheet.cell(row=summary_row, column=2, value=f"Row {anchor_r}")
                summary_sheet.cell(row=summary_row, column=3, value=team_count)
                summary_sheet.cell(row=summary_row, column=4, value=len(items))
                summary_row += 1

                max_widths, sn_width = {}, 11.9
                ws_out.column_dimensions['A'].width = sn_width
                flat_teams = [t for block in matrix_blocks for t in block]
                avg, rem = divmod(len(items), len(flat_teams))
                item_ptr, out_r = 1, 1

                for block in matrix_blocks:
                    max_block_h = 0
                    for team in block:
                        col_let_sn = ws_out.cell(row=1, column=team['col']).column_letter
                        ws_out.column_dimensions[col_let_sn].width = sn_width
                        c1, c2 = ws_out.cell(row=out_r, column=team['col'], value=team['s']), ws_out.cell(row=out_r, column=team['col']+1, value=team['n'])
                        for cell in [c1, c2]:
                            cell.font, cell.fill, cell.border = Font(bold=True), team_header_fill, no_border
                        
                        col_let_data = ws_out.cell(row=1, column=team['col']+1).column_letter
                        max_widths[col_let_data] = max(max_widths.get(col_let_data, 0), len(str(team['n'])))

                    for team in block:
                        t_idx = flat_teams.index(team)
                        count = avg + (1 if t_idx < rem else 0)
                        max_block_h = max(max_block_h, count)
                        for i in range(1, count + 1):
                            if item_ptr <= len(items):
                                val = items[item_ptr-1]
                                cell_idx, cell_val = ws_out.cell(row=out_r + i, column=team['col'], value=i), ws_out.cell(row=out_r + i, column=team['col']+1, value=val)
                                cell_idx.fill, cell_idx.border, cell_val.border = sn_column_fill, no_border, no_border
                                cell_val.number_format = '@'
                                col_let_data = ws_out.cell(row=1, column=team['col']+1).column_letter
                                max_widths[col_let_data] = max(max_widths.get(col_let_data, 0), len(str(val)))
                                item_ptr += 1
                    out_r += (max_block_h + 3)

                for col_let, length in max_widths.items():
                    ws_out.column_dimensions[col_let].width = length + 2

            # Final Summary Cleanup
            summary_sheet.protection.set_password(SHEET_PASSWORD)
            for col in ['A', 'B', 'C', 'D']: summary_sheet.column_dimensions[col].width = 25
            for row in summary_sheet.iter_rows():
                for cell in row: cell.border = no_border

            # --- PREPARE DOWNLOAD ---
            output = io.BytesIO()
            wb_out.save(output)
            st.success(f"✅ Success! Created {sheets_created} distribution sheets.")
            st.download_button(
                label="📥 Download Excel File",
                data=output.getvalue(),
                file_name="Sequential_Matrix_Distribution.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:

            st.error(f"An error occurred: {e}")
