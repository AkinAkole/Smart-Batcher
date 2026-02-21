import streamlit as st
import pandas as pd
import random
import re
import io
import plotly.express as px
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment

# --- 1. APP CONFIG & STYLES ---
st.set_page_config(page_title="Smart Batcher", page_icon="📊", layout="wide")
st.title("📊 Smart Batcher")

# Global Styles
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

# --- 2. SIDEBAR (Security & Search) ---
st.sidebar.header("🛡️ Security & Search")
enable_protection = st.sidebar.checkbox("Enable Password Protection", value=True)
custom_password = st.sidebar.text_input("Set Password", value="Smart_File_Lock", type="password") if enable_protection else "Smart_File_Lock"

# --- 3. MAIN TABS ---
tab1, tab2 = st.tabs(["🚀 Generator", "📖 User Manual"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        items_file = st.file_uploader("Upload Participants List (Class List)", type=['xlsx'])
    with col2:
        template_file = st.file_uploader("Upload Team Names (Template)", type=['xlsx'])

    if items_file and template_file:
        if st.button("🚀 Generate Distribution"):
            try:
                # LOAD DATA
                df_items = pd.read_excel(items_file)
                group_headers = sorted([c for c in df_items.columns if "Unnamed" not in str(c)])
                
                wb_tmpl = load_workbook(template_file, data_only=True)
                ws_tmpl = wb_tmpl.worksheets[0]
                first_col_map = {clean_val(ws_tmpl.cell(row=r, column=1).value): r 
                                 for r in range(1, ws_tmpl.max_row + 1) if ws_tmpl.cell(row=r, column=1).value}

                # WORKBOOK SETUP
                wb_out = Workbook()
                summary_sheet = wb_out.active
                summary_sheet.title = "Distribution Summary"
                
                headers = ["Group Header", "Status / Row", "Teams Count", "Total Items"]
                for idx, text in enumerate(headers, 1):
                    cell = summary_sheet.cell(row=4, column=idx, value=text)
                    cell.font, cell.fill, cell.alignment = summary_text_font, summary_header_fill, Alignment(horizontal="center")

                summary_row = 5
                browser_data = []
                total_teams_acc = 0
                total_items_acc = 0

                # CORE LOGIC LOOP
                for header in group_headers:
                    target = clean_val(header)
                    if target not in first_col_map:
                        browser_data.append({"Group": header, "Teams": 0, "Items": 0, "Status": "❌ Skipped"})
                        continue
                    
                    anchor_r = first_col_map[target]
                    items = df_items[header].dropna().apply(format_as_integer_string).tolist()
                    random.shuffle(items)

                    # Extract Matrix
                    matrix_blocks, curr_r, team_count = [], anchor_r, 0
                    while curr_r <= ws_tmpl.max_row:
                        row_teams = []
                        for c in range(2, ws_tmpl.max_column + 1, 2):
                            s_v, n_v = ws_tmpl.cell(row=curr_r, column=c).value, ws_tmpl.cell(row=curr_r, column=c+1).value
                            if s_v is not None and n_v is not None:
                                row_teams.append({'s': s_v, 'n': str(n_v).strip(), 'col': c})
                                team_count += 1
                        if not row_teams:
                            if matrix_blocks: break
                        else: matrix_blocks.append(row_teams)
                        curr_r += 1

                    if not items or not matrix_blocks:
                        browser_data.append({"Group": header, "Teams": team_count, "Items": len(items), "Status": "⚠️ Empty"})
                        continue

                    # Create Sheet
                    ws_out = wb_out.create_sheet(title=str(header)[:30])
                    if enable_protection: ws_out.protection.set_password(custom_password)

                    # Fill Summary
                    summary_sheet.cell(row=summary_row, column=1, value=header)
                    summary_sheet.cell(row=summary_row, column=3, value=team_count)
                    summary_sheet.cell(row=summary_row, column=4, value=len(items))
                    
                    total_teams_acc += team_count
                    total_items_acc += len(items)
                    summary_row += 1
                    browser_data.append({"Group": header, "Teams": team_count, "Items": len(items), "Status": "✅ Success"})

                    # Distribution & Auto-Width
                    max_widths = {}
                    flat_teams = [t for block in matrix_blocks for t in block]
                    avg, rem = divmod(len(items), len(flat_teams))
                    item_ptr, out_r = 1, 1

                    for block in matrix_blocks:
                        max_block_h = 0
                        for team in block:
                            col_let_data = ws_out.cell(row=1, column=team['col']+1).column_letter
                            max_widths[col_let_data] = max(max_widths.get(col_let_data, 0), len(str(team['n'])))
                            
                            c1, c2 = ws_out.cell(row=out_r, column=team['col'], value=team['s']), ws_out.cell(row=out_r, column=team['col']+1, value=team['n'])
                            for cell in [c1, c2]: cell.font, cell.fill = Font(bold=True), team_header_fill

                        for team in block:
                            t_idx = flat_teams.index(team)
                            count = avg + (1 if t_idx < rem else 0)
                            max_block_h = max(max_block_h, count)
                            for i in range(1, count + 1):
                                if item_ptr <= len(items):
                                    val = items[item_ptr-1]
                                    cell_v = ws_out.cell(row=out_r + i, column=team['col']+1, value=val)
                                    ws_out.cell(row=out_r+i, column=team['col'], value=i).fill = sn_column_fill
                                    col_let_v = ws_out.cell(row=1, column=team['col']+1).column_letter
                                    max_widths[col_let_v] = max(max_widths.get(col_let_v, 0), len(str(val)))
                                    item_ptr += 1
                        out_r += (max_block_h + 3)

                    for col_let, length in max_widths.items():
                        ws_out.column_dimensions[col_let].width = length + 3

                # Finalize Excel Summary Totals
                sum_label = summary_sheet.cell(row=summary_row, column=1, value="GRAND TOTAL")
                sum_teams = summary_sheet.cell(row=summary_row, column=3, value=total_teams_acc)
                sum_items = summary_sheet.cell(row=summary_row, column=4, value=total_items_acc)
                for c in [sum_label, sum_teams, sum_items]: c.font = Font(bold=True)

                # DASHBOARD
                st.divider()
                m1, m2, m3 = st.columns(3)
                m1.metric("Classes", len(group_headers))
                m2.metric("Total Items", total_items_acc)
                m3.metric("Total Teams", total_teams_acc)
                
                sum_df = pd.DataFrame(browser_data)
                st.dataframe(sum_df.style.background_gradient(cmap="Blues", subset=["Items"]), use_container_width=True, hide_index=True)

                # SAVE & DOWNLOAD
                output = io.BytesIO()
                wb_out.save(output)
                st.session_state['wb_final'] = wb_out # Store for Search
                st.download_button("📥 Download Excel", output.getvalue(), "Smart_Batcher_Results.xlsx")

            except Exception as e:
                st.error(f"Something went wrong: {e}")

with tab2:
    st.header("📘 How to use")
    st.write("1. Upload Participant List (Headers = Class Names)")
    st.write("2. Upload Template (Column A = Class Names, Column B+ = Teams)")

# SEARCH IN SIDEBAR (If data exists)
if 'wb_final' in st.session_state:
    st.sidebar.divider()
    search_q = st.sidebar.text_input("🔍 Search Name:").strip().lower()
    if search_q:
        for sn in st.session_state['wb_final'].sheetnames:
            if sn == "Distribution Summary": continue
            for row in st.session_state['wb_final'][sn].iter_rows(values_only=True):
                for val in row:
                    if val and search_q in str(val).lower():
                        st.sidebar.success(f"**Found!**\n\nName: {val}\n\nClass: {sn}")
