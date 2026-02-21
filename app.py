import streamlit as st
import pandas as pd
import random
import re
import io
import plotly.express as px
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment

# --- STREAMLIT UI SETUP ---
st.set_page_config(page_title="Smart Batcher", page_icon="📊", layout="wide")
st.title("📊 Smart Batcher - Easy Batching")

# --- SIDEBAR SETTINGS ---
st.sidebar.header("Security Settings")
enable_protection = st.sidebar.checkbox("Enable Password Protection", value=True)
custom_password = "Gemini2026"
if enable_protection:
    custom_password = st.sidebar.text_input("Set Sheet Password", value="Gemini2026", type="password")

# --- TABS ---
tab1, tab2 = st.tabs(["🚀 Generator", "📖 User Manual & Samples"])

with tab1:
    st.markdown("### Step 1: Upload your files")
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        items_file = st.file_uploader("Upload Participants List (Class List)", type=['xlsx'])
    with col_u2:
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
                
                # Setup Styles
                summary_header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
                summary_text_font = Font(color="FFFFFF", bold=True)
                team_header_fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")
                sn_column_fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")
                no_border = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))

                headers = ["Group Header", "Status / Row", "Teams Count", "Total Items"]
                for idx, text in enumerate(headers, 1):
                    cell = summary_sheet.cell(row=4, column=idx, value=text)
                    cell.font, cell.fill, cell.alignment = summary_text_font, summary_header_fill, Alignment(horizontal="center")

                summary_row, sheets_created = 5, 0
                browser_summary_data = []
                
                # Totals Trackers
                grand_total_teams = 0
                grand_total_items = 0

                # 3. Process Logic
                for header in group_headers:
                    target = clean_val(header)
                    if target not in first_col_map:
                        summary_sheet.cell(row=summary_row, column=1, value=header)
                        summary_sheet.cell(row=summary_row, column=2, value="Skipped: No match")
                        summary_row += 1
                        browser_summary_data.append({"Group": header, "Teams": 0, "Items": 0, "Status": "❌ Skipped"})
                        continue
                    
                    anchor_r = first_col_map[target]
                    items = df_items[header].dropna().apply(lambda x: re.sub(r'\.0$', '', str(x).strip())).tolist()
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
                        browser_summary_data.append({"Group": header, "Teams": team_count, "Items": len(items), "Status": "⚠️ Empty"})
                        continue

                    # Success Logic
                    ws_out = wb_out.create_sheet(title=str(header))
                    sheets_created += 1
                    if enable_protection: ws_out.protection.set_password(custom_password)

                    summary_sheet.cell(row=summary_row, column=1, value=header)
                    summary_sheet.cell(row=summary_row, column=2, value=f"Row {anchor_r}")
                    summary_sheet.cell(row=summary_row, column=3, value=team_count)
                    summary_sheet.cell(row=summary_row, column=4, value=len(items))
                    
                    # Accumulate Totals
                    grand_total_teams += team_count
                    grand_total_items += len(items)
                    
                    summary_row += 1
                    browser_summary_data.append({"Group": header, "Teams": team_count, "Items": len(items), "Status": "✅ Success"})

                    # ... [Internal Distribution Logic same as before] ...
                    # (Simplified for brevity in display, keep your existing logic here)

                # --- 4. ADD TOTALS ROW TO EXCEL SUMMARY ---
                total_label_cell = summary_sheet.cell(row=summary_row, column=1, value="GRAND TOTAL")
                total_teams_cell = summary_sheet.cell(row=summary_row, column=3, value=grand_total_teams)
                total_items_cell = summary_sheet.cell(row=summary_row, column=4, value=grand_total_items)
                
                for cell in [total_label_cell, total_teams_cell, total_items_cell]:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="ECF0F1", end_color="ECF0F1", fill_type="solid")

                # Final Excel Cleanup
                if enable_protection: summary_sheet.protection.set_password(custom_password)
                for col in ['A', 'B', 'C', 'D']: summary_sheet.column_dimensions[col].width = 25

                # --- 5. BROWSER DASHBOARD ---
                st.divider()
                sum_df = pd.DataFrame(browser_summary_data)
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Classes", len(sum_df))
                m2.metric("Total Items", grand_total_items)
                m3.metric("Total Teams", grand_total_teams)

                left_col, right_col = st.columns([1, 1])
                with left_col:
                    st.dataframe(sum_df.style.background_gradient(cmap="Blues", subset=["Items"]), use_container_width=True, hide_index=True)
                with right_col:
                    fig = px.pie(sum_df[sum_df["Items"] > 0], values='Items', names='Group', hole=0.4)
                    st.plotly_chart(fig, use_container_width=True)

                output = io.BytesIO()
                wb_out.save(output)
                st.download_button("📥 Download Final Output", output.getvalue(), "Smart_Batcher_Results.xlsx")

            except Exception as e:
                st.error(f"Error: {e}")
