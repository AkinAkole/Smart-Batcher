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

# Create Tabs
tab1, tab2 = st.tabs(["🚀 Generator", "📖 User Manual"])

with tab1:
    # --- YOUR EXISTING GENERATOR CODE GOES HERE ---
    st.markdown("Upload your files below to generate the distribution.")
    # (Move your columns, uploaders, and buttons inside this 'with tab1' block)

with tab2:
    st.header("📘 Smart Batcher User Manual")
    
    st.subheader("1. How it Works")
    st.write("""
    Smart Batcher automates the randomization of participants into team structures 
    defined by your template. It ensures no team is over-filled and shuffles 
    participants to maintain fairness.
    """)

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.info("**File A: Participants List**")
        st.write("""
        - **Format:** Excel (.xlsx)
        - **Headers:** Top row must be Class Names (e.g., *Class A*, *Class B*).
        - **Data:** List names directly under the headers.
        """)
        
    with col_m2:
        st.info("**File B: Team Template**")
        st.write("""
        - **Format:** Excel (.xlsx)
        - **Column A:** Must contain the Class Name (must match File A).
        - **Columns B, C, etc:** Contains the Team Names for that class.
        """)

    st.divider()
    
    st.subheader("2. Step-by-Step Guide")
    st.markdown("""
    1. **Set Security:** Use the sidebar to enable/disable password protection.
    2. **Upload Files:** Drag and drop your Participants List and Team Names.
    3. **Generate:** Click the 'Generate' button to see the dashboard.
    4. **Download:** Save your final report via the download button.
    """)

    # Optional: Keep the PDF download button here too!
    try:
        with open("User_Manual.pdf", "rb") as f:
            st.download_button("📥 Download PDF version", f, "Manual.pdf")
    except:
        st.caption("PDF version currently unavailable.")
# --- SIDEBAR SETTINGS ---
st.sidebar.header("Security Settings")
enable_protection = st.sidebar.checkbox("Enable Password Protection", value=True)
custom_password = "Smart_File_Lock"
if enable_protection:
    custom_password = st.sidebar.text_input("Set Sheet Password", value="Smart_File_Lock", type="password")

#st.markdown("Upload your files below to generate the distribution.")

# --- SETTINGS & STYLES ---
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
            
            if enable_protection:
                summary_sheet.cell(row=2, column=1, value="Note: Sheets are password protected.").font = Font(size=9, color="FF0000")
            
            headers = ["Group Header", "Status / Row", "Teams Count", "Total Items"]
            for idx, text in enumerate(headers, 1):
                cell = summary_sheet.cell(row=4, column=idx, value=text)
                cell.font, cell.fill, cell.alignment = summary_text_font, summary_header_fill, Alignment(horizontal="center")

            summary_row, sheets_created = 5, 0
            browser_summary_data = []

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
                    browser_summary_data.append({"Group": header, "Teams": team_count, "Items": len(items), "Status": "⚠️ Empty"})
                    continue

                ws_out = wb_out.create_sheet(title=str(header))
                sheets_created += 1
                
                # --- CONDITIONAL PROTECTION ---
                if enable_protection:
                    ws_out.protection.set_password(custom_password)

                summary_sheet.cell(row=summary_row, column=1, value=header)
                summary_sheet.cell(row=summary_row, column=2, value=f"Row {anchor_r}")
                summary_sheet.cell(row=summary_row, column=3, value=team_count)
                summary_sheet.cell(row=summary_row, column=4, value=len(items))
                summary_row += 1
                
                browser_summary_data.append({"Group": header, "Teams": team_count, "Items": len(items), "Status": "✅ Success"})

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
            if enable_protection:
                summary_sheet.protection.set_password(custom_password)
            for col in ['A', 'B', 'C', 'D']: summary_sheet.column_dimensions[col].width = 25
            for row in summary_sheet.iter_rows():
                for cell in row: cell.border = no_border

            # --- STREAMLIT DASHBOARD DISPLAY ---
            st.divider()
            sum_df = pd.DataFrame(browser_summary_data)
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Classes", len(sum_df))
            m2.metric("Total Items Distributed", sum_df["Items"].sum())
            m3.metric("Total Teams Formed", sum_df["Teams"].sum())

            left_col, right_col = st.columns([1, 1])
            with left_col:
                st.subheader("📋 Browser Summary")
                st.dataframe(
                    sum_df.style.set_properties(**{'background-color': '#F9F9F9', 'color': '#2C3E50'})
                    .set_table_styles([{'selector': 'th', 'props': [('background-color', '#2C3E50'), ('color', 'white')]}])
                    .background_gradient(cmap="Blues", subset=["Items"]),
                    use_container_width=True, hide_index=True
                )

            with right_col:
                st.subheader("🥧 Distribution Composition")
                fig = px.pie(sum_df[sum_df["Items"] > 0], values='Items', names='Group', 
                             hole=0.4, color_discrete_sequence=px.colors.qualitative.Prism)
                fig.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)

            output = io.BytesIO()
            wb_out.save(output)
            st.success(f"✅ Distribution Ready!")
            st.download_button(
                label="📥 Download Smart Batcher Excel",
                data=output.getvalue(),
                file_name=f"Smart_Batcher_{datetime.now().strftime('%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")




