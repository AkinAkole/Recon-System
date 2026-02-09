import streamlit as st
import pandas as pd
import os
import re
import io
import base64
import plotly.express as px
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- APP START ---
st.set_page_config(page_title="Recon Tool", layout="wide")

# --- SECURITY GATEKEEPER ---
def check_password():
    """Returns True if the user had the correct password."""
    if st.session_state.get("password_correct", False):
        return True

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.title("Secure Access")
        user_pwd = st.text_input("Institutional Password", type="password")
        if st.button("Unlock System"):
            # Check if password exists in secrets first to avoid KeyError
            if "INSTITUTIONAL_PASSWORD" in st.secrets and user_pwd == st.secrets["INSTITUTIONAL_PASSWORD"]: 
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("🚫 Access Denied")
    return False

# ONLY PROCEED IF PASSWORD IS CORRECT
if check_password():
    # --- STYLING CONSTANTS ---
    NAVY_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    GREY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    WHITE_TEXT = Font(color="FFFFFF", bold=True)
    BLACK_BOLD = Font(color="000000", bold=True)
    RED_BOLD = Font(color="CC0000", bold=True)
    DOUBLE_BORDER = Border(top=Side(style='double'), bottom=Side(style='double'))
    THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    st.title("📂 Institutional Reconciliation System")

    # --- SIDEBAR TOOLS ---
    st.sidebar.header("🔍 Quick Search")
    search_ref = st.sidebar.text_input("Enter FCM Reference to Track")

    col1, col2 = st.columns(2)
    with col1:
        gl_uploads = st.file_uploader("Upload GL Excel Files", type=['xlsx', 'xls'], accept_multiple_files=True)
    with col2:
        csv_uploads = st.file_uploader("Upload NIBSS CSV Files", type=['csv'], accept_multiple_files=True)

    def heavy_clean_file(uploaded_file):
        try:
            content = uploaded_file.read()
            # Safety: reset file pointer if read elsewhere
            uploaded_file.seek(0)
            text = content.replace(b'\x00', b'').decode('utf-8', errors='ignore')
            if '\t' in text: text = text.replace(',', ';').replace('\t', ',')
            return io.StringIO(text)
        except: return None

    def clean_num(col):
        col_clean = col.astype(str).str.replace(r'[^-0-9.]', '', regex=True)
        return pd.to_numeric(col_clean, errors='coerce').fillna(0)

    if st.button("🚀 Run Reconciliation"):
        if gl_uploads and csv_uploads:
            run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # --- DATA LOADING & PROCESSING ---
            all_gl, all_csv, csv_dict = [], [], {}
            for f in gl_uploads:
                df = pd.read_excel(f)
                df.columns = [str(c).strip() for c in df.columns]
                if 'Description' not in df.columns:
                    st.error(f"❌ Error: File '{f.name}' missing 'Description'.")
                    st.stop()
                for c in ['Deposit', 'Withdrawal', 'Balance']:
                    if c in df.columns: df[c] = clean_num(df[c])
                all_gl.append(df)
            df_gl_input = pd.concat(all_gl, ignore_index=True) if all_gl else pd.DataFrame()

            for f in csv_uploads:
                buffer = heavy_clean_file(f)
                if buffer:
                    df = pd.read_csv(buffer, on_bad_lines='skip', engine='python')
                    df.columns = [str(c).strip().lower() for c in df.columns]
                    df.rename(columns={'uniquereference': 'unique_reference', 'remittedamount': 'remitted_amount', 'mdaname': 'mda_name'}, inplace=True)
                    for c in ['remitted_amount', 'collected_amount', 'fee']:
                        if c in df.columns: df[c] = clean_num(df[c])
                    df['source_file'] = f.name
                    if 'mda_name' in df.columns:
                        df['mda_name'] = df['mda_name'].astype(str).str.strip().str.upper()
                    all_csv.append(df)
                    csv_dict[f.name] = df
            df_csv_input = pd.concat(all_csv, ignore_index=True) if all_csv else pd.DataFrame()

            gl_review = df_gl_input.copy() if not df_gl_input.empty else pd.DataFrame(columns=['Description', 'Deposit', 'Withdrawal', 'Reference'])
            
            # GL_Reference logic
            if 'Reference' in gl_review.columns:
                gl_review['GL_Reference'] = gl_review['Reference'].astype(str).replace('nan', '')
            else:
                gl_review['GL_Reference'] = ""

            if 'Description' in gl_review.columns:
                gl_review['Reference'] = gl_review['Description'].astype(str).str.extract(r'(FCM\d{17})', expand=False).fillna("")
            else:
                gl_review['Reference'] = ""
            
            # Column Reordering Safety
            cols = list(gl_review.columns)
            if 'GL_Reference' in cols:
                cols.insert(2, cols.pop(cols.index('GL_Reference')))
                gl_review = gl_review[cols]
                
            csv_totals = df_csv_input.groupby('unique_reference')['remitted_amount'].sum().to_dict() if not df_csv_input.empty else {}
            gl_review['NIBSS_remitted'] = gl_review['Reference'].map(csv_totals).fillna(0)
            gl_review['NIBSS_reference'] = gl_review.apply(lambda x: x['Reference'] if x['Reference'] in csv_totals and x['Reference'] != "" else "", axis=1)
            gl_review['Variance'] = gl_review['NIBSS_remitted'] - gl_review['Deposit']

            gl_match_map = gl_review.groupby('Reference')['Deposit'].sum().to_dict()
            nibss_review = df_csv_input.copy() if not df_csv_input.empty else pd.DataFrame(columns=['unique_reference', 'remitted_amount', 'bank_id'])

            if not nibss_review.empty:
                b_idx = list(nibss_review.columns).index('bank_id') + 1 if 'bank_id' in nibss_review.columns else 1
                nibss_review.insert(b_idx, 'Kachasi_ref', nibss_review['unique_reference'].map(lambda x: x if x in gl_match_map else ""))
                nibss_review.insert(b_idx+1, 'Kachasi_In_GL', nibss_review['unique_reference'].map(gl_match_map).fillna(0))
                nibss_review.insert(b_idx+2, 'Settle_Variance', nibss_review['remitted_amount'] - nibss_review['Kachasi_In_GL'])

            # --- CALCULATIONS ---
            matched_mask_gl = (gl_review['NIBSS_reference'] != "")
            total_matched_gl_dep = gl_review[matched_mask_gl]['Deposit'].sum()
            total_matched_gl_nibss = gl_review[matched_mask_gl]['NIBSS_remitted'].sum()
            unmatched_gl_dep = gl_review[~matched_mask_gl & (gl_review['Deposit'] > 0)]['Deposit'].sum()
            bridging_diff = total_matched_gl_dep - total_matched_gl_nibss
            csv_vs_kachasi_diff = total_matched_gl_nibss - total_matched_gl_dep
            
            is_m = (nibss_review['Kachasi_ref'] != "") if 'Kachasi_ref' in nibss_review.columns else pd.Series([False]*len(nibss_review))
            is_c = ((nibss_review['unique_reference'].str.startswith('F', na=False)) & (nibss_review.get('mda_name', pd.Series()) == 'NIGERIA CUSTOM SERVICES'))
            
            unmatched_csv_customs = nibss_review[is_c & ~is_m]['remitted_amount'].sum() if not nibss_review.empty else 0
            unmatched_csv_others = nibss_review[~is_m & ~is_c]['remitted_amount'].sum() if not nibss_review.empty else 0

            # --- BROWSER DASHBOARD ---
            st.markdown("---")
            st.subheader("📊 Executive Insights")
            
            v1, v2 = st.columns([1, 1.5])
            with v1:
                if 'mda_name' in df_csv_input.columns and not df_csv_input.empty:
                    mda_share = df_csv_input.groupby('mda_name')['remitted_amount'].sum().reset_index()
                    mda_share = mda_share[mda_share['remitted_amount'] > 0]
                    if not mda_share.empty:
                        fig = px.pie(mda_share, values='remitted_amount', names='mda_name', hole=0.4, title="Remittance Market Share")
                        fig.update_layout(showlegend=False, height=350)
                        st.plotly_chart(fig, use_container_width=True)

            with v2:
                m1, m2, m3 = st.columns(3)
                m1.metric("Overall GL Deposit", f"₦{gl_review['Deposit'].sum():,.2f}")
                m2.metric("Overall NIBSS Remittance", f"₦{df_csv_input['remitted_amount'].sum() if not df_csv_input.empty else 0:,.2f}")
                m3.metric("Net System Variance", f"₦{gl_review['Variance'].sum():,.2f}")

            st.markdown("### Detailed Reconciliation Breakdown")
            d1, d2, d3 = st.columns(3)
            with d1:
                st.write("**Bridging (Excel to CSV)**")
                df_bridge = pd.DataFrame({"Description": ["Matched GL Deposit", "Matched NIBSS Remitted", "Difference", "Unmatched GL Dep"], "Value": [total_matched_gl_dep, total_matched_gl_nibss, bridging_diff, unmatched_gl_dep]}).set_index("Description")
                st.table(df_bridge.style.format("₦{:,.2f}"))
            with d2:
                st.write("**CSV to Excel Analysis**")
                df_comp = pd.DataFrame({"Description": ["Total Matched CSV", "Total Matched GL", "Difference"], "Value": [total_matched_gl_nibss, total_matched_gl_dep, csv_vs_kachasi_diff]}).set_index("Description")
                st.table(df_comp.style.format("₦{:,.2f}"))
            with d3:
                st.write("**Unmatched CSV Categorization**")
                st.table(pd.DataFrame({"Description": ["Customs (NCS)", "Other MDAs", "Total Unmatched"], "Value": [unmatched_csv_customs, unmatched_csv_others, unmatched_csv_customs + unmatched_csv_others]}).set_index("Description").style.format("₦{:,.2f}"))

            # --- EXCEL OUTPUT GENERATION ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # [Rest of your Excel logic remains the same...]
                summary_sections = [
                    ['EXECUTIVE RECONCILIATION DASHBOARD', ''],
                    ['Run Timestamp', run_time],
                    ['Overall GL Deposit', gl_review['Deposit'].sum()],
                    ['Overall NIBSS Remittance', df_csv_input['remitted_amount'].sum() if not df_csv_input.empty else 0],
                    ['Net System Variance', gl_review['Variance'].sum()]
                ]
                pd.DataFrame(summary_sections).to_excel(writer, sheet_name='Executive Summary', index=False, header=False)
                
                # ... (Keep your write_block and write_grand_total functions here)
                def write_block(ws, df, start_row, label, sum_cols):
                    if df.empty: return start_row
                    df_clean = df.fillna('')
                    df_clean.to_excel(writer, sheet_name=ws.title, startrow=start_row, index=False)
                    h_row = start_row + 1 
                    for r in range(h_row, h_row + len(df_clean) + 1):
                        for c in range(1, len(df_clean.columns) + 1):
                            cell = ws.cell(row=r, column=c)
                            cell.border = THIN_BORDER
                            if r == h_row:
                                cell.fill = NAVY_FILL; cell.font = WHITE_TEXT; cell.alignment = Alignment(horizontal='center')
                            else:
                                if r % 2 == 0: cell.fill = GREY_FILL
                                col_name = str(df_clean.columns[c-1]).lower()
                                if any(x in col_name for x in ['amount', 'fee', 'deposit', 'withdrawal', 'variance', 'remitted', 'in_gl']):
                                    cell.number_format = '#,##0.00'; cell.alignment = Alignment(horizontal='right')
                    t_row = h_row + len(df_clean) + 1
                    ws.cell(row=t_row, column=2, value=f"SUB-TOTAL: {label}").font = BLACK_BOLD
                    for i, col in enumerate(df_clean.columns, 1):
                        if col in sum_cols:
                            cell = ws.cell(row=t_row, column=i, value=df[col].sum())
                            cell.font = BLACK_BOLD; cell.number_format = '#,##0.00'
                    return t_row + 2

                def write_grand_total(ws, df, row, label, sum_cols):
                    ws.cell(row=row, column=2, value=label).font = RED_BOLD
                    for i, col in enumerate(df.columns, 1):
                        if col in sum_cols:
                            cell = ws.cell(row=row, column=i, value=df[col].sum())
                            cell.font = BLACK_BOLD; cell.border = DOUBLE_BORDER; cell.number_format = '#,##0.00'
                    return row + 2

                # Sheets
                ws_gl = writer.book.create_sheet('GL_Review')
                gl_sums = ['Deposit', 'Withdrawal', 'NIBSS_remitted', 'Variance']
                r = write_block(ws_gl, gl_review[gl_review['NIBSS_reference'] != ""], 0, "MATCHED GL", gl_sums)
                r = write_block(ws_gl, gl_review[(gl_review['NIBSS_reference'] == "") & (gl_review['Deposit'] > 0)], r, "UNMATCHED DEPOSIT", gl_sums)
                
                ws_nr = writer.book.create_sheet('NIBSS_Review')
                nr_sums = ['remitted_amount', 'collected_amount', 'fee', 'Kachasi_In_GL', 'Settle_Variance']
                r_n = write_block(ws_nr, nibss_review[is_m], 0, "MATCHED NIBSS", nr_sums)

                # MDA Extracts and log logic remains as you wrote...
                # (Omitted for brevity but included in your final logic)

            st.download_button(label="📥 Download Executive Report", data=output.getvalue(), file_name="Executive_Recon_Report.xlsx")
        else:
            st.error("Please upload both GL and NIBSS files.")
