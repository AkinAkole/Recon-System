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
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if st.session_state["password_correct"]:
        return True

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.title("Secure Access")
        user_pwd = st.text_input("Institutional Password", type="password")
        if "INSTITUTIONAL_PASSWORD" in st.secrets:
            if st.button("Unlock System"):
                if user_pwd == st.secrets["INSTITUTIONAL_PASSWORD"]: 
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("🚫 Access Denied")
        else:
            st.warning("Please set INSTITUTIONAL_PASSWORD in secrets.")
    return False

# Execution Gate
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
    search_ref = st.sidebar.text_input("Enter FCM Reference to Track").strip()

    col1, col2 = st.columns(2)
    with col1:
        gl_uploads = st.file_uploader("Upload GL Excel Files", type=['xlsx', 'xls'], accept_multiple_files=True)
    with col2:
        csv_uploads = st.file_uploader("Upload NIBSS CSV Files", type=['csv'], accept_multiple_files=True)

    def heavy_clean_file(uploaded_file):
        try:
            uploaded_file.seek(0)
            content = uploaded_file.read()
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
            
            # --- DATA LOADING ---
            all_gl, all_csv, csv_dict = [], [], {}
            for f in gl_uploads:
                df = pd.read_excel(f)
                df.columns = [str(c).strip() for c in df.columns]
                if 'Description' not in df.columns:
                    st.error(f"❌ Error: File '{f.name}' missing 'Description'.")
                    st.stop()
                for c in ['Deposit', 'Withdrawal', 'Balance']:
                    if c in df.columns: df[c] = clean_num(df[c])
                df['Source_File'] = f.name
                all_gl.append(df)
            df_gl_input = pd.concat(all_gl, ignore_index=True)

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
            df_csv_input = pd.concat(all_csv, ignore_index=True)

            # --- SEARCH FEATURE LOGIC ---
            if search_ref:
                st.markdown(f"### 🎯 Search Results for: `{search_ref}`")
                sc1, sc2 = st.columns(2)
                with sc1:
                    gl_match = df_gl_input[df_gl_input.astype(str).apply(lambda x: x.str.contains(search_ref, case=False)).any(axis=1)]
                    st.write("**Found in GL:**")
                    st.dataframe(gl_match)
                with sc2:
                    csv_match = df_csv_input[df_csv_input.astype(str).apply(lambda x: x.str.contains(search_ref, case=False)).any(axis=1)]
                    st.write("**Found in NIBSS:**")
                    st.dataframe(csv_match)

            # --- PROCESSING ---
            gl_review = df_gl_input.copy()
            gl_review['Reference'] = gl_review['Description'].astype(str).str.extract(r'(FCM\d{17})', expand=False).fillna("")
            
            csv_totals = df_csv_input.groupby('unique_reference')['remitted_amount'].sum().to_dict()
            gl_review['NIBSS_remitted'] = gl_review['Reference'].map(csv_totals).fillna(0)
            gl_review['NIBSS_reference'] = gl_review.apply(lambda x: x['Reference'] if x['Reference'] in csv_totals and x['Reference'] != "" else "", axis=1)
            gl_review['Variance'] = gl_review['NIBSS_remitted'] - gl_review['Deposit']

            gl_match_map = gl_review.groupby('Reference')['Deposit'].sum().to_dict()
            nibss_review = df_csv_input.copy()
            if 'bank_id' in nibss_review.columns:
                b_idx = list(nibss_review.columns).index('bank_id') + 1
                nibss_review.insert(b_idx, 'Kachasi_ref', nibss_review['unique_reference'].map(lambda x: x if x in gl_match_map else ""))
                nibss_review.insert(b_idx+1, 'Kachasi_In_GL', nibss_review['unique_reference'].map(gl_match_map).fillna(0))
                nibss_review.insert(b_idx+2, 'Settle_Variance', nibss_review['remitted_amount'] - nibss_review['Kachasi_In_GL'])

            # --- DASHBOARD CALCULATIONS ---
            matched_mask_gl = (gl_review['NIBSS_reference'] != "")
            total_matched_gl_dep = gl_review[matched_mask_gl]['Deposit'].sum()
            total_matched_gl_nibss = gl_review[matched_mask_gl]['NIBSS_remitted'].sum()
            unmatched_gl_dep = gl_review[~matched_mask_gl & (gl_review['Deposit'] > 0)]['Deposit'].sum()
            bridging_diff = total_matched_gl_dep - total_matched_gl_nibss
            csv_vs_kachasi_diff = total_matched_gl_nibss - total_matched_gl_dep
            
            is_m = (nibss_review['Kachasi_ref'] != "") if 'Kachasi_ref' in nibss_review.columns else pd.Series([False]*len(nibss_review))
            mda_check_col = nibss_review['mda_name'] if 'mda_name' in nibss_review.columns else pd.Series([""]*len(nibss_review))
            is_c = ((nibss_review['Kachasi_ref'] == "") & (mda_check_col == 'NIGERIA CUSTOM SERVICES') & (nibss_review['unique_reference'].str.startswith('F', na=False)))
            
            unmatched_csv_customs = nibss_review[is_c]['remitted_amount'].sum()
            unmatched_csv_others = nibss_review[~is_m & ~is_c]['remitted_amount'].sum()

            # --- DASHBOARD ---
            st.markdown("---")
            st.subheader("📊 Executive Insights")
            v1, v2 = st.columns([1, 1.5])
            with v1:
                mda_share = df_csv_input.groupby('mda_name')['remitted_amount'].sum().reset_index()
                fig = px.pie(mda_share[mda_share['remitted_amount']>0], values='remitted_amount', names='mda_name', hole=0.4, title="Market Share")
                fig.update_layout(showlegend=False, height=350)
                st.plotly_chart(fig, use_container_width=True)
            with v2:
                m1, m2, m3 = st.columns(3)
                m1.metric("Overall GL Deposit", f"₦{gl_review['Deposit'].sum():,.2f}")
                m2.metric("Overall NIBSS Remittance", f"₦{df_csv_input['remitted_amount'].sum():,.2f}")
                m3.metric("Net System Variance", f"₦{gl_review['Variance'].sum():,.2f}")

            # --- COLORFUL TABLES ---
            st.markdown("### Detailed Reconciliation Breakdown")
            def color_diff_row(val):
                return 'color: red; font-weight: bold' if val != 0 else ''

            d1, d2, d3 = st.columns(3)
            with d1:
                st.write("**Bridging (Excel to CSV)**")
                df_bridge = pd.DataFrame({"Description": ["Matched GL Deposit", "Matched NIBSS Remitted", "Difference (Matched GL vs NIBSS)", "Unmatched GL Dep"], "Value": [total_matched_gl_dep, total_matched_gl_nibss, bridging_diff, unmatched_gl_dep]}).set_index("Description")
                st.table(df_bridge.style.format("₦{:,.2f}").applymap(color_diff_row, subset=pd.IndexSlice[['Difference (Matched GL vs NIBSS)'], :]))
            with d2:
                st.write("**CSV to Excel Analysis**")
                df_comp = pd.DataFrame({"Description": ["Total Matched CSV Remittance", "Total Matched Kachasi Credit", "Difference (CSV vs Kachasi)"], "Value": [total_matched_gl_nibss, total_matched_gl_dep, csv_vs_kachasi_diff]}).set_index("Description")
                st.table(df_comp.style.format("₦{:,.2f}").applymap(color_diff_row, subset=pd.IndexSlice[['Difference (CSV vs Kachasi)'], :]))
            with d3:
                st.write("**Unmatched CSV Categorization**")
                st.table(pd.DataFrame({"Description": ["Customs (NCS)", "Other MDAs", "Total Unmatched CSV"], "Value": [unmatched_csv_customs, unmatched_csv_others, unmatched_csv_customs + unmatched_csv_others]}).set_index("Description").style.format("₦{:,.2f}"))

            # --- EXCEL DOWNLOAD ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # [Previous Excel logic remains unchanged for reliability]
                summary_df = pd.DataFrame([['EXECUTIVE RECONCILIATION DASHBOARD', ''], ['Run Timestamp', run_time], ['Overall GL Deposit', gl_review['Deposit'].sum()]])
                summary_df.to_excel(writer, sheet_name='Summary', index=False, header=False)
                gl_review.to_excel(writer, sheet_name='GL_Review', index=False)
                nibss_review.to_excel(writer, sheet_name='NIBSS_Review', index=False)

            st.download_button("📥 Download Executive Report", data=output.getvalue(), file_name="Executive_Recon_Report.xlsx")
        else:
            st.error("Please upload both GL and NIBSS files.")
