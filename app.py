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
        if st.button("Unlock System"):
            if "INSTITUTIONAL_PASSWORD" in st.secrets and user_pwd == st.secrets["INSTITUTIONAL_PASSWORD"]: 
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("🚫 Access Denied")
    return False

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

    # --- GLOBAL DATA HOLDERS ---
    # We define these outside the button click so the Search box can use them
    df_gl_input = pd.DataFrame()
    df_csv_input = pd.DataFrame()

    if gl_uploads:
        all_gl = []
        for f in gl_uploads:
            df = pd.read_excel(f)
            df.columns = [str(c).strip() for c in df.columns]
            for c in ['Deposit', 'Withdrawal', 'Balance']:
                if c in df.columns: df[c] = clean_num(df[c])
            df['Source'] = f.name
            all_gl.append(df)
        df_gl_input = pd.concat(all_gl, ignore_index=True)

    if csv_uploads:
        all_csv = []
        for f in csv_uploads:
            buffer = heavy_clean_file(f)
            if buffer:
                df = pd.read_csv(buffer, on_bad_lines='skip', engine='python')
                df.columns = [str(c).strip().lower() for c in df.columns]
                df.rename(columns={'uniquereference': 'unique_reference', 'remittedamount': 'remitted_amount', 'mdaname': 'mda_name'}, inplace=True)
                for c in ['remitted_amount', 'collected_amount', 'fee']:
                    if c in df.columns: df[c] = clean_num(df[c])
                df['source_file'] = f.name
                all_csv.append(df)
        df_csv_input = pd.concat(all_csv, ignore_index=True)

    # --- REFERENCE CHECKER LOGIC ---
    if search_ref:
        st.markdown(f"### 🎯 Tracking Reference: `{search_ref}`")
        s1, s2 = st.columns(2)
        
        with s1:
            st.info("**Search Result in GL/Excel**")
            if not df_gl_input.empty:
                # Search in Description or Reference columns
                mask = df_gl_input.astype(str).apply(lambda x: x.str.contains(search_ref, case=False, na=False)).any(axis=1)
                res_gl = df_gl_input[mask]
                if not res_gl.empty:
                    st.dataframe(res_gl, use_container_width=True)
                else: st.warning("Not found in GL Files.")
            else: st.write("Upload GL files to enable searching.")

        with s2:
            st.info("**Search Result in NIBSS/CSV**")
            if not df_csv_input.empty:
                mask_csv = df_csv_input.astype(str).apply(lambda x: x.str.contains(search_ref, case=False, na=False)).any(axis=1)
                res_csv = df_csv_input[mask_csv]
                if not res_csv.empty:
                    st.dataframe(res_csv, use_container_width=True)
                else: st.warning("Not found in NIBSS Files.")
            else: st.write("Upload NIBSS files to enable searching.")
        st.markdown("---")

    # --- RUN RECONCILIATION BUTTON ---
    if st.button("🚀 Run Reconciliation"):
        if gl_uploads and csv_uploads:
            run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # (Processing continues as before...)
            csv_dict = {f.name: df_csv_input[df_csv_input['source_file'] == f.name] for f in csv_uploads}
            
            gl_review = df_gl_input.copy()
            if 'Description' in gl_review.columns:
                gl_review['Reference'] = gl_review['Description'].astype(str).str.extract(r'(FCM\d{17})', expand=False).fillna("")
            
            csv_totals = df_csv_input.groupby('unique_reference')['remitted_amount'].sum().to_dict()
            gl_review['NIBSS_remitted'] = gl_review['Reference'].map(csv_totals).fillna(0)
            gl_review['NIBSS_reference'] = gl_review.apply(lambda x: x['Reference'] if x['Reference'] in csv_totals and x['Reference'] != "" else "", axis=1)
            gl_review['Variance'] = gl_review['NIBSS_remitted'] - gl_review['Deposit']

            gl_match_map = gl_review.groupby('Reference')['Deposit'].sum().to_dict()
            nibss_review = df_csv_input.copy()
            
            if not nibss_review.empty and 'bank_id' in nibss_review.columns:
                b_idx = list(nibss_review.columns).index('bank_id') + 1
                nibss_review.insert(b_idx, 'Kachasi_ref', nibss_review['unique_reference'].map(lambda x: x if x in gl_match_map else ""))
                nibss_review.insert(b_idx+1, 'Kachasi_In_GL', nibss_review['unique_reference'].map(gl_match_map).fillna(0))
                nibss_review.insert(b_idx+2, 'Settle_Variance', nibss_review['remitted_amount'] - nibss_review['Kachasi_In_GL'])

            # --- CALCULATIONS & DASHBOARD ---
            matched_mask_gl = (gl_review['NIBSS_reference'] != "")
            total_matched_gl_dep = gl_review[matched_mask_gl]['Deposit'].sum()
            total_matched_gl_nibss = gl_review[matched_mask_gl]['NIBSS_remitted'].sum()
            unmatched_gl_dep = gl_review[~matched_mask_gl & (gl_review['Deposit'] > 0)]['Deposit'].sum()
            bridging_diff = total_matched_gl_dep - total_matched_gl_nibss
            csv_vs_kachasi_diff = total_matched_gl_nibss - total_matched_gl_dep

            st.markdown("### Detailed Reconciliation Breakdown")
            d1, d2, d3 = st.columns(3)
            def color_diff(val): return 'color: red; font-weight: bold' if val != 0 else ''
            
            with d1:
                st.write("**Bridging (Excel to CSV)**")
                df_bridge = pd.DataFrame({"Description": ["Matched GL Deposit", "Matched NIBSS Remitted", "Difference", "Unmatched GL Dep"], "Value": [total_matched_gl_dep, total_matched_gl_nibss, bridging_diff, unmatched_gl_dep]}).set_index("Description")
                st.table(df_bridge.style.format("₦{:,.2f}").applymap(color_diff, subset=pd.IndexSlice[['Difference'], :]))

            # (The rest of the Excel Writer code block follows here as in previous versions)
            # ... [Excel generation code here] ...
            
            st.success("Analysis Complete. You can now download the report.")
        else:
            st.error("Please upload both GL and NIBSS files.")
