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

if not check_password():
    st.stop()

# --- STYLING CONSTANTS ---
NAVY_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
GREY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
WHITE_TEXT = Font(color="FFFFFF", bold=True)
BLACK_BOLD = Font(color="000000", bold=True)
RED_BOLD = Font(color="CC0000", bold=True)
DOUBLE_BORDER = Border(top=Side(style='double'), bottom=Side(style='double'))
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

st.title("📂 Institutional Reconciliation System")

# --- SIDEBAR ---
st.sidebar.header("🔍 Quick Search")
search_ref = st.sidebar.text_input("Enter FCM Reference to Track")

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

# --- RECONCILIATION ENGINE ---
if st.button("🚀 Run Reconciliation"):
    if gl_uploads and csv_uploads:
        run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        all_gl, all_csv, csv_dict = [], [], {}
        
        # Load GL
        for f in gl_uploads:
            df = pd.read_excel(f)
            df.columns = [str(c).strip() for c in df.columns]
            for c in ['Deposit', 'Withdrawal', 'Balance']:
                if c in df.columns: df[c] = clean_num(df[c])
            all_gl.append(df)
        df_gl_input = pd.concat(all_gl, ignore_index=True)

        # Load CSV
        for f in csv_uploads:
            buffer = heavy_clean_file(f)
            if buffer:
                df = pd.read_csv(buffer, on_bad_lines='skip', engine='python')
                df.columns = [str(c).strip().lower() for c in df.columns]
                df.rename(columns={'uniquereference': 'unique_reference', 'remittedamount': 'remitted_amount', 'mdaname': 'mda_name'}, inplace=True)
                for c in ['remitted_amount', 'collected_amount', 'fee']:
                    if c in df.columns: df[c] = clean_num(df[c])
                if 'mda_name' in df.columns:
                    df['mda_name'] = df['mda_name'].astype(str).str.strip().str.upper()
                all_csv.append(df)
                csv_dict[f.name] = df
        df_csv_input = pd.concat(all_csv, ignore_index=True)

        # Calculations
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

        # Browser Dashboard
        st.markdown("---")
        st.subheader("📊 Executive Summary")
        m1, m2, m3 = st.columns(3)
        m1.metric("Overall GL Deposit", f"₦{gl_review['Deposit'].sum():,.2f}")
        m2.metric("Overall NIBSS Remitted", f"₦{df_csv_input['remitted_amount'].sum():,.2f}")
        m3.metric("Net Variance", f"₦{gl_review['Variance'].sum():,.2f}")

        # Detailed Tables
        st.markdown("### Detailed Reconciliation Breakdown")
        def color_diff(val): return 'color: red; font-weight: bold' if val != 0 else ''
        
        d1, d2, d3 = st.columns(3)
        matched_mask = (gl_review['NIBSS_reference'] != "")
        total_matched_gl = gl_review[matched_mask]['Deposit'].sum()
        total_matched_nibss = gl_review[matched_mask]['NIBSS_remitted'].sum()
        
        with d1:
            st.write("**Excel to CSV Bridging**")
            df_b = pd.DataFrame({"Description": ["Matched GL Deposit", "Matched NIBSS Remitted", "Difference", "Unmatched GL"], 
                                 "Value": [total_matched_gl, total_matched_nibss, total_matched_gl - total_matched_nibss, gl_review[~matched_mask & (gl_review['Deposit']>0)]['Deposit'].sum()]}).set_index("Description")
            st.table(df_b.style.format("₦{:,.2f}").applymap(color_diff, subset=pd.IndexSlice[['Difference'], :]))

        # --- EXCEL GENERATION (FULL RESTORE) ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Summary Sheet
            summary_sections = [['EXECUTIVE RECONCILIATION DASHBOARD', ''], ['Overall GL Deposit', gl_review['Deposit'].sum()], ['Overall NIBSS Remittance', df_csv_input['remitted_amount'].sum()]]
            pd.DataFrame(summary_sections).to_excel(writer, sheet_name='Executive Summary', index=False, header=False)
            
            # Review Sheets
            gl_review.to_excel(writer, sheet_name='GL_Review', index=False)
            nibss_review.to_excel(writer, sheet_name='NIBSS_Review', index=False)

            # MDA Extracts
            mda_configs = [('NIGERIA CUSTOM SERVICES', 'Customs_Extract'), ('FEDERAL MINISTRY OF FINANCE, BUDGET AND NATIONAL PLANNING - HQTRS', 'NESS_Extract')]
            for target, sname in mda_configs:
                ext = df_csv_input[df_csv_input['mda_name'] == target]
                ext.to_excel(writer, sheet_name=sname, index=False)

        st.download_button("📥 Download Executive Report", data=output.getvalue(), file_name="Recon_Report.xlsx")
    else:
        st.error("Please upload both GL and NIBSS files.")
