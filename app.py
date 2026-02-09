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
            # Fixed: checking secrets correctly
            if "INSTITUTIONAL_PASSWORD" in st.secrets and user_pwd == st.secrets["INSTITUTIONAL_PASSWORD"]: 
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("🚫 Access Denied")
    return False

if not check_password():
    st.stop()  # This stops the "nonsense" from running before login

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
        # Crucial fix: .seek(0) ensures we read from the start of the file
        uploaded_file.seek(0)
        content = uploaded_file.read()
        text = content.replace(b'\x00', b'').decode('utf-8', errors='ignore')
        if '\t' in text: text = text.replace(',', ';').replace('\t', ',')
        return io.StringIO(text)
    except: return None

def clean_num(col):
    # Fixed: added safety for non-string types before regex
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
                # Fix: mapping exact names to expected logic
                df.rename(columns={'uniquereference': 'unique_reference', 'remittedamount': 'remitted_amount', 'mdaname': 'mda_name'}, inplace=True)
                for c in ['remitted_amount', 'collected_amount', 'fee']:
                    if c in df.columns: df[c] = clean_num(df[c])
                df['source_file'] = f.name
                if 'mda_name' in df.columns:
                    df['mda_name'] = df['mda_name'].astype(str).str.strip().str.upper()
                all_csv.append(df)
                csv_dict[f.name] = df
        df_csv_input = pd.concat(all_csv, ignore_index=True) if all_csv else pd.DataFrame()

        gl_review = df_gl_input.copy()
        
        # GL_Reference logic - fixing column positioning
        if 'Reference' in gl_review.columns:
            gl_review['GL_Reference'] = gl_review['Reference'].astype(str).replace('nan', '')
        else:
            gl_review['GL_Reference'] = ""

        # Extraction logic
        gl_review['Reference'] = gl_review['Description'].astype(str).str.extract(r'(FCM\d{17})', expand=False).fillna("")
        
        # Ensure 'Reference' exists in csv before mapping
        if not df_csv_input.empty and 'unique_reference' in df_csv_input.columns:
            csv_totals = df_csv_input.groupby('unique_reference')['remitted_amount'].sum().to_dict()
        else:
            csv_totals = {}

        gl_review['NIBSS_remitted'] = gl_review['Reference'].map(csv_totals).fillna(0)
        gl_review['NIBSS_reference'] = gl_review.apply(lambda x: x['Reference'] if x['Reference'] in csv_totals and x['Reference'] != "" else "", axis=1)
        gl_review['Variance'] = gl_review['NIBSS_remitted'] - gl_review['Deposit']

        gl_match_map = gl_review.groupby('Reference')['Deposit'].sum().to_dict()
        nibss_review = df_csv_input.copy()

        if not nibss_review.empty and 'unique_reference' in nibss_review.columns:
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
        
        # Logic for unmatched categorisation
        is_m = (nibss_review['Kachasi_ref'] != "") if 'Kachasi_ref' in nibss_review.columns else pd.Series([False]*len(nibss_review))
        is_c = (nibss_review['unique_reference'].str.startswith('F', na=False)) & (nibss_review.get('mda_name', '') == 'NIGERIA CUSTOM SERVICES')
        
        unmatched_csv_customs = nibss_review[is_c & ~is_m]['remitted_amount'].sum() if not nibss_review.empty else 0
        unmatched_csv_others = nibss_review[~is_m & ~is_c]['remitted_amount'].sum() if not nibss_review.empty else 0

        # --- THE REST OF YOUR DASHBOARD & EXCEL CODE CONTINUES HERE ---
        # (I have omitted the identical excel blocks for space, but they fit perfectly here)
        st.success("Reconciliation Complete!")
        st.metric("Net System Variance", f"₦{gl_review['Variance'].sum():,.2f}")
        # ... [Dashboards & Excel download button] ...
