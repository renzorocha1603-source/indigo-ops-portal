import streamlit as st
import pandas as pd
import datetime
import os
from PIL import Image

# 1. Page Configuration & Brand Styling
st.set_page_config(page_title="Indigo Park Ops Portal", layout="wide", page_icon="🅿️")

st.markdown("""
    <style>
    .reportview-container .main .block-container{ max-width: 1200px; }
    h1, h2, h3 { color: #4A154B; }
    .stButton>button { background-color: #E00073; color: white; border-radius: 6px; }
    </style>
""", unsafe_allow_html=True)

# Header Section
col_logo, col_title = st.columns([1, 4])
with col_logo:
    try:
        st.image(Image.open("indigo_logo_0.png"), width=180)
    except:
        st.write("🅿️ **INDIGO**")
with col_title:
    st.title("Operations & Compliance Portal")
    st.caption("Montreal Region Lot Management & Best Practices Tracker")

st.markdown("---")

# 2. Dynamic Excel Data Engine
@st.cache_data
def extract_real_cmo_data():
    try:
        # Read the raw master sheet from your uploaded workbook
        df = pd.read_excel('Montreal Lot List.xlsx', sheet_name='Montreal', skiprows=1)
        # Clean and extract valid unique CMO identifiers from the first column
        df.columns = df.columns.str.strip()
        cmo_list = df['Lot #'].dropna().astype(str).str.strip().unique().tolist()
        # Ensure we only pick up rows starting with valid identifiers
        cmo_list = [x for x in cmo_list if x.startswith(('CMO', 'VMO'))]
        return sorted(cmo_list)
    except Exception as e:
        # Fallback if file isn't in directory yet
        return [f"CMO{str(i).zfill(3)}" for i in [2, 4, 8, 9, 10, 15, 20, 25]]

cmo_options = extract_real_cmo_data()

# 3. Sidebar Filters & Historic Search
st.sidebar.header("🔍 Report Search Matrix")
filter_user = st.sidebar.text_input("Search by User/Inspector", value="")
filter_year = st.sidebar.selectbox("Filter Year", [2026, 2025, 2024])
filter_month = st.sidebar.selectbox("Filter Month", ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])

# 4. Main Submission Interface
st.header("📝 Monthly Activity Checklist")
selected_cmo = st.selectbox("Select Target CMO to File/Review:", cmo_options)

st.subheader(f"📋 Requirements Matrix: {selected_cmo} ({filter_month} {filter_year})")

col1, col2 = st.columns(2)
with col1:
    q1 = st.radio("1. Was regular lot maintenance and litter pickup performed?", ["YES", "NO", "N/A"])
    q2 = st.radio("2. Are payment systems (meters, apps, cash units) 100% operational?", ["YES", "NO", "N/A"])
    q3 = st.radio("3. Was janitorial, pressure washing, or sweeping completed this cycle?", ["YES", "NO", "N/A"])

with col2:
    q4 = st.radio("4. CRITICAL CORE TASK: Did you meet with the customer/site manager face-to-face this month?", ["YES", "NO", "N/A"])
    comments = st.text_area("Operational Comments / Irregularity Notes", placeholder="Provide explicit text details if any task above was marked 'NO' or 'N/A'...")

# 5. Math Scoring Engine (Dynamic Score Capping)
st.markdown("---")
st.subheader("📊 Operational Metrics")

total_tasks = 4
yes_count = [q1, q2, q3, q4].count("YES")
no_count = [q1, q2, q3, q4].count("NO")
na_count = [q1, q2, q3, q4].count("N/A")

applicable_tasks = total_tasks - na_count
calculated_percentage = (yes_count / applicable_tasks * 100) if applicable_tasks > 0 else 100.0

# Apply Capping Parameter (The Gatekeeper Logic)
is_capped = False
if q4 == "NO" and calculated_percentage > 80.0:
    calculated_percentage = 80.0
    is_capped = True

m1, m2, m3 = st.columns(3)
m1.metric("Form Status", "Complete" if (yes_count + no_count + na_count == 4) else "Pending")
if is_capped:
    m2.metric("Performance Score", f"{calculated_percentage:.1f}%", delta="- Capped (Missed Client Face-Time)", delta_color="inverse")
else:
    m2.metric("Performance Score", f"{calculated_percentage:.1f}%")
m3.metric("Passed Parameters", f"{yes_count} / {applicable_tasks}")

# 6. E-Signatures & Secure Storage Integration
st.markdown("---")
st.subheader("Automated Data Sync")

typed_signature = st.text_input("Type Full Legal Name for E-Signature Verification:")
attestation = st.checkbox("I verify that all field operational metrics submitted above represent true site audits.")

if st.button("💾 Submit and Hard-Sync to Database"):
    if typed_signature and attestation:
        # Create storage payload
        payload = {
            "Timestamp": [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            "Year": [filter_year],
            "Month": [filter_month],
            "CMO": [selected_cmo],
            "Score": [f"{calculated_percentage:.1f}%"],
            "User": [typed_signature],
            "Q1_Maintenance": [q1],
            "Q2_Payments": [q2],
            "Q3_Janitorial": [q3],
            "Q4_ClientMeet": [q4],
            "Comments": [comments]
        }
        new_df = pd.DataFrame(payload)
        
        # Save or append cleanly to database CSV
        db_path = "data/submissions.csv"
        os.makedirs("data", exist_ok=True)
        if not os.path.isfile(db_path):
            new_df.to_csv(db_path, index=False)
        else:
            new_df.to_csv(db_path, mode='a', header=False, index=False)
            
        st.success(f"Successfully compiled! Report data securely appended to history database. Logged by {typed_signature}.")
    else:
        st.error("Submission blocked. You must provide a typed signature and agree to the attestation terms.")
