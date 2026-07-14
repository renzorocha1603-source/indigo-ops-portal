import streamlit as st
import pandas as pd
import datetime
import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# --- Setup ---
st.set_page_config(page_title="Indigo Matrix", layout="wide")
DATA_FILE = "data/submissions.csv"
os.makedirs("data", exist_ok=True)
LOGO_URL = "https://i.ibb.co/DHgswzDq/indigo-park-canada-logo.jpg"

# --- Data Handling ---
def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_csv(DATA_FILE)
    return pd.DataFrame()

def save_data(df):
    df.to_csv(DATA_FILE, index=False)

# --- Multilingual Dictionary ---
LANGS = {
    "English": {
        "title": "City Reporting Matrix",
        "t1": "Form", "t2": "History",
        "cmo": "CMO ID", "sign": "Signature",
        "submit": "Submit Report", "del": "Delete Selected",
        "tasks": [f"Task {i}" for i in range(1, 10)], # Replace with your real tasks
        "comm_req": "Comment required for NO/NA"
    },
    "Français": {
        "title": "Matrice de Rapports",
        "t1": "Formulaire", "t2": "Historique",
        "cmo": "ID CMO", "sign": "Signature",
        "submit": "Soumettre", "del": "Supprimer la sélection",
        "tasks": [f"Tâche {i}" for i in range(1, 10)], 
        "comm_req": "Commentaire requis pour NO/NA"
    }
}

# --- Sidebar ---
st.sidebar.image(LOGO_URL, width=150)
lang = st.sidebar.selectbox("Language / Langue", ["English", "Français"])
T = LANGS[lang]

# --- Main App ---
st.title(T["title"])
tab1, tab2 = st.tabs([T["t1"], T["t2"]])

with tab1:
    cmo = st.text_input(T["cmo"])
    responses = {}
    comments = {}
    total_score = 0
    valid_tasks = 0

    for task in T["tasks"]:
        col1, col2 = st.columns([2, 1])
        col1.write(task)
        val = col2.radio(f"r_{task}", ["YES", "NO", "N/A"], horizontal=True, index=None, label_visibility="collapsed")
        responses[task] = val
        
        if val in ["NO", "N/A"]:
            comments[task] = st.text_input(f"{T['comm_req']} ({task})")
        else:
            comments[task] = ""
            if val == "YES":
                total_score += 1
                valid_tasks += 1
            elif val == "N/A":
                pass # N/A doesn't count towards score
            else: # NO
                valid_tasks += 1

    pct = (total_score / valid_tasks * 100) if valid_tasks > 0 else 0
    st.metric("Completion %", f"{pct:.1f}%")
    
    signature = st.text_input(T["sign"])
    if st.button(T["submit"]):
        if not signature or not cmo:
            st.error("Missing Info")
        else:
            new_row = {"Date": datetime.datetime.now().strftime("%Y-%m-%d"), "CMO": cmo, "User": signature, "Score": pct}
            new_row.update(responses)
            new_row.update({f"Comm_{k}": v for k,v in comments.items()})
            df = load_data()
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_data(df)
            st.success("Saved!")

with tab2:
    df = load_data()
    if not df.empty:
        # Search Filters
        c_search, c_date = st.columns(2)
        s_cmo = c_search.text_input("Filter by CMO")
        if s_cmo: df = df[df['CMO'].str.contains(s_cmo, case=False)]
        
        st.dataframe(df)
        
        # Download Buttons
        col_dl1, col_dl2 = st.columns(2)
        
        # Excel
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False)
        col_dl1.download_button("Download Excel", buffer.getvalue(), "data.xlsx")
        
        # PDF Simple
        if col_dl2.button("Generate PDF"):
            pdf_buf = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buf, pagesize=letter)
            story = [Paragraph("Report", getSampleStyleSheet()['Title'])]
            for _, row in df.iterrows():
                story.append(Paragraph(f"CMO: {row['CMO']} | User: {row['User']}", getSampleStyleSheet()['Normal']))
            doc.build(story)
            col_dl2.download_button("Download PDF", pdf_buf.getvalue(), "report.pdf")

        # Delete Logic
        to_del = st.selectbox("Select ID to Delete", df.index)
        if st.button(T["del"]):
            df = df.drop(to_del)
            save_data(df)
            st.rerun()
