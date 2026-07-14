import streamlit as st
import pandas as pd
import datetime
import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# --- Setup ---
st.set_page_config(page_title="Indigo Ops Portal", layout="wide", page_icon="🅿️")
DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "submissions.csv")
os.makedirs(DATA_DIR, exist_ok=True)
LOGO_URL = "https://i.ibb.co/DHgswzDq/indigo-park-canada-logo.jpg"

# --- PDF Generation Function (Vertical) ---
def create_vertical_pdf(df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()
    
    elements.append(Paragraph("Indigo Park - City Reporting Matrix Report", styles['Title']))
    elements.append(Spacer(1, 12))
    
    for index, row in df.iterrows():
        elements.append(Paragraph(f"Record #{index + 1}", styles['Heading2']))
        
        data = []
        for col, val in row.items():
            # Force word wrap using Paragraph
            key_para = Paragraph(str(col), styles['Normal'])
            val_para = Paragraph(str(val), styles['Normal'])
            data.append([key_para, val_para])
        
        table = Table(data, colWidths=[250, 150])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 20)) 
    
    doc.build(elements)
    return buffer.getvalue()

# --- Data Management ---
def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_csv(DATA_FILE)
    return pd.DataFrame()

def save_data(df):
    df.to_csv(DATA_FILE, index=False)

# --- Config ---
CMO_LIST = ["CMO001", "CMO002", "CMO020", "CMO037", "CMO101", "CMO108", "CMO111", "CMO145"]
YEARS = list(range(2024, 2031))

TASKS = {
    "English": [
        "Tier 1 monthly report completed", "CRITICAL: Scheduled monthly meeting/call completed",
        "Unplanned monthly contact made by the General Manager", "Industry news shared every month",
        "IPC/Indigo Group news shared every month", "Monthly SMILE audit completed",
        "Marketing and special/sports events reporting completed", "Value-add propositions delivered to clients each month",
        "Other reference benchmarks or points of interest discussed"
    ],
    "Français": [
        "Rapport mensuel de niveau 1 terminé", "CRITICAL: Appel/réunion mensuel programmé terminé",
        "Contact mensuel non planifié effectué par le directeur général", "Actualités de l'industrie partagées chaque mois",
        "Actualités IPC/Indigo Group partagées chaque mois", "Audit mensuel SMILE terminé",
        "Marketing et rapports d'événements spéciaux/sportifs terminés", "Des propositions à valeur ajoutée livrées aux clients chaque mois",
        "Autres points de référence ou d'intérêt"
    ]
}

MONTHS = {
    "English": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    "Français": ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
}

LANGS = {
    "English": {
        "title": "City Reporting Matrix", "tab1": "Form", "tab2": "History",
        "cmo": "Select CMO ID", "year": "Select Year", "month": "Select Month",
        "sign": "Full Name", "submit": "Submit Report", "del": "Delete Selected Record",
        "comm": "Justification (Required for NO/NA)", "attest": "I certify the information is accurate.",
        "success": "Record Saved!", "dl_excel": "📥 Export Excel (Vertical)", "dl_pdf": "📥 Export PDF (Vertical)"
    },
    "Français": {
        "title": "Matrice de Rapports", "tab1": "Formulaire", "tab2": "Historique",
        "cmo": "Sélectionner le code CMO", "year": "Sélectionner l'année", "month": "Sélectionner le mois",
        "sign": "Nom complet", "submit": "Soumettre", "del": "Supprimer l'enregistrement",
        "comm": "Justification (Requise pour NO/NA)", "attest": "Je certifie que les informations sont exactes.",
        "success": "Données enregistrées !", "dl_excel": "📥 Exporter Excel (Vertical)", "dl_pdf": "📥 Exporter PDF (Vertical)"
    }
}

# --- Sidebar ---
st.sidebar.image(LOGO_URL, width=150)
lang = st.sidebar.selectbox("Language / Langue", ["English", "Français"])
T = LANGS[lang]
task_list = TASKS[lang]
month_list = MONTHS[lang]

# --- Main App ---
st.title(T["title"])
tab1, tab2 = st.tabs([T["tab1"], T["tab2"]])

with tab1:
    col_a, col_b, col_c = st.columns(3)
    cmo = col_a.selectbox(T["cmo"], CMO_LIST)
    year = col_b.selectbox(T["year"], YEARS)
    month = col_c.selectbox(T["month"], month_list)
    
    st.divider()
    
    responses = {}
    comments = {}
    yes_count, total_valid = 0, 0
    
    # IMPORTANT: index=None ensures no radio button is pre-selected
    for i, task in enumerate(task_list):
        st.markdown(f"**{i+1}. {task}**")
        val = st.radio(f"r_{i}", ["YES", "NO", "N/A"], horizontal=True, index=None, label_visibility="collapsed")
        responses[task] = val
        
        if val in ["NO", "N/A"]:
            comments[task] = st.text_input(T["comm"], key=f"c_{i}")
            if val == "NO": total_valid += 1
        elif val == "YES":
            yes_count += 1
            total_valid += 1
            
    pct = (yes_count / total_valid * 100) if total_valid > 0 else 0
    st.metric("Completion %", f"{pct:.1f}%")
    
    signature = st.text_input(T["sign"])
    attest = st.checkbox(T["attest"])
    
    if st.button(T["submit"]):
        if not signature or not attest: st.error("Missing info.")
        else:
            new_row = {"Date": datetime.datetime.now().strftime("%Y-%m-%d"), "Year": year, "Month": month, "CMO": cmo, "User": signature, "Score": pct}
            new_row.update(responses)
            new_row.update({f"Comm_{k}": v for k, v in comments.items()})
            df = load_data()
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_data(df)
            st.success(T["success"])

with tab2:
    df = load_data()
    if not df.empty:
        col_f1, col_f2 = st.columns(2)
        s_cmo = col_f1.text_input("Filter by CMO")
        s_year = col_f2.selectbox("Filter by Year", ["All"] + sorted(df['Year'].unique().tolist()))
        
        if s_cmo: df = df[df['CMO'].str.contains(s_cmo, case=False)]
        if s_year != "All": df = df[df['Year'] == s_year]
        
        st.dataframe(df, use_container_width=True)
        
        col1, col2, col3 = st.columns(3)
        
        # Excel Vertical
        buffer_xls = io.BytesIO()
        df.T.to_excel(buffer_xls) 
        col1.download_button(T["dl_excel"], buffer_xls.getvalue(), "report.xlsx")
        
        # PDF Vertical
        pdf_data = create_vertical_pdf(df)
        col2.download_button(T["dl_pdf"], pdf_data, "report.pdf")
        
        target_idx = st.selectbox("Select row to delete", df.index)
        if col3.button(T["del"]):
            df = df.drop(target_idx)
            save_data(df)
            st.rerun()
