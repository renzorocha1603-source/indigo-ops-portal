import streamlit as st
import pandas as pd
import datetime
import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- Hardcoded Client List ---
LOT_DATA = [
    {"cmo": "M002", "name": "Youville", "client": "Les Placements St-Paul Inc"},
    {"cmo": "M020", "name": "2984 Taschereau", "client": "Northwest Healthcare Properties"},
    {"cmo": "M037", "name": "Faubourg Ste Catherine", "client": "Faubourg Ste Catherine"},
    {"cmo": "M101", "name": "Place Ville Marie", "client": "Ivanhoe Cambridge"},
    {"cmo": "M102", "name": "Place Bonaventure", "client": "Kevric"},
    {"cmo": "M111", "name": "1981 McGill College", "client": "Kevric"},
    {"cmo": "M119", "name": "Place Montréal Trust", "client": "Ivanhoe Cambridge"},
    {"cmo": "M132", "name": "Complexe Desjardins", "client": "Groupe Immobilier Desjardins"},
    {"cmo": "M141", "name": "Quartier International", "client": "Cité Internationale"},
    {"cmo": "M145", "name": "Dominion Square", "client": "Kevric Allied"},
    {"cmo": "M179", "name": "Gare Centrale CN", "client": "Complexe de la Gare Centrale 2 S.E.C."},
    {"cmo": "M230", "name": "Le Seaforth", "client": "Groupe Accueil International"},
    {"cmo": "M242", "name": "Place Bell", "client": "Evenko Montreal"},
    {"cmo": "M275", "name": "2200 McGill", "client": "Allied"},
    {"cmo": "M276", "name": "2000 McGill", "client": "Allied"},
    {"cmo": "M281", "name": "Centre EATON", "client": "Ivanhoe Cambridge"},
    {"cmo": "M296", "name": "Espace Montmorency", "client": "Espace Montmorency"},
    {"cmo": "M305", "name": "1350-1360 René-Lévesque", "client": "Northwest Healthcare Properties"},
    {"cmo": "M504", "name": "VIA Rail Dorval", "client": "VIA Rail"}
]
LOT_OPTIONS = [f"{d['cmo']} - {d['name']} ({d['client']})" for d in LOT_DATA]

# --- PDF Generation (Clean & Simple) ---
def create_vertical_pdf(df, task_list):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    bold_style = styles['Normal']
    bold_style.fontName = 'Helvetica-Bold'
    
    elements.append(Paragraph("Indigo Park - Monthly Reporting Report", title_style))
    elements.append(Spacer(1, 12))
    
    for index, row in df.iterrows():
        status = "COMPLETED" if row.get('Completed_Flag', False) else "PENDING"
        elements.append(Paragraph(f"<b>Record #{index + 1} | {row['CMO_Info']} | Status: {status}</b>", styles['Heading2']))
        elements.append(Paragraph(f"Date: {row['Date']} | Score: {float(row['Score']):.0f}%", styles['Normal']))
        elements.append(Spacer(1, 10))
        
        for task in task_list:
            val = row.get(task, "N/A")
            comm = row.get(f"Comm_{task}", "")
            elements.append(Paragraph(f"<b>{task}:</b> {val}", styles['Normal']))
            if pd.notna(comm) and str(comm).strip() != "":
                elements.append(Paragraph(f"<i>Comment: {comm}</i>", styles['Normal']))
            elements.append(Spacer(1, 5))
        
        elements.append(Spacer(1, 20))
        elements.append(Paragraph("-" * 80, styles['Normal']))
    
    doc.build(elements)
    return buffer.getvalue()

# --- Helpers ---
def load_data(): return pd.read_csv(DATA_FILE) if os.path.exists(DATA_FILE) else pd.DataFrame()
def save_data(df): df.to_csv(DATA_FILE, index=False)

# --- Configuration ---
DATA_DIR = "data"; DATA_FILE = os.path.join(DATA_DIR, "submissions.csv"); os.makedirs(DATA_DIR, exist_ok=True)
TASKS = {
    "English": ["Tier 1 monthly report completed", "CRITICAL: Scheduled monthly meeting/call completed", "Unplanned monthly contact made", "Industry news shared", "IPC/Indigo Group news shared", "Monthly SMILE audit completed", "Marketing/Events reporting completed", "Value-add propositions delivered", "Other points of interest"],
    "Français": ["Rapport mensuel niveau 1 terminé", "CRITICAL: Réunion mensuelle complétée", "Contact mensuel imprévu fait", "News industrie partagées", "News Indigo partagées", "Audit mensuel SMILE complété", "Rapports Marketing/Événements", "Propositions valeur ajoutée", "Autres points d'intérêt"]
}
LANGS = {
    "English": {"title": "City Reporting Matrix", "tab1": "Form", "tab2": "History", "cmo": "Lot / Client", "year": "Year", "month": "Month", "sign": "Full Name", "submit": "Submit", "comm": "Comment", "attest": "Verified", "mark_complete": "Mark as Completed"},
    "Français": {"title": "Matrice de Rapports", "tab1": "Formulaire", "tab2": "Historique", "cmo": "Lot / Client", "year": "Année", "month": "Mois", "sign": "Nom complet", "submit": "Soumettre", "comm": "Commentaire", "attest": "Vérifié", "mark_complete": "Marquer comme complet"}
}

# --- App ---
lang = st.sidebar.selectbox("Language / Langue", ["English", "Français"])
T = LANGS[lang]
task_list = TASKS[lang]

st.title(T["title"])
tab1, tab2 = st.tabs([T["tab1"], T["tab2"]])

with tab1:
    col1, col2, col3 = st.columns(3)
    cmo = col1.selectbox(T["cmo"], LOT_OPTIONS)
    year = col2.selectbox(T["year"], list(range(2024, 2031)))
    month = col3.selectbox(T["month"], ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    
    responses, comments, yes_count, total = {}, {}, 0, 0
    for i, task in enumerate(task_list):
        st.markdown(f"**{i+1}. {task}**")
        val = st.radio(f"r_{i}", ["YES", "NO", "N/A"], horizontal=True, index=None, label_visibility="collapsed")
        comments[task] = st.text_input(T["comm"], key=f"c_{i}")
        responses[task] = val
        if val == "YES": yes_count += 1
        if val in ["YES", "NO"]: total += 1
            
    pct = (yes_count / total * 100) if total > 0 else 0
    st.metric("Completion %", f"{pct:.0f}%")
    
    is_completed = st.checkbox(T["mark_complete"])
    sig = st.text_input(T["sign"])
    
    if st.button(T["submit"]):
        if not sig: st.error("Name required")
        else:
            row = {"Date": datetime.datetime.now().strftime("%Y-%m-%d"), "Year": year, "Month": month, "CMO_Info": cmo, "User": sig, "Score": pct, "Completed_Flag": is_completed}
            row.update(responses)
            row.update({f"Comm_{k}": v for k, v in comments.items()})
            df = load_data()
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            save_data(df)
            st.success("Success!")

with tab2:
    df = load_data()
    if not df.empty:
        st.dataframe(df, use_container_width=True)
        # Download
        c1, c2 = st.columns(2)
        buff = io.BytesIO()
        df.T.to_excel(buff)
        c1.download_button("📥 Excel", buff.getvalue(), "report.xlsx")
        c2.download_button("📥 PDF", create_vertical_pdf(df, task_list), "report.pdf")
