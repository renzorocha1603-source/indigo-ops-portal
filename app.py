import streamlit as st
import pandas as pd
import datetime
import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
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

# --- Professional PDF Generator ---
def generate_professional_pdf(df, task_list):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    
    # Custom Styles
    title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], fontSize=18, textColor=colors.HexColor('#003366'), spaceAfter=12)
    heading_style = ParagraphStyle('HeadingStyle', parent=styles['Heading2'], fontSize=12, spaceAfter=6, textColor=colors.darkblue)

    for _, row in df.iterrows():
        elements.append(Paragraph("INDIGO PARK - CITY REPORTING MATRIX", title_style))
        status = "COMPLETED" if row.get('Completed_Flag', False) else "PENDING"
        elements.append(Paragraph(f"<b>Status:</b> {status} | <b>Date:</b> {row['Date']} | <b>Score:</b> {float(row['Score']):.0f}%", styles['Normal']))
        elements.append(Paragraph(f"<b>Location:</b> {row['CMO_Info']}", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Prepare Table Data
        table_data = [["Task", "Status", "Comment"]]
        for task in task_list:
            val = row.get(task, "N/A")
            comm = row.get(f"Comm_{task}", "")
            table_data.append([Paragraph(str(task), styles['Normal']), str(val), Paragraph(str(comm) if pd.notna(comm) else "", styles['Normal'])])
        
        # Create Table
        table = Table(table_data, colWidths=[250, 60, 200])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d9d9d9')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.whitesmoke]),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 30))
        elements.append(Paragraph("__________________________________________________________________________", styles['Normal']))
        elements.append(Spacer(1, 20))
    
    doc.build(elements)
    return buffer.getvalue()

# --- Config ---
DATA_DIR = "data"; DATA_FILE = os.path.join(DATA_DIR, "submissions.csv"); os.makedirs(DATA_DIR, exist_ok=True)
def load_data(): return pd.read_csv(DATA_FILE) if os.path.exists(DATA_FILE) else pd.DataFrame()
def save_data(df): df.to_csv(DATA_FILE, index=False)

TASKS = {
    "English": ["Tier 1 monthly report completed", "CRITICAL: Scheduled monthly meeting/call completed", "Unplanned monthly contact made", "Industry news shared", "IPC/Indigo Group news shared", "Monthly SMILE audit completed", "Marketing/Events reporting completed", "Value-add propositions delivered", "Other points of interest"],
    "Français": ["Rapport mensuel niveau 1 terminé", "CRITICAL: Réunion mensuelle complétée", "Contact mensuel imprévu fait", "News industrie partagées", "News Indigo partagées", "Audit mensuel SMILE complété", "Rapports Marketing/Événements", "Propositions valeur ajoutée", "Autres points d'intérêt"]
}

# --- Main App ---
st.set_page_config(layout="wide")
st.title("🅿️ Indigo Reporting Portal")
tab1, tab2 = st.tabs(["📋 Submission Form", "📂 History & Management"])

with tab1:
    col1, col2, col3 = st.columns(3)
    cmo = col1.selectbox("Select Lot / Client", LOT_OPTIONS)
    year = col2.selectbox("Year", list(range(2024, 2031)))
    month = col3.selectbox("Month", ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    
    responses, comments, yes_count, total = {}, {}, 0, 0
    for i, task in enumerate(TASKS["English"]):
        st.markdown(f"**{i+1}. {task}**")
        val = st.radio(f"r_{i}", ["YES", "NO", "N/A"], horizontal=True, index=None, label_visibility="collapsed")
        comments[task] = st.text_input("Comment", key=f"c_{i}")
        responses[task] = val
        if val == "YES": yes_count += 1
        if val in ["YES", "NO"]: total += 1
            
    pct = (yes_count / total * 100) if total > 0 else 0
    st.metric("Completion %", f"{pct:.0f}%")
    
    is_completed = st.checkbox("Mark report as COMPLETED")
    sig = st.text_input("Full Name")
    
    if st.button("Submit Report"):
        if not sig: st.error("Please enter your name.")
        else:
            row = {"Date": datetime.datetime.now().strftime("%Y-%m-%d"), "Year": year, "Month": month, "CMO_Info": cmo, "User": sig, "Score": pct, "Completed_Flag": is_completed}
            row.update(responses)
            row.update({f"Comm_{k}": v for k, v in comments.items()})
            df = load_data()
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            save_data(df)
            st.success("Report Saved Successfully!")

with tab2:
    df = load_data()
    if not df.empty:
        st.subheader("Manage Records")
        st.dataframe(df, use_container_width=True)
        
        col_export, col_delete = st.columns([2, 1])
        
        with col_export:
            buff = io.BytesIO()
            df.T.to_excel(buff)
            st.download_button("📥 Export All to Excel", buff.getvalue(), "report.xlsx")
            st.download_button("📥 Export Professional PDF", generate_professional_pdf(df, TASKS["English"]), "report.pdf")
            
        with col_delete:
            st.warning("⚠️ Correction Area")
            target_idx = st.selectbox("Select row to delete if there is a mistake", df.index)
            if st.button("Delete Selected Report"):
                df = df.drop(target_idx)
                save_data(df)
                st.rerun()
    else:
        st.info("No records found.")
