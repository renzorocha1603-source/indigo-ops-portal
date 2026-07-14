import streamlit as st
import pandas as pd
import datetime
import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.graphics.shapes import Drawing, Wedge, String
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from openpyxl.chart import DoughnutChart, Reference

# --- Full Hardcoded Client List ---
LOT_DATA = [
    {"cmo": "M002", "name": "Youville", "client": "Les Placements St-Paul Inc"},
    {"cmo": "M020", "name": "2984 Taschereau", "client": "Northwest Healthcare Properties Corporation"},
    {"cmo": "M037", "name": "Faubourg Ste Catherine", "client": "Faubourg Ste Catherine"},
    {"cmo": "M101", "name": "Place Ville Marie", "client": "Ivanhoe Cambridge"},
    {"cmo": "M102", "name": "Place Bonaventure", "client": "Kevric"},
    {"cmo": "M111", "name": "1981 McGill College", "client": "Kevric"},
    {"cmo": "M119", "name": "Place Montréal Trust", "client": "Ivanhoe Cambridge"},
    {"cmo": "M132", "name": "Complexe Desjardins", "client": "Groupe Immobilier Desjardins Inc"},
    {"cmo": "M141", "name": "Quartier International", "client": "Cité Internationale"},
    {"cmo": "M145", "name": "Dominion Square", "client": "Kevric Allied"},
    {"cmo": "M179", "name": "Gare Centrale CN – Belmont", "client": "Complexe de la Gare Centrale 2 S.E.C."},
    {"cmo": "M230", "name": "Le Seaforth", "client": "Groupe Accueil International Ltee"},
    {"cmo": "M242", "name": "Place Bell", "client": "Evenko Montreal"},
    {"cmo": "M275", "name": "2200 McGill", "client": "Allied"},
    {"cmo": "M276", "name": "2000 McGill", "client": "Allied"},
    {"cmo": "M281", "name": "Centre EATON", "client": "Ivanhoe Cambridge"},
    {"cmo": "M296", "name": "Espace Montmorency", "client": "Espace Montmorency"},
    {"cmo": "M305", "name": "1350-1360 René-Lévesque Ouest", "client": "Northwest Healthcare Properties"},
    {"cmo": "M504", "name": "VIA Rail Dorval", "client": "VIA Rail"}
]
LOT_OPTIONS = [f"{d['cmo']} - {d['name']} ({d['client']})" for d in LOT_DATA]

# --- PDF Graphic Helper ---
def create_progress_circle(pct):
    d = Drawing(50, 50)
    d.add(Wedge(25, 25, 20, 0, 360, fillColor=colors.whitesmoke, strokeColor=colors.lightgrey))
    angle = (pct / 100) * 360
    d.add(Wedge(25, 25, 20, 90, 90 - angle, fillColor=colors.HexColor('#003366')))
    d.add(String(15, 20, f"{int(pct)}%", fontSize=10, fontName='Helvetica-Bold'))
    return d

# --- Professional PDF Generator ---
def generate_professional_pdf(df, task_list):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], fontSize=16, textColor=colors.HexColor('#003366'), spaceAfter=12)
    
    for _, row in df.iterrows():
        elements.append(Paragraph("INDIGO PARK - CITY REPORTING MATRIX", title_style))
        status = "COMPLETED" if row.get('Completed_Flag', False) else "PENDING"
        elements.append(Table([[
            Paragraph(f"<b>Status:</b> {status}<br/><b>Date:</b> {row['Date']}<br/><b>Location:</b> {row['CMO_Info']}", styles['Normal']),
            create_progress_circle(float(row['Score']))
        ]], colWidths=[400, 60]))
        elements.append(Spacer(1, 20))
        table_data = [["Task", "Status", "Comment"]]
        for task in task_list:
            table_data.append([Paragraph(str(task), styles['Normal']), str(row.get(task, "N/A")), Paragraph(str(row.get(f"Comm_{task}", "")), styles['Normal'])])
        t = Table(table_data, colWidths=[250, 60, 200])
        t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d9d9d9')), ('GRID', (0, 0), (-1, -1), 0.5, colors.black), ('VALIGN', (0, 0), (-1, -1), 'TOP')]))
        elements.append(t)
        elements.append(Spacer(1, 40))
    doc.build(elements)
    return buffer.getvalue()

# --- Excel Vertical + Circle Graphic ---
def generate_vertical_excel(df, task_list):
    output = io.BytesIO()
    row = df.iloc[-1]
    data = []
    for t in task_list:
        data.append({"Question": t, "Answer": row.get(t, "N/A"), "Comment": row.get(f"Comm_{t}", "")})
    df_vertical = pd.DataFrame(data)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_vertical.to_excel(writer, index=False, sheet_name='Report')
        ws = writer.book['Report']
        ws['E2'] = "Completion %"
        ws['E3'] = float(row['Score'])
        chart = DoughnutChart()
        chart.title = "Report Completion"
        chart.add_data(Reference(ws, min_col=5, min_row=3, max_row=3), from_rows=True)
        ws.add_chart(chart, "G2")
    return output.getvalue()

# --- App Config ---
DATA_FILE = "submissions.csv"
def load_data(): return pd.read_csv(DATA_FILE) if os.path.exists(DATA_FILE) else pd.DataFrame()
def save_data(df): df.to_csv(DATA_FILE, index=False)

TASKS = {
    "English": ["Tier 1 monthly report", "CRITICAL: Monthly meeting/call", "Unplanned monthly contact", "Industry news shared", "IPC/Indigo Group news shared", "Monthly SMILE audit", "Marketing/Events reporting", "Value-add propositions", "Other points of interest"],
    "Français": ["Rapport mensuel niveau 1", "CRITICAL: Réunion mensuelle", "Contact mensuel imprévu", "News industrie partagées", "News Indigo partagées", "Audit mensuel SMILE", "Rapports Marketing/Événements", "Propositions valeur ajoutée", "Autres points d'intérêt"]
}
LANGS = {
    "English": {"title": "City Reporting Matrix", "tab1": "Form", "tab2": "History", "cmo": "Lot / Client", "year": "Year", "month": "Month", "sign": "Full Name", "submit": "Submit Report", "comm": "Comment", "mark": "Mark as COMPLETED"},
    "Français": {"title": "Matrice de Rapports", "tab1": "Formulaire", "tab2": "Historique", "cmo": "Lot / Client", "year": "Année", "month": "Mois", "sign": "Nom complet", "submit": "Soumettre", "comm": "Commentaire", "mark": "Marquer comme TERMINÉ"}
}

st.set_page_config(layout="wide", page_title="Indigo Ops")
st.sidebar.image("https://i.ibb.co/DHgswzDq/indigo-park-canada-logo.jpg", width=150)
lang = st.sidebar.selectbox("Language / Langue", ["English", "Français"])
T = LANGS[lang]
task_list = TASKS[lang]

st.title(f"🅿️ {T['title']}")
tab1, tab2 = st.tabs([T["tab1"], T["tab2"]])

with tab1:
    c1, c2, c3 = st.columns(3)
    cmo = c1.selectbox(T["cmo"], LOT_OPTIONS)
    year = c2.selectbox(T["year"], list(range(2024, 2031)))
    month = c3.selectbox(T["month"], ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    res, comments, yes_count, total = {}, {}, 0, 0
    for i, task in enumerate(task_list):
        st.markdown(f"**{i+1}. {task}**")
        val = st.radio(f"r_{i}", ["YES", "NO", "N/A"], horizontal=True, index=None, label_visibility="collapsed")
        comments[task] = st.text_input(T["comm"], key=f"c_{i}")
        res[task] = val
        if val == "YES": yes_count += 1
        if val in ["YES", "NO"]: total += 1
    pct = (yes_count / total * 100) if total > 0 else 0
    st.metric("Completion %", f"{pct:.0f}%")
    is_done = st.checkbox(T["mark"])
    sig = st.text_input(T["sign"])
    if st.button(T["submit"]):
        if not sig: st.error("Please enter your name.")
        else:
            row = {"Date": datetime.datetime.now().strftime("%Y-%m-%d"), "Year": year, "Month": month, "CMO_Info": cmo, "User": sig, "Score": pct, "Completed_Flag": is_done}
            row.update(res); row.update({f"Comm_{k}": v for k, v in comments.items()})
            save_data(pd.concat([load_data(), pd.DataFrame([row])], ignore_index=True))
            st.success("Report Saved!")

with tab2:
    df = load_data()
    if not df.empty:
        st.dataframe(df, use_container_width=True)
        c1, c2 = st.columns([2, 1])
        c1.download_button("📥 Excel Vertical + Chart", generate_vertical_excel(df, task_list), "report.xlsx")
        c1.download_button("📥 PDF Professional + Circle", generate_professional_pdf(df, task_list), "report.pdf")
        idx = c2.selectbox("Select row to DELETE if mistake:", df.index)
        if c2.button("DELETE REPORT"):
            save_data(df.drop(idx))
            st.rerun()
