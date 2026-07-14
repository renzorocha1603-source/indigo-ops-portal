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
LOT_LIST_FILE = "Montreal Lot List.xlsx - Montreal.csv"
os.makedirs(DATA_DIR, exist_ok=True)
LOGO_URL = "https://i.ibb.co/DHgswzDq/indigo-park-canada-logo.jpg"

# --- Data Loading (Lots) ---
def get_lot_options():
    if os.path.exists(LOT_LIST_FILE):
        df = pd.read_csv(LOT_LIST_FILE)
        # Create a display string: CMO - Client (Manager)
        options = []
        for _, row in df.iterrows():
            cmo = str(row.get('Lot #', 'N/A'))
            client = str(row.get('Client / Owner', 'Unknown'))
            mgr = str(row.get('INDIGO Day-to-Day Manager', 'Unknown'))
            options.append(f"{cmo} | {client} | Mgr: {mgr}")
        return options
    return ["CMO001 | Error Loading File", "CMO002 | Please check file name"]

# --- PDF Generation (Vertical) ---
def create_vertical_pdf(df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()
    elements.append(Paragraph("Indigo Park - City Reporting Matrix", styles['Title']))
    elements.append(Spacer(1, 12))
    
    for index, row in df.iterrows():
        elements.append(Paragraph(f"Record #{index + 1}", styles['Heading2']))
        data = [[Paragraph(str(col), styles['Normal']), Paragraph(str(val), styles['Normal'])] for col, val in row.items()]
        table = Table(data, colWidths=[200, 300])
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BACKGROUND', (0, 0), (0, -1), colors.whitesmoke),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 20))
    doc.build(elements)
    return buffer.getvalue()

# --- Helpers ---
def load_data():
    return pd.read_csv(DATA_FILE) if os.path.exists(DATA_FILE) else pd.DataFrame()

def save_data(df):
    df.to_csv(DATA_FILE, index=False)

# --- Config ---
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

LANGS = {
    "English": {
        "title": "City Reporting Matrix", "tab1": "Form", "tab2": "History",
        "cmo": "Select Lot / Client", "year": "Year", "month": "Month",
        "sign": "Full Name", "submit": "Submit", "del": "Delete",
        "comm": "Comment / Note", "attest": "I certify this is accurate.",
        "dl_excel": "📥 Excel", "dl_pdf": "📥 PDF"
    },
    "Français": {
        "title": "Matrice de Rapports", "tab1": "Formulaire", "tab2": "Historique",
        "cmo": "Sélectionner Lot / Client", "year": "Année", "month": "Mois",
        "sign": "Nom complet", "submit": "Soumettre", "del": "Supprimer",
        "comm": "Commentaire / Note", "attest": "Je certifie que c'est exact.",
        "dl_excel": "📥 Excel", "dl_pdf": "📥 PDF"
    }
}

# --- App ---
st.sidebar.image(LOGO_URL, width=150)
lang = st.sidebar.selectbox("Language / Langue", ["English", "Français"])
T = LANGS[lang]
task_list = TASKS[lang]

st.title(T["title"])
tab1, tab2 = st.tabs([T["tab1"], T["tab2"]])

with tab1:
    col_a, col_b, col_c = st.columns([2,1,1])
    cmo_selection = col_a.selectbox(T["cmo"], get_lot_options())
    year = col_b.selectbox(T["year"], list(range(2024, 2031)))
    month = col_c.selectbox(T["month"], ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    
    responses = {}
    comments = {}
    yes_count, total_valid = 0, 0
    
    for i, task in enumerate(task_list):
        st.markdown(f"**{i+1}. {task}**")
        val = st.radio(f"r_{i}", ["YES", "NO", "N/A"], horizontal=True, index=None, label_visibility="collapsed")
        comments[task] = st.text_input(f"{T['comm']} ({i+1})", key=f"c_{i}")
        
        responses[task] = val
        if val == "YES":
            yes_count += 1
            total_valid += 1
        elif val == "NO":
            total_valid += 1
            
    pct = (yes_count / total_valid * 100) if total_valid > 0 else 0
    st.metric("Completion %", f"{pct:.1f}%")
    
    signature = st.text_input(T["sign"])
    attest = st.checkbox(T["attest"])
    
    if st.button(T["submit"]):
        if not signature or not attest or not cmo_selection: st.error("Fill mandatory fields.")
        else:
            new_row = {"Date": datetime.datetime.now().strftime("%Y-%m-%d"), "Year": year, "Month": month, "CMO_Info": cmo_selection, "User": signature, "Score": pct}
            new_row.update(responses)
            new_row.update({f"Comm_{k}": v for k, v in comments.items()})
            df = load_data()
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_data(df)
            st.success("Success!")

with tab2:
    df = load_data()
    if not df.empty:
        st.dataframe(df, use_container_width=True)
        col1, col2, col3 = st.columns(3)
        
        buffer_xls = io.BytesIO()
        df.T.to_excel(buffer_xls)
        col1.download_button(T["dl_excel"], buffer_xls.getvalue(), "report.xlsx")
        
        pdf_data = create_vertical_pdf(df)
        col2.download_button(T["dl_pdf"], pdf_data, "report.pdf")
        
        target_idx = st.selectbox("Select row to delete", df.index)
        if col3.button(T["del"]):
            df = df.drop(target_idx)
            save_data(df)
            st.rerun()
