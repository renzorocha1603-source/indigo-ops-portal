import streamlit as st
import pandas as pd
import datetime
import os
import io

# Importation sécurisée
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
except ImportError:
    os.system('pip install reportlab')
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

# ==========================================
# 1. CONFIGURATION
# ==========================================
st.set_page_config(page_title="Indigo Park City Matrix Portal", layout="wide", page_icon="🅿️")

st.markdown("""
    <style>
    .block-container { max-width: 95% !important; padding-top: 1.5rem !important; }
    h1, h2, h3 { color: #2D144B; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stButton>button { background-color: #E00073; color: white; border-radius: 4px; font-weight: 600; padding: 0.5rem 1.5rem; border: none; width: 100%; }
    .stButton>button:hover { background-color: #2D144B; color: white; }
    .reason-box { padding: 10px; background-color: #F9F9F9; border-left: 4px solid #E00073; margin-top: 5px; margin-bottom: 10px; border-radius: 4px; }
    </style>
""", unsafe_allow_html=True)

# Path Handling
FILENAME = 'Montreal Lot List.xlsx'
# Check if file exists in the current directory or a data folder
if os.path.exists(FILENAME):
    master_excel_file = FILENAME
else:
    # Fallback to absolute path check
    master_excel_file = os.path.abspath(FILENAME)

logo_url = "https://i.ibb.co/DHgswzDq/indigo-park-canada-logo.jpg"

if "success_message" not in st.session_state:
    st.session_state.success_message = None

# ==========================================
# 2. DICTIONNAIRE BILINGUE
# ==========================================
LANG_DICT = {
    "English": {
        "title": "City Reporting Matrix Portal",
        "subtitle": "Montreal Region Compliance Tracker (2026)",
        "sidebar_lang": "Language",
        "sidebar_ref": "Reference Documents",
        "ref_btn1": "Download: Activity Management",
        "ref_btn2": "Download: Standards Overview",
        "ref_btn3": "Download: Best Practices Matrix",
        "tab_new": "📝 Fill Matrix Form",
        "tab_history": "📜 Database Audit Log",
        "select_cmo": "Select Target Lot ID (CMO):",
        "select_year": "Year:",
        "select_month": "Month:",
        "form_header": "Operational Indicator Survey",
        "comment_placeholder": "Provide specific justification details...",
        "comment_warning": "⚠️ Complete all 9 items and fill mandatory justifications for NO or N/A.",
        "metrics_header": "📊 KPI Performance Metrics",
        "m_status": "Form Status",
        "m_complete": "Completed",
        "m_incomplete": "Incomplete",
        "m_score": "Performance Score",
        "m_capped": "Max Cap Applied",
        "m_passed": "Validated Tasks (YES)",
        "sign_header": "🖋️ Sign-off & Electronic Attestation",
        "sign_label": "Type Full Name for Digital Signature:",
        "attest_label": "I certify that the information filled above is accurate.",
        "submit_btn": "💾 Commit & Sync Matrix Data",
        "err_missing": "Action Blocked: Complete all selections, add required comments, sign your name, and check the attestation checkbox.",
        "success_log": "🎉 Success! Matrix report securely synchronized by",
        "no_history": "No submitted records found in the database yet.",
        "history_title": "🔎 Query & Filter Historical Submissions",
        "search_user": "Search User:",
        "search_cmo": "Search Lot ID:",
        "search_month": "Month:",
        "search_year": "Year:",
        "all": "All",
        "dl_excel": "📥 Export Vertical Excel Log",
        "dl_pdf": "📥 Export Executive PDF Report",
        "tasks": [
            "Tier 1 monthly report completed",
            "CRITICAL: Scheduled monthly meeting/call completed",
            "Unplanned monthly contact made by the General Manager",
            "Industry news shared every month",
            "IPC/Indigo Group news shared every month",
            "Monthly SMILE audit completed",
            "Marketing and special/sports events reporting completed",
            "Value-add propositions delivered to clients each month",
            "Other reference benchmarks or points of interest discussed"
        ]
    },
    "Français": {
        "title": "Portail City Reporting Matrix",
        "subtitle": "Suivi de conformité Région de Montréal (2026)",
        "sidebar_lang": "Langue",
        "sidebar_ref": "Documents de Référence",
        "ref_btn1": "Télécharger : Gestion des activités",
        "ref_btn2": "Télécharger : Aperçu des normes",
        "ref_btn3": "Télécharger : Matrice meilleures pratiques",
        "tab_new": "📝 Saisie de la Matrice",
        "tab_history": "📜 Historique des Données",
        "select_cmo": "Code Emplacement (CMO) :",
        "select_year": "Année :",
        "select_month": "Mois :",
        "form_header": "Formulaire d'Évaluation des Indicateurs",
        "comment_placeholder": "Détails et contexte obligatoire...",
        "comment_warning": "⚠️ Veuillez répondre aux 9 questions et justifier obligatoirement les choix 'NO' ou 'N/A'.",
        "metrics_header": "📊 Indicateurs Clés de Performance (KPI)",
        "m_status": "Statut",
        "m_complete": "Formulaire Conforme",
        "m_incomplete": "Incomplet",
        "m_score": "Note globale",
        "m_capped": "Pénalité maximum appliquée",
        "m_passed": "Tâches Validées",
        "sign_header": "🖋️ Signature Électronique et Validation",
        "sign_label": "Nom et Prénom pour signature numérique :",
        "attest_label": "Je certifie que les déclarations ci-dessus sont exactes et véridiques.",
        "submit_btn": "💾 Enregistrer et Synchroniser la Matrice",
        "err_missing": "Erreur : Formulaire incomplet, justifications manquantes, signature omise ou case d'attestation non cochée.",
        "success_log": "🎉 Succès ! Données enregistrées avec succès par",
        "no_history": "Aucun historique disponible dans la base de données.",
        "history_title": "🔎 Recherche et Filtrage de l'Historique",
        "search_user": "Utilisateur :",
        "search_cmo": "No de Lot :",
        "search_month": "Mois :",
        "search_year": "Année :",
        "all": "Tous",
        "dl_excel": "📥 Télécharger Log Excel Vertical",
        "dl_pdf": "📥 Télécharger Rapport PDF Officiel",
        "tasks": [
            "Rapport mensuel de niveau 1 terminé",
            "CRITICAL: Appel/réunion mensuel programmé terminé",
            "Contact mensuel non planifié effectué par le directeur général",
            "Actualités de l'industrie partagées chaque mois",
            "Actualités IPC/Indigo Group partagées chaque mois",
            "Audit mensuel SMILE terminé",
            "Marketing et rapports d'événements spéciaux/sportifs terminés",
            "Des propositions à valeur ajoutée livrées aux clients chaque mois",
            "Autres points de référence ou d'intérêt"
        ]
    }
}

# ==========================================
# 3. SIDEBAR
# ==========================================
with st.sidebar:
    st.image(logo_url, width=150)
    selected_lang = st.selectbox("🌐 Language / Langue", ["Français", "English"])
    T = LANG_DICT[selected_lang]
    st.markdown("---")
    st.markdown(f"### {T['sidebar_ref']}")
    
    # Helper avec débogage de chemin
    def excel_dl_button(sheet_name, label):
        if os.path.exists(master_excel_file):
            try:
                df = pd.read_excel(master_excel_file, sheet_name=sheet_name)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                st.download_button(label=label, data=buffer.getvalue(), file_name=f"{sheet_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Error loading {sheet_name}: {e}")
        else:
            st.error(f"File not found. Looking at: {os.path.abspath(master_excel_file)}")

    excel_dl_button('Gestion des activités', T['ref_btn1'])
    excel_dl_button('Aperçu des normes', T['ref_btn2'])
    excel_dl_button('Matrice des meilleures pratique', T['ref_btn3'])

st.title(T["title"])
st.caption(T["subtitle"])

# Data loading functions
def get_customer_name(cmo_id):
    if os.path.exists(master_excel_file):
        try:
            df = pd.read_excel(master_excel_file, sheet_name='City Reporting Matrix 2026', skiprows=9)
            df.columns = df.columns.str.strip()
            result = df.loc[df['CMO'] == cmo_id, 'Customer Name']
            return result.values[0] if not result.empty else "N/A"
        except: return "N/A"
    return "N/A"

def load_cmo_codes():
    if os.path.exists(master_excel_file):
        try:
            df = pd.read_excel(master_excel_file, sheet_name='City Reporting Matrix 2026', skiprows=9)
            df.columns = df.columns.str.strip()
            return sorted(list(set(df['CMO'].dropna().astype(str).tolist())))
        except: pass
    return ["CMO001", "CMO002"]

cmo_options = load_cmo_codes()
years_options = list(range(2024, 2036))
months_options = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"] if selected_lang == "Français" else ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

# ==========================================
# 4. FORMULAIRE
# ==========================================
t_form, t_hist = st.tabs([T["tab_new"], T["tab_history"]])

with t_form:
    if st.session_state.success_message:
        st.success(st.session_state.success_message)
        st.session_state.success_message = None

    c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
    selected_cmo = c1.selectbox(T["select_cmo"], cmo_options)
    c2.text_input("Customer Name:", value=str(get_customer_name(selected_cmo)), disabled=True)
    chosen_year = c3.selectbox(T["select_year"], years_options, index=0)
    chosen_month = c4.selectbox(T["select_month"], months_options, index=6)

    st.markdown(f"### {T['form_header']}")
    responses, task_comments = {}, {}
    incomplete_selections_flag, missing_comments_flag = False, False
    
    for i, task in enumerate(T["tasks"], start=1):
        c_text, c_radio = st.columns([5, 2])
        c_text.markdown(f"**{i}.** {task}")
        responses[f"task_{i}"] = c_radio.radio(f"Radio_{i}", ["YES", "NO", "N/A"], index=None, key=f"r_{i}", label_visibility="collapsed")
            
        if responses[f"task_{i}"] is None: incomplete_selections_flag = True
        elif responses[f"task_{i}"] in ["NO", "N/A"]:
            task_comments[f"comment_{i}"] = c_text.text_input(f"Justification:", key=f"c_{i}", placeholder=T["comment_placeholder"], label_visibility="collapsed")
            if not task_comments[f"comment_{i}"].strip(): missing_comments_flag = True
        else: task_comments[f"comment_{i}"] = ""
        st.divider()

    # Calculations
    answers = [v for v in responses.values() if v is not None]
    yes_count = answers.count("YES")
    na_count = answers.count("N/A")
    score = (yes_count / (9 - na_count) * 100) if (9 - na_count) > 0 else 0
    is_capped = (responses.get("task_2") == "NO" and score > 85.0)
    if is_capped: score = 85.0

    st.metric(label=T["m_score"], value=f"{score:.1f}%")
    
    typed_signature = st.text_input(T["sign_label"])
    attest = st.checkbox(T["attest_label"])
    
    if st.button(T["submit_btn"]):
        if incomplete_selections_flag or missing_comments_flag or not typed_signature or not attest:
            st.error(T["err_missing"])
        else:
            payload = {"Timestamp": [datetime.datetime.now()], "User": [typed_signature], "CMO_ID": [selected_cmo], "Score": [score]}
            os.makedirs("data", exist_ok=True)
            pd.DataFrame(payload).to_csv("data/submissions.csv", mode='a', header=not os.path.exists("data/submissions.csv"), index=False)
            st.session_state.success_message = "Succès!"
            st.rerun()

with t_hist:
    if os.path.isfile("data/submissions.csv"):
        st.dataframe(pd.read_csv("data/submissions.csv"), use_container_width=True)
    else:
        st.info(T["no_history"])
