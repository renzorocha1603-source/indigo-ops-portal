import streamlit as st
import pandas as pd
import datetime
import os
from PIL import Image

# ==========================================
# 1. PAGE CONFIGURATION & BRAND STYLING
# ==========================================
st.set_page_config(page_title="Indigo Park Ops Portal", layout="wide", page_icon="🅿️")

# Custom CSS for Indigo Park Corporate Identity
st.markdown("""
    <style>
    .reportview-container .main .block-container{ max-width: 1200px; }
    h1, h2, h3 { color: #2D144B; } /* Indigo Dark Purple */
    .stButton>button { background-color: #E00073; color: white; border-radius: 6px; font-weight: bold; } /* Magenta accent */
    div.stRadio > div{ flex-direction:row; gap: 20px; } /* Horizontal Radio buttons */
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. BILINGUAL TRANSLATION DICTIONARY
# ==========================================
# Multi-language dictionary setup
LANG_DICT = {
    "English": {
        "title": "Operations & Compliance Portal",
        "subtitle": "Montreal Region Lot Management & Best Practices Digital Tracker",
        "sidebar_lang": "🌐 Select Language / Choisir la langue",
        "sidebar_search": "🔍 Search & Filter Settings",
        "user_filter": "Filter Database by User Name",
        "year_filter": "Select Year",
        "month_filter": "Select Month",
        "tab_new": "📝 File New Report",
        "tab_history": "📜 View & Search History",
        "select_cmo": "Select Target Lot ID (CMO):",
        "form_header": "Monthly Validation Form",
        "form_instruction": "Fill out the status for all 14 mandatory compliance items:",
        "comments_label": "Operational Comments & Justifications (Required for 'NO' or 'N/A'):",
        "metrics_header": "📊 Performance Metrics Results",
        "m_status": "Form Status",
        "m_complete": "Form Completed",
        "m_incomplete": "Incomplete",
        "m_score": "Operational Performance Score",
        "m_capped": "Penalized Max Capping (No Client Meeting)",
        "m_passed": "Validated Tasks (YES)",
        "sign_header": "🖋️ Digital Signature & Verification",
        "sign_label": "Type your Full Name to sign electronically:",
        "attest_label": "I officially attest that these metrics represent true site audits matching real field conditions.",
        "submit_btn": "💾 Save & Sync Report to Database",
        "err_missing": "Submission Blocked: You must enter your name and check the attestation box.",
        "success_log": "🎉 Report securely saved to history and signed by",
        "no_history": "No records found in the database yet.",
        "history_title": "🔎 Historical Audit Database Matrix",
        "tasks": [
            "Complete operational site report",
            "Site visit by site manager/supervisor - internal",
            "Site visit by site manager/supervisor - external",
            "Account manager call to the client",
            "Account manager site visit / internal review",
            "Account manager site visit / external review",
            "Account manager - Contact with site team",
            "CRITICAL: Account manager - Meeting with the client (Face-to-Face)",
            "General Manager call to the client",
            "General Manager contacts the client in person",
            "Site visit by General Manager / internal and external review",
            "Vice President of Operations contacts the client in person",
            "Debriefs and marketing for special/sporting events",
            "Propose discussions/propositions focused on value-add, revenue generation, or pricing recommendations"
        ]
    },
    "Français": {
        "title": "Portail d'Opérations & Conformité",
        "subtitle": "Suivi numérique des lotissements et meilleures pratiques - Région de Montréal",
        "sidebar_lang": "🌐 Choisir la langue / Select Language",
        "sidebar_search": "🔍 Paramètres de recherche et filtres",
        "user_filter": "Filtrer la base par nom d'utilisateur",
        "year_filter": "Sélectionner l'année",
        "month_filter": "Sélectionner le mois",
        "tab_new": "📝 Remplir un nouveau rapport",
        "tab_history": "📜 Consulter et chercher l'historique",
        "select_cmo": "Sélectionner l'emplacement cible (CMO) :",
        "form_header": "Formulaire de validation mensuelle",
        "form_instruction": "Indiquez le statut pour l'ensemble des 14 indicateurs obligatoires :",
        "comments_label": "Remarques et justifications (Obligatoire si des éléments sont cochés 'NO' ou 'N/A') :",
        "metrics_header": "📊 Résultats Métriques de Performance",
        "m_status": "Statut de saisie",
        "m_complete": "Formulaire Complété",
        "m_incomplete": "Saisie Incomplète",
        "m_score": "Indice de Performance Réel",
        "m_capped": "Pénalité maximum appliquée (Réunion Client manquante)",
        "m_passed": "Tâches Validées (YES)",
        "sign_header": "🖋️ Signature et Validation Légale",
        "sign_label": "Saisissez votre Nom et Prénom complet pour la signature électronique :",
        "attest_label": "J'atteste sur l'honneur avoir complété cette évaluation conformément aux réalités observées sur le terrain.",
        "submit_btn": "💾 Valider et Enregistrer dans la Base Historique",
        "err_missing": "Action Bloquée : Vous devez saisir votre signature nominative et cocher la case d'attestation.",
        "success_log": "🎉 Rapport synchronisé avec succès ! Enregistré de manière sécurisée par",
        "no_history": "Aucun enregistrement trouvé dans la base de données pour le moment.",
        "history_title": "🔎 Matrice Historique des Audits Enregistrés",
        "tasks": [
            "Rapport opérationnel complet du site",
            "Visite du site par le responsable/superviseur du site - interne",
            "Visite du site par le responsable/superviseur du site - externe",
            "Appel du gestionnaire de compte au client",
            "Visite du gestionnaire de compte sur le site/examen interne",
            "Visite du gestionnaire de compte sur le site/examen externe",
            "Gestionnaire de compte - Contact avec l'équipe du site",
            "CRITICAL: Gestionnaire de compte - Réunion avec le client (Face-to-Face Meeting)",
            "Appel du directeur général au client",
            "Le directeur général contacte le client en personne",
            "Visite du site par le directeur général / examen interne et externe",
            "Le vice-président des opérations contacte le client en personne",
            "Comptes rendus et marketing des événements spéciaux/sportifs",
            "Proposer des discussions/propositions axées sur la valeur ajoutée (Génération de revenus, Recommandations tarifaires)"
        ]
    }
}

# ==========================================
# 3. SIDEBAR CONTROLS (GLOBAL FILTERS)
# ==========================================
# Language selection setting
selected_lang = st.sidebar.selectbox(LANG_DICT["English"]["sidebar_lang"], ["Français", "English"])
T = LANG_DICT[selected_lang]

st.sidebar.markdown("---")
st.sidebar.header(T["sidebar_search"])
filter_user = st.sidebar.text_input(T["user_filter"], value="", placeholder="e.g. Francis")
filter_year = st.sidebar.selectbox(T["year_filter"], [2026, 2025, 2024])

months_list = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"] if selected_lang == "Français" else ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
filter_month = st.sidebar.selectbox(T["month_filter"], months_list)

# ==========================================
# 4. HEADER LAYOUT WITH ORIGINAL LOGO
# ==========================================
col_logo, col_title = st.columns([1, 4])
with col_logo:
    try:
        st.image(Image.open("indigo_logo_0.png"), width=180)
    except:
        st.write("🅿️ **INDIGO PARK**")
with col_title:
    st.title(T["title"])
    st.caption(T["subtitle"])

st.markdown("---")

# ==========================================
# 5. EXCEL DATA PARSER (LOAD CMO LIST)
# ==========================================
@st.cache_data
def extract_real_cmo_data():
    try:
        df = pd.read_excel('Montreal Lot List.xlsx', sheet_name='Montreal', skiprows=1)
        df.columns = df.columns.str.strip()
        cmo_list = df['Lot #'].dropna().astype(str).str.strip().unique().tolist()
        cmo_list = [x for x in cmo_list if x.startswith(('CMO', 'VMO'))]
        return sorted(cmo_list)
    except:
        return [f"CMO{str(i).zfill(3)}" for i in [2, 4, 8, 9, 10, 15, 20, 25, 37, 101, 102]]

cmo_options = extract_real_cmo_data()

# Create Multi-tab Dashboard view
tab1, tab2 = st.tabs([T["tab_new"], T["tab_history"]])

# ==========================================
# TAB 1: FILE NEW MONTHLY COMPLIANCE REPORT
# ==========================================
with tab1:
    selected_cmo = st.selectbox(T["select_cmo"], cmo_options)
    st.subheader(f"{T['form_header']} : {selected_cmo} — ({filter_month} {filter_year})")
    st.write(T["form_instruction"])
    
    # Render the 14 rows dynamically
    responses = {}
    for idx, task in enumerate(T["tasks"], start=1):
        st.markdown(f"**{idx}. {task}**")
        responses[f"task_{idx}"] = st.radio(
            f"Statut {idx}", ["YES", "NO", "N/A"], key=f"task_radio_{idx}", label_visibility="collapsed"
        )
        st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

    st.markdown("---")
    comments = st.text_area(T["comments_label"])

    # Math Scoring Parameters Matrix Engine
    st.markdown("---")
    st.subheader(T["metrics_header"])
    
    answers = list(responses.values())
    yes_count = answers.count("YES")
    no_count = answers.count("NO")
    na_count = answers.count("N/A")
    
    applicable_count = len(T["tasks"]) - na_count
    base_score = (yes_count / applicable_count * 100) if applicable_count > 0 else 100.0
    
    # Capping evaluation metric penalty rule (Task #8 Check)
    is_capped = False
    if responses["task_8"] == "NO" and base_score > 85.0:
        base_score = 85.0
        is_capped = True
        
    # Render Dashboard Metrics tags
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric(label=T["m_status"], value=T["m_complete"] if (yes_count + no_count + na_count == 14) else T["m_incomplete"])
    with col_m2:
        if is_capped:
            st.metric(label=T["m_score"], value=f"{base_score:.1f}%", delta=f"- {T['m_capped']}", delta_color="inverse")
        else:
            st.metric(label=T["m_score"], value=f"{base_score:.1f}%")
    with col_m3:
        st.metric(label=T["m_passed"], value=f"{yes_count} / {applicable_count}")

    # E-Signature Tracker Block Configuration
    st.markdown("---")
    st.subheader(T["sign_header"])
    typed_signature = st.text_input(T["sign_label"])
    attestation_check = st.checkbox(T["attest_label"])
    
    if st.button(T["submit_btn"]):
        if typed_signature.strip() == "" or not attestation_check:
            st.error(T["err_missing"])
        else:
            # Package structural data file entry payload
            payload = {
                "Timestamp": [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                "Year": [filter_year],
                "Month": [filter_month],
                "User": [typed_signature.strip()],
                "CMO_ID": [selected_cmo],
                "Final_Score": [f"{base_score:.1f}%"],
                "Capped_Flag": ["YES" if is_capped else "NO"],
                "Comments": [comments]
            }
            # Append separate task variables to row
            for i in range(1, 15):
                payload[f"Task_{i}"] = [responses[f"task_{i}"]]
                
            new_row_df = pd.DataFrame(payload)
            save_path = "data/submissions.csv"
            os.makedirs("data", exist_ok=True)
            
            if not os.path.isfile(save_path):
                new_row_df.to_csv(save_path, index=False)
            else:
                new_row_df.to_csv(save_path, mode='a', header=False, index=False)
                
            st.success(f"{T['success_log']} : {typed_signature} !")

# ==========================================
# TAB 2: HISTORY RECORD LOOKUP & SEARCH VIEW
# ==========================================
with tab2:
    st.subheader(T["history_title"])
    save_path = "data/submissions.csv"
    
    if not os.path.isfile(save_path):
        st.info(T["no_history"])
    else:
        # Load the recorded historical logs
        history_df = pd.read_csv(save_path)
        
        # Apply reactive query searches from sidebar filters dynamically
        if filter_user.strip() != "":
            # Search matches by text character sequence (case insensitive)
            history_df = history_df[history_df['User'].str.contains(filter_user, case=False, na=False)]
            
        # Filter down by Year and Month columns
        history_df = history_df[(history_df['Year'] == filter_year)]
        
        # Display matching rows in an interactive table view grid
        if history_df.empty:
            st.warning("No records found matching current search criteria.")
        else:
            st.dataframe(history_df, use_container_width=True)
