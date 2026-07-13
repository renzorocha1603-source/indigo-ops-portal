import streamlit as st
import pandas as pd
import datetime
import os
import io
from PIL import Image

# Secure ReportLab Import
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
except ImportError:
    os.system('pip install reportlab')
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors

# ==========================================
# 1. PAGE CONFIGURATION & BRAND STYLING
# ==========================================
st.set_page_config(page_title="Indigo Park Ops Portal", layout="wide", page_icon="🅿️")

st.markdown("""
    <style>
    .reportview-container .main .block-container{ max-width: 1200px; }
    h1, h2, h3 { color: #2D144B; } 
    .stButton>button { background-color: #E00073; color: white; border-radius: 6px; font-weight: bold; }
    div.stRadio > div{ flex-direction:row; gap: 20px; } 
    .reason-box { padding: 5px 10px; background-color: #FFF2F7; border-left: 3px solid #E00073; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. BILINGUAL TRANSLATION DICTIONARY
# ==========================================
LANG_DICT = {
    "English": {
        "title": "Operations & Compliance Portal",
        "subtitle": "Montreal Region Lot Management & Best Practices Digital Tracker",
        "sidebar_lang": "🌐 Select Language / Choisir la langue",
        "tab_new": "📝 File New Report",
        "tab_history": "📜 View & Search History",
        "select_cmo": "Select Target Lot ID (CMO):",
        "select_year": "Select Report Target Year:",
        "select_month": "Select Report Target Month:",
        "form_header": "Monthly Validation Form",
        "form_instruction": "Fill out the status for all 14 mandatory compliance items:",
        "comment_placeholder": "Provide specific justification details for this indicator...",
        "comment_warning": "⚠️ Please fill out all context comments for items marked NO or N/A.",
        "metrics_header": "📊 Performance Metrics Results",
        "m_status": "Form Status",
        "m_complete": "Form Completed",
        "m_incomplete": "Incomplete",
        "m_score": "Operational Performance Score",
        "m_capped": "Penalized Max Capping (No Client Meeting)",
        "m_passed": "Validated Tasks (YES)",
        "sign_header": "🖋️ Digital Signature & Verification",
        "sign_label": "Type your Full Name to sign electronically:",
        "attest_label": "I verify that the answers provided above are accurate and true.",
        "submit_btn": "💾 Save & Sync Report to Database",
        "err_missing": "Submission Blocked: You must enter your name, check the attestation box, and fill out all conditional comments.",
        "success_log": "🎉 Report securely saved to history and signed by",
        "no_history": "No records found in the database yet.",
        "history_title": "🔎 Filter and Search Historical Audits",
        "search_user": "Search by User Name:",
        "search_cmo": "Search by Lot ID (CMO):",
        "search_month": "Search by Month:",
        "search_year": "Search by Year:",
        "all": "All",
        "dl_excel": "📥 Download Filtered Selection as Excel",
        "dl_pdf": "📥 Download Filtered Selection as PDF",
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
        "tab_new": "📝 Remplir un nouveau rapport",
        "tab_history": "📜 Consulter et chercher l'historique",
        "select_cmo": "Sélectionner l'emplacement cible (CMO) :",
        "select_year": "Sélectionner l'année cible du rapport :",
        "select_month": "Sélectionner le mois cible du rapport :",
        "form_header": "Formulaire de validation mensuelle",
        "form_instruction": "Indiquez le statut pour l'ensemble des 14 indicateurs obligatoires :",
        "comment_placeholder": "Fournir les détails de justification pour cet indicateur...",
        "comment_warning": "⚠️ Veuillez remplir tous les commentaires de justification pour les éléments cochés 'NO' ou 'N/A'.",
        "metrics_header": "📊 Résultats Métriques de Performance",
        "m_status": "Statut de saisie",
        "m_complete": "Formulaire Complété",
        "m_incomplete": "Saisie Incomplète",
        "m_score": "Indice de Performance Réel",
        "m_capped": "Pénalité maximum appliquée (Réunion Client manquante)",
        "m_passed": "Tâches Validées (YES)",
        "sign_header": "🖋️ Signature et Validation Légale",
        "sign_label": "Saisissez votre Nom et Prénom complet pour la signature électronique :",
        "attest_label": "Je vérifie que les réponses fournies ci-dessus sont exactes.",
        "submit_btn": "💾 Valider et Enregistrer dans la Base Historique",
        "err_missing": "Action Bloquée : Vous devez saisir votre signature nominative, cocher l'attestation, et remplir tous les commentaires obligatoires.",
        "success_log": "🎉 Rapport synchronisé avec succès ! Enregistré de manière sécurisée par",
        "no_history": "Aucun enregistrement trouvé dans la base de données pour le moment.",
        "history_title": "🔎 Filtrer et chercher dans l'historique des audits",
        "search_user": "Chercher par nom d'utilisateur :",
        "search_cmo": "Chercher par No de Lot (CMO) :",
        "search_month": "Chercher par Mois :",
        "search_year": "Chercher par Année :",
        "all": "Tous",
        "dl_excel": "📥 Télécharger la sélection sous Excel",
        "dl_pdf": "📥 Télécharger la sélection sous PDF",
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

# Language Selector Sidebar
selected_lang = st.sidebar.selectbox(LANG_DICT["English"]["sidebar_lang"], ["Français", "English"])
T = LANG_DICT[selected_lang]

# ==========================================
# 3. HEADER LAYOUT WITH FIXED CORPORATE LOGO
# ==========================================
col_logo, col_title = st.columns([1, 4])
with col_logo:
    try:
        st.image(Image.open("indigo-park-canada-logo.jpg"), width=180)
    except:
        st.write("🅿️ **INDIGO PARK**")
with col_title:
    st.title(T["title"])
    st.caption(T["subtitle"])

st.markdown("---")

# ==========================================
# 4. EXCEL CMO DATA ENGINE
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

years_options = [2026, 2025, 2024]
months_options = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"] if selected_lang == "Français" else ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

tab1, tab2 = st.tabs([T["tab_new"], T["tab_history"]])

# ==========================================
# TAB 1: NEW AUDIT FORM
# ==========================================
with tab1:
    col_cmo, col_yr, col_mnth = st.columns(3)
    with col_cmo:
        selected_cmo = st.selectbox(T["select_cmo"], cmo_options)
    with col_yr:
        chosen_year = st.selectbox(T["select_year"], years_options, index=0)
    with col_mnth:
        chosen_month = st.selectbox(T["select_month"], months_options, index=6) # Default to July

    st.subheader(f"{T['form_header']} : {selected_cmo} — ({chosen_month} {chosen_year})")
    st.write(T["form_instruction"])
    
    responses = {}
    task_comments = {}
    missing_comments_flag = False
    
    for idx, task in enumerate(T["tasks"], start=1):
        st.markdown(f"**{idx}. {task}**")
        responses[f"task_{idx}"] = st.radio(
            f"Statut {idx}", ["YES", "NO", "N/A"], key=f"task_radio_{idx}", label_visibility="collapsed"
        )
        
        if responses[f"task_{idx}"] in ["NO", "N/A"]:
            st.markdown(f"<div class='reason-box'>", unsafe_allow_html=True)
            task_comments[f"comment_{idx}"] = st.text_input(
                f"Justification / Reason ({responses[f'task_{idx}']}) *",
                key=f"comment_input_{idx}",
                placeholder=T["comment_placeholder"]
            )
            st.markdown("</div>", unsafe_allow_html=True)
            if not task_comments[f"comment_{idx}"].strip():
                missing_comments_flag = True
        else:
            task_comments[f"comment_{idx}"] = ""
            
        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

    # Scoring calculation
    st.markdown("---")
    st.subheader(T["metrics_header"])
    
    answers = list(responses.values())
    yes_count = answers.count("YES")
    no_count = answers.count("NO")
    na_count = answers.count("N/A")
    
    applicable_count = len(T["tasks"]) - na_count
    base_score = (yes_count / applicable_count * 100) if applicable_count > 0 else 100.0
    
    is_capped = False
    if responses["task_8"] == "NO" and base_score > 85.0:
        base_score = 85.0
        is_capped = True
        
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

    st.markdown("---")
    st.subheader(T["sign_header"])
    typed_signature = st.text_input(T["sign_label"])
    attestation_check = st.checkbox(T["attest_label"])
    
    if missing_comments_flag:
        st.warning(T["comment_warning"])
        
    if st.button(T["submit_btn"]):
        if typed_signature.strip() == "" or not attestation_check or missing_comments_flag:
            st.error(T["err_missing"])
        else:
            payload = {
                "Timestamp": [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                "Year": [chosen_year],
                "Month": [chosen_month],
                "User": [typed_signature.strip()],
                "CMO_ID": [selected_cmo],
                "Final_Score": [f"{base_score:.1f}%"],
                "Capped_Flag": ["YES" if is_capped else "NO"]
            }
            for i in range(1, 15):
                payload[f"Task_{i}_Status"] = [responses[f"task_{i}"]]
                payload[f"Task_{i}_Comment"] = [task_comments[f"comment_{i}"]]
                
            new_row_df = pd.DataFrame(payload)
            save_path = "data/submissions.csv"
            os.makedirs("data", exist_ok=True)
            
            if not os.path.isfile(save_path):
                new_row_df.to_csv(save_path, index=False)
            else:
                new_row_df.to_csv(save_path, mode='a', header=False, index=False)
                
            st.success(f"{T['success_log']} : {typed_signature} !")

# ==========================================
# TAB 2: HISTORY - WITH IN-PAGE LIVE FILTERS
# ==========================================
with tab2:
    st.subheader(T["history_title"])
    save_path = "data/submissions.csv"
    
    if not os.path.isfile(save_path):
        st.info(T["no_history"])
    else:
        # Always pull clean ground truth from database file
        history_df = pd.read_csv(save_path)
        
        # In-page search configuration rows
        h_col1, h_col2, h_col3, h_col4 = st.columns(4)
        with h_col1:
            s_user = st.text_input(T["search_user"], value="")
        with h_col2:
            s_cmo = st.text_input(T["search_cmo"], value="")
        with h_col3:
            s_month = st.selectbox(T["search_month"], [T["all"]] + months_options)
        with h_col4:
            s_year = st.selectbox(T["search_year"], [T["all"]] + [str(y) for y in years_options])
            
        # Process user data matches dynamically
        filtered_df = history_df.copy()
        if s_user.strip():
            filtered_df = filtered_df[filtered_df['User'].str.contains(s_user, case=False, na=False)]
        if s_cmo.strip():
            filtered_df = filtered_df[filtered_df['CMO_ID'].str.contains(s_cmo, case=False, na=False)]
        if s_month != T["all"]:
            filtered_df = filtered_df[filtered_df['Month'] == s_month]
        if s_year != T["all"]:
            filtered_df = filtered_df[filtered_df['Year'] == int(s_year)]
            
        if filtered_df.empty:
            st.warning("No recorded matrix logs found matching these exact search parameters.")
        else:
            st.dataframe(filtered_df, use_container_width=True)
            
            st.markdown("---")
            dl_col1, dl_col2 = st.columns(2)
            
            with dl_col1:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Filtered Ops Data')
                
                st.download_button(
                    label=T["dl_excel"],
                    data=excel_buffer.getvalue(),
                    file_name=f"Indigo_Filtered_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with dl_col2:
                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
                styles = getSampleStyleSheet()
                
                story = []
                story.append(Paragraph(f"<b>INDIGO PARK CANADA — OPÉRATIONS HISTORIQUE EXPORT</b>", styles['Heading1']))
                story.append(Paragraph(f"Exported On: {datetime.date.today().strftime('%Y-%m-%d')}", styles['Normal']))
                story.append(Spacer(1, 15))
                
                table_data = [["CMO ID", "Auditor User", "Month", "Year", "Score"]]
                for _, row in filtered_df.iterrows():
                    table_data.append([str(row['CMO_ID']), str(row['User']), str(row['Month']), str(row['Year']), str(row['Final_Score'])])
                
                pdf_table = Table(table_data, colWidths=[100, 150, 100, 80, 100])
                pdf_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2D144B')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0,0), (-1,0), 8),
                    ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F9F9F9')),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ]))
                
                story.append(pdf_table)
                doc.build(story)
                
                st.download_button(
                    label=T["dl_pdf"],
                    data=pdf_buffer.getvalue(),
                    file_name=f"Indigo_Filtered_Summary.pdf",
                    mime="application/pdf"
                )
