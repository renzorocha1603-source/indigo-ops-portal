import streamlit as st
import pandas as pd
import datetime
import os
import io

# Importation sécurisée de ReportLab pour les exports PDF
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
except ImportError:
    os.system('pip install reportlab')
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

# ==========================================
# 1. CONFIGURATION DE LA PAGE
# ==========================================
st.set_page_config(page_title="Indigo Park City Matrix Portal", layout="wide", page_icon="🅿️")

st.markdown("""
    <style>
    .block-container { max-width: 95% !important; padding-top: 1.5rem !important; }
    h1, h2, h3 { color: #2D144B; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stButton>button { background-color: #E00073; color: white; border-radius: 4px; font-weight: 600; padding: 0.5rem 1.5rem; border: none; width: 100%; }
    .stButton>button:hover { background-color: #2D144B; color: white; }
    .reason-box { padding: 10px; background-color: #F9F9F9; border-left: 4px solid #E00073; margin-top: 5px; margin-bottom: 10px; border-radius: 4px; }
    div.stRadio > div { flex-direction: row; gap: 25px; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; border-bottom: 2px solid #E0E0E0; }
    .stTabs [data-baseweb="tab"] { background-color: #F4F5F7; border: 1px solid #E0E0E0; border-bottom: none; padding: 8px 20px; border-radius: 4px 4px 0px 0px; font-weight: bold; color: #555555; }
    .stTabs [aria-selected="true"] { background-color: #2D144B !important; color: white !important; border-color: #2D144B !important; }
    </style>
""", unsafe_allow_html=True)

# Détermination dynamique du fichier Excel
possible_paths = ['/home/bard/Montreal Lot List.xlsx', 'Montreal Lot List.xlsx', os.path.abspath('Montreal Lot List.xlsx')]
master_excel_file = next((p for p in possible_paths if os.path.exists(p)), 'Montreal Lot List.xlsx')
logo_url = "https://i.ibb.co/DHgswzDq/indigo-park-canada-logo.jpg"

if "success_message" not in st.session_state:
    st.session_state.success_message = None

# ==========================================
# 2. DICTIONNAIRE BILINGUE STRUCTURÉ
# ==========================================
LANG_DICT = {
    "English": {
        "title": "City Reporting Matrix Portal",
        "subtitle": "Montreal Region Compliance Tracker & Reference Hub (2026)",
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
        "comment_warning": "⚠️ Complete all 9 items and fill mandatory justifications for NO or N/A choices.",
        "metrics_header": "📊 KPI Performance Metrics",
        "m_status": "Form Status",
        "m_complete": "Completed",
        "m_incomplete": "Incomplete",
        "m_score": "Performance Score",
        "m_capped": "Max Cap Applied (No Client Meeting)",
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
        "subtitle": "Suivi de conformité Région de Montréal & Centre de Références (2026)",
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
        "m_capped": "Pénalité maximum appliquée (Réunion Client manquante)",
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
# 3. SIDEBAR & LOGIQUE DYNAMIQUE
# ==========================================
with st.sidebar:
    st.image(logo_url, width=150)
    selected_lang = st.selectbox("🌐 Language / Langue", ["Français", "English"])
    T = LANG_DICT[selected_lang]
    st.markdown("---")
    st.markdown(f"### {T['sidebar_ref']}")
    
    # Helper pour bouton de téléchargement Excel
    def excel_dl_button(sheet_name, label):
        if os.path.exists(master_excel_file):
            df = pd.read_excel(master_excel_file, sheet_name=sheet_name)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            st.download_button(label=label, data=buffer.getvalue(), file_name=f"{sheet_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("File not found")

    excel_dl_button('Gestion des activités', T['ref_btn1'])
    excel_dl_button('Aperçu des normes', T['ref_btn2'])
    excel_dl_button('Matrice des meilleures pratique', T['ref_btn3'])

st.title(T["title"])
st.caption(T["subtitle"])
st.markdown("---")

def get_customer_name(cmo_id):
    if master_excel_file and os.path.exists(master_excel_file):
        try:
            df = pd.read_excel(master_excel_file, sheet_name='City Reporting Matrix 2026', skiprows=9)
            df.columns = df.columns.str.strip()
            if 'CMO' in df.columns and 'Customer Name' in df.columns:
                result = df.loc[df['CMO'] == cmo_id, 'Customer Name']
                if not result.empty: return result.values[0]
        except Exception: pass
    return "N/A"

def load_cmo_codes():
    if master_excel_file and os.path.exists(master_excel_file):
        try:
            df = pd.read_excel(master_excel_file, sheet_name='City Reporting Matrix 2026', skiprows=9)
            df.columns = df.columns.str.strip()
            cmo_list = df['CMO'].dropna().astype(str).str.strip().tolist()
            if cmo_list: return sorted(list(set(cmo_list)))
        except: pass
    return [f"CMO{str(i).zfill(3)}" for i in [2, 20, 37, 101, 102, 108, 111, 119, 132, 141, 145, 146, 242, 275, 296, 305]]

cmo_options = load_cmo_codes()
years_options = list(range(2024, 2036))
months_options = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"] if selected_lang == "Français" else ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

# ==========================================
# 4. ONGLET PRINCIPAL
# ==========================================
t_form, t_hist = st.tabs([T["tab_new"], T["tab_history"]])

with t_form:
    if st.session_state.success_message:
        st.success(st.session_state.success_message)
        st.balloons()
        st.session_state.success_message = None

    col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
    with col1: selected_cmo = st.selectbox(T["select_cmo"], cmo_options)
    with col2: st.text_input("Customer Name (Info):", value=str(get_customer_name(selected_cmo)), disabled=True)
    with col3: chosen_year = st.selectbox(T["select_year"], years_options, index=0)
    with col4: chosen_month = st.selectbox(T["select_month"], months_options, index=6)

    st.markdown(f"### {T['form_header']} — {selected_cmo} ({chosen_month} {chosen_year})")
    responses, task_comments = {}, {}
    incomplete_selections_flag, missing_comments_flag = False, False
    st.markdown("---")
    
    for i, task in enumerate(LANG_DICT["English"]["tasks"], start=1):
        col_task_text, col_task_radio = st.columns([5, 2])
        with col_task_text: st.markdown(f"**{i}.** {T['tasks'][i-1]}")
        with col_task_radio:
            responses[f"task_{i}"] = st.radio(f"Radio_{i}", ["YES", "NO", "N/A"], index=None, key=f"radio_task_{i}", label_visibility="collapsed")
            
        if responses[f"task_{i}"] is None: incomplete_selections_flag = True
        elif responses[f"task_{i}"] in ["NO", "N/A"]:
            with col_task_text:
                st.markdown(f"<div class='reason-box'>", unsafe_allow_html=True)
                task_comments[f"comment_{i}"] = st.text_input(f"Justification {responses[f'task_{i}']} *", key=f"comm_{i}", placeholder=T["comment_placeholder"], label_visibility="collapsed")
                st.markdown("</div>", unsafe_allow_html=True)
                if not task_comments[f"comment_{i}"].strip(): missing_comments_flag = True
        else: task_comments[f"comment_{i}"] = ""
        st.markdown("<hr style='margin:0.5rem 0; border:0; border-top:1px solid #F00;' />", unsafe_allow_html=True)

    answers = [v for v in responses.values() if v is not None]
    yes_count = answers.count("YES")
    na_count = answers.count("N/A")
    if not incomplete_selections_flag:
        applicable_count = 9 - na_count
        base_score = (yes_count / applicable_count * 100) if applicable_count > 0 else 100.0
        is_capped = (responses["task_2"] == "NO" and base_score > 85.0)
        if is_capped: base_score = 85.0
        score_display, status_display = f"{base_score:.1f}%", T["m_complete"]
    else: is_capped, score_display, status_display = False, "--", T["m_incomplete"]

    st.markdown(f"### {T['metrics_header']}")
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric(label=T["m_status"], value=status_display)
    c_m2.metric(label=T["m_score"], value=score_display, delta=T["m_capped"] if is_capped else None, delta_color="inverse")
    c_m3.metric(label=T["m_passed"], value=f"{yes_count} / {9 - na_count if not incomplete_selections_flag else 9}")

    st.markdown("---")
    st.markdown(f"### {T['sign_header']}")
    c_sig, c_att = st.columns([2, 2])
    typed_signature = c_sig.text_input(T["sign_label"])
    attestation_check = c_att.checkbox(T["attest_label"])
        
    if incomplete_selections_flag or missing_comments_flag: st.warning(T["comment_warning"])
    if st.button(T["submit_btn"]):
        if incomplete_selections_flag or missing_comments_flag or typed_signature.strip() == "" or not attestation_check:
            st.error(T["err_missing"])
        else:
            unique_report_id = f"REF-{int(datetime.datetime.now().timestamp())}"
            payload = {"Report_ID": [unique_report_id], "Timestamp": [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")], "Year": [chosen_year], "Month": [chosen_month], "User": [typed_signature.strip()], "CMO_ID": [selected_cmo], "Final_Score": [score_display], "Capped_Flag": ["YES" if is_capped else "NO"]}
            for i in range(1, 10):
                payload[f"Q{i}_Task_Text"] = [LANG_DICT["English"]["tasks"][i-1]]
                payload[f"Q{i}_Status"] = [responses[f"task_{i}"]]
                payload[f"Q{i}_Justification"] = [task_comments[f"comment_{i}"]]
            new_row_df = pd.DataFrame(payload)
            os.makedirs("data", exist_ok=True)
            if not os.path.isfile("data/submissions.csv"): new_row_df.to_csv("data/submissions.csv", index=False)
            else: new_row_df.to_csv("data/submissions.csv", mode='a', header=False, index=False)
            st.session_state.success_message = f"{T['success_log']} : **{typed_signature}** ! ID: {unique_report_id}"
            st.rerun()

with t_hist:
    st.markdown(f"### {T['history_title']}")
    if not os.path.isfile("data/submissions.csv"): st.info(T["no_history"])
    else:
        df = pd.read_csv("data/submissions.csv")
        h1, h2, h3, h4 = st.columns(4)
        s_user = h1.text_input(T["search_user"], key="filter_user")
        s_cmo = h2.text_input(T["search_cmo"], key="filter_cmo")
        s_month = h3.selectbox(T["search_month"], [T["all"]] + months_options, key="filter_month")
        s_year = h4.selectbox(T["search_year"], [T["all"]] + [str(y) for y in years_options], key="filter_year")
        f_df = df.copy()
        if s_user.strip(): f_df = f_df[f_df['User'].str.contains(s_user, case=False, na=False)]
        if s_cmo.strip(): f_df = f_df[f_df['CMO_ID'].str.contains(s_cmo, case=False, na=False)]
        if s_month != T["all"]: f_df = f_df[f_df['Month'] == s_month]
        if s_year != T["all"]: f_df = f_df[f_df['Year'] == int(s_year)]
        if f_df.empty: st.warning("Aucun enregistrement trouvé.")
        else:
            st.dataframe(f_df[["Report_ID", "Timestamp", "CMO_ID", "User", "Month", "Year", "Final_Score", "Capped_Flag"]], use_container_width=True)
            dl1, dl2 = st.columns(2)
            # Logique Export...
            with dl1:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer) as writer: f_df.transpose().to_excel(writer, sheet_name='Audit Log')
                st.download_button(T["dl_excel"], buffer.getvalue(), "Log.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with dl2:
                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
                styles = getSampleStyleSheet()
                story = [Paragraph("<b>INDIGO COMPLIANCE REPORT</b>", styles['Heading2'])]
                for _, row in f_df.iterrows():
                    story.append(Paragraph(f"<b>CMO: {row['CMO_ID']}</b>", styles['Heading3']))
                doc.build(story)
                st.download_button(T["dl_pdf"], pdf_buffer.getvalue(), "Report.pdf", "application/pdf")
