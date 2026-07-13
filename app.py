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
# 1. CONFIGURATION DE LA PAGE & DESIGN EXCEL/INDIGO
# ==========================================
st.set_page_config(page_title="Indigo Park City Matrix Portal", layout="wide", page_icon="🅿️")

# Injection CSS pour donner un aspect "Tableau de bord / Feuille Excel"
st.markdown("""
    <style>
    /* Élargir le conteneur principal */
    .block-container { max-width: 95% !important; padding-top: 1.5rem !important; }
    
    /* Titres corporatifs */
    h1, h2, h3 { color: #2D144B; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    
    /* Boutons style Excel / Indigo */
    .stButton>button { background-color: #E00073; color: white; border-radius: 4px; font-weight: 600; padding: 0.5rem 1.5rem; border: none; }
    .stButton>button:hover { background-color: #2D144B; color: white; }
    
    /* Zone d'alerte/justification épurée */
    .reason-box { padding: 10px; background-color: #F9F9F9; border-left: 4px solid #E00073; margin-top: 5px; margin-bottom: 10px; border-radius: 4px; }
    
    /* Rendre les radios horizontaux plus compacts */
    div.stRadio > div { flex-direction: row; gap: 25px; }
    
    /* Style pour simuler les onglets Excel */
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
        "tab_new": "📝 Fill Matrix Form",
        "tab_history": "📜 Database Audit Log",
        "tab_ref1": "📋 Activity Management",
        "tab_ref2": "📋 Standards Overview",
        "tab_ref3": "📋 Best Practices Matrix",
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
        "tab_new": "📝 Saisie de la Matrice",
        "tab_history": "📜 Historique des Données",
        "tab_ref1": "📋 Gestion des Activités (Réf)",
        "tab_ref2": "📋 Aperçu des Normes (Réf)",
        "tab_ref3": "📋 Meilleures Pratiques (Réf)",
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

# Configuration de la barre supérieure pour le choix de la langue
col_header_logo, col_header_space, col_header_lang = st.columns([1.5, 4, 1.2])
with col_header_logo:
    st.image(logo_url, width=160)
with col_header_lang:
    selected_lang = st.selectbox("🌐 Language / Langue", ["Français", "English"], label_visibility="collapsed")

T = LANG_DICT[selected_lang]

st.title(T["title"])
st.caption(T["subtitle"])
st.markdown("---")

# ==========================================
# 3. BASE DE DONNÉES DE SECOURS (FALLBACK)
# ==========================================
FALLBACK_DATA = {
    'Gestion des activités': pd.DataFrame({
        "Tâche [Niveau 1 System]": ["Mettre à jour la matrice de rapports municipaux", "Examiner la matrice de rapports municipaux", "Compléter la liste des lotissements", "Effectuer les tâches terrain"],
        "Responsabilité": ["RR (Responsable Régional)", "DG (Directeur Général)", "DG", "RR et Superviseur"],
        "Fréquence / Échéance": ["Mensuelle", "Révision le 21 du mois, finale le 1er du mois suivant", "Révision mensuelle", "Mensuelle"]
    }),
    'Aperçu des normes': pd.DataFrame({
        "Excellence (Quoi)": ["1. Personnalisation & Customisation", "2. Réunions Clients / Comptes rendus", "3. Signalisation & Image de marque", "10. Professionnalisme Niveau 1"],
        "Résultats attendus (Comment)": ["Prendre des notes détaillées sur chaque réunion, carnets distincts", "Le DG doit assister au moins une fois par trimestre aux réunions mensuelles", "Niveaux de livraison les plus élevés, vérifier et remplacer régulièrement", "S'assurer que l'équipe premium est formée à l'étiquette de haut niveau"],
        "Responsabilité": ["Opérations juniors / Superviseur", "Directeur Général", "Responsable Opérationnel / DG", "Directeur Général / Tous"]
    }),
    'Matrice des meilleures pratique': pd.DataFrame({
        "Indicateurs Opérationnels": ["Rapport opérationnel complet du site", "CRITICAL: Réunion/appel mensuel programmé", "Audit mensuel SMILE", "Propositions à valeur ajoutée (Revenus/Économies)"],
        "Niveau Minimum Requis": ["Niveau 1, 2 et 3 - Mensuel", "Niveau 1 - Obligatoire (Plafonnement si absent)", "Niveau 1, 2 - Mensuel", "Niveau 1 - Mensuel / Trimestriel"]
    })
}

def get_sheet_data(sheet_name, fallback_key):
    if master_excel_file and os.path.exists(master_excel_file):
        try:
            return pd.read_excel(master_excel_file, sheet_name=sheet_name)
        except:
            pass
    return FALLBACK_DATA[fallback_key]

# Dynamic CMO collection
def load_cmo_codes():
    if master_excel_file and os.path.exists(master_excel_file):
        try:
            df = pd.read_excel(master_excel_file, sheet_name='City Reporting Matrix 2026', skiprows=9)
            df.columns = df.columns.str.strip()
            raw_codes = df.iloc[0].dropna().astype(str).str.strip().tolist()
            cmo_list = [x for x in raw_codes if x.upper().startswith(('CMO', 'VMO'))]
            if cmo_list:
                return sorted(list(set(cmo_list)))
        except:
            pass
    return [f"CMO{str(i).zfill(3)}" for i in [2, 20, 37, 101, 102, 108, 111, 119, 132, 141, 145, 146, 242, 275, 296, 305]]

cmo_options = load_cmo_codes()
# Mise à jour de la plage : de 2024 jusqu'à 2035
years_options = list(range(2024, 2036))
months_options = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"] if selected_lang == "Français" else ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

# ==========================================
# 4. SYSTÈME D'ONGLETS HORIZONTAUX (STYLE EXCEL)
# ==========================================
t_form, t_hist, t_ref1, t_ref2, t_ref3 = st.tabs([
    T["tab_new"], T["tab_history"], T["tab_ref1"], T["tab_ref2"], T["tab_ref3"]
])

# Onglet 1 : Formulaire de saisie
with t_form:
    if st.session_state.success_message:
        st.success(st.session_state.success_message)
        st.balloons()
        st.session_state.success_message = None

    # En-tête du formulaire style tableur
    col_cmo, col_yr, col_mnth = st.columns([2, 1, 1])
    with col_cmo:
        selected_cmo = st.selectbox(T["select_cmo"], cmo_options)
    with col_yr:
        chosen_year = st.selectbox(T["select_year"], years_options, index=0)
    with col_mnth:
        chosen_month = st.selectbox(T["select_month"], months_options, index=6)

    st.markdown(f"### {T['form_header']} — {selected_cmo} ({chosen_month} {chosen_year})")
    
    responses = {}
    task_comments = {}
    incomplete_selections_flag = False
    missing_comments_flag = False
    
    st.markdown("---")
    for i, task in enumerate(LANG_DICT["English"]["tasks"], start=1):
        col_task_text, col_task_radio = st.columns([5, 2])
        with col_task_text:
            st.markdown(f"**{i}.** {T['tasks'][i-1]}")
        with col_task_radio:
            responses[f"task_{i}"] = st.radio(
                f"Radio_{i}", ["YES", "NO", "N/A"], index=None, key=f"radio_task_{i}", label_visibility="collapsed"
            )
            
        if responses[f"task_{i}"] is None:
            incomplete_selections_flag = True
            task_comments[f"comment_{i}"] = ""
        elif responses[f"task_{i}"] in ["NO", "N/A"]:
            with col_task_text:
                st.markdown(f"<div class='reason-box'>", unsafe_allow_html=True)
                task_comments[f"comment_{i}"] = st.text_input(
                    f"Justification {responses[f'task_{i}']} *",
                    key=f"comm_{i}",
                    placeholder=T["comment_placeholder"],
                    label_visibility="collapsed"
                )
                st.markdown("</div>", unsafe_allow_html=True)
                if not task_comments[f"comment_{i}"].strip():
                    missing_comments_flag = True
        else:
            task_comments[f"comment_{i}"] = ""
        st.markdown("<hr style='margin:0.5rem 0; border:0; border-top:1px solid #F00;' />", unsafe_allow_html=True)

    # Calculs Métriques
    answers = [v for v in responses.values() if v is not None]
    yes_count = answers.count("YES")
    no_count = answers.count("NO")
    na_count = answers.count("N/A")
    
    if not incomplete_selections_flag:
        applicable_count = 9 - na_count
        base_score = (yes_count / applicable_count * 100) if applicable_count > 0 else 100.0
        is_capped = False
        if responses["task_2"] == "NO" and base_score > 85.0:
            base_score = 85.0
            is_capped = True
        score_display = f"{base_score:.1f}%"
        status_display = T["m_complete"]
    else:
        is_capped = False
        score_display = "--"
        status_display = T["m_incomplete"]

    st.markdown(f"### {T['metrics_header']}")
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric(label=T["m_status"], value=status_display)
    with col_m2:
        if is_capped:
            st.metric(label=T["m_score"], value=score_display, delta=T["m_capped"], delta_color="inverse")
        else:
            st.metric(label=T["m_score"], value=score_display)
    with col_m3:
        st.metric(label=T["m_passed"], value=f"{yes_count} / {9 - na_count if not incomplete_selections_flag else 9}")

    st.markdown("---")
    st.markdown(f"### {T['sign_header']}")
    col_sig, col_att = st.columns([2, 2])
    with col_sig:
        typed_signature = st.text_input(T["sign_label"])
    with col_att:
        st.markdown("<br>", unsafe_allow_html=True)
        attestation_check = st.checkbox(T["attest_label"])
        
    if incomplete_selections_flag or missing_comments_flag:
        st.warning(T["comment_warning"])
        
    if st.button(T["submit_btn"]):
        if incomplete_selections_flag or missing_comments_flag or typed_signature.strip() == "" or not attestation_check:
            st.error(T["err_missing"])
        else:
            unique_report_id = f"REF-{int(datetime.datetime.now().timestamp())}"
            payload = {
                "Report_ID": [unique_report_id],
                "Timestamp": [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                "Year": [chosen_year],
                "Month": [chosen_month],
                "User": [typed_signature.strip()],
                "CMO_ID": [selected_cmo],
                "Final_Score": [score_display],
                "Capped_Flag": ["YES" if is_capped else "NO"]
            }
            for i in range(1, 10):
                payload[f"Q{i}_Task_Text"] = [LANG_DICT["English"]["tasks"][i-1]]
                payload[f"Q{i}_Status"] = [responses[f"task_{i}"]]
                payload[f"Q{i}_Justification"] = [task_comments[f"comment_{i}"]]
                
            new_row_df = pd.DataFrame(payload)
            save_path = "data/submissions.csv"
            os.makedirs("data", exist_ok=True)
            if not os.path.isfile(save_path):
                new_row_df.to_csv(save_path, index=False)
            else:
                new_row_df.to_csv(save_path, mode='a', header=False, index=False)
                
            st.session_state.success_message = f"{T['success_log']} : **{typed_signature}** ! ID: {unique_report_id}"
            st.rerun()

# Onglet 2 : Historique et téléchargement
with t_hist:
    st.markdown(f"### {T['history_title']}")
    save_path = "data/submissions.csv"
    
    if not os.path.isfile(save_path) or os.path.getsize(save_path) == 0:
        st.info(T["no_history"])
    else:
        history_df = pd.read_csv(save_path)
        h_col1, h_col2, h_col3, h_col4 = st.columns(4)
        with h_col1:
            s_user = st.text_input(T["search_user"], key="filter_user")
        with h_col2:
            s_cmo = st.text_input(T["search_cmo"], key="filter_cmo")
        with h_col3:
            s_month = st.selectbox(T["search_month"], [T["all"]] + months_options, key="filter_month")
        with h_col4:
            s_year = st.selectbox(T["search_year"], [T["all"]] + [str(y) for y in years_options], key="filter_year")
            
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
            st.warning("Aucun enregistrement trouvé.")
        else:
            st.dataframe(filtered_df[["Report_ID", "Timestamp", "CMO_ID", "User", "Month", "Year", "Final_Score", "Capped_Flag"]], use_container_width=True)
            
            st.markdown("---")
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                vertical_df = filtered_df.set_index("Report_ID").transpose()
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    vertical_df.to_excel(writer, sheet_name='Audit Log')
                st.download_button(label=T["dl_excel"], data=excel_buffer.getvalue(), file_name="Indigo_Vertical_Matrix_Log.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
            with dl_col2:
                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                styles = getSampleStyleSheet()
                style_normal = ParagraphStyle(name='WrapText', parent=styles['Normal'], fontSize=8, leading=10)
                style_header = ParagraphStyle(name='HeaderStyle', parent=styles['Normal'], fontSize=8, leading=10, fontName='Helvetica-Bold', textColor=colors.whitesmoke)
                
                story = [Paragraph("<b>INDIGO PARK CANADA — COMPLIANCE LOG REPORT</b>", styles['Heading2']), Spacer(1, 15)]
                for idx, row in filtered_df.iterrows():
                    story.append(Paragraph(f"<b>Matrix Profile: {row['CMO_ID']} — {row['Month']} {row['Year']}</b> (Score: {row['Final_Score']})", styles['Heading3']))
                    pdf_table_data = [[Paragraph("No.", style_header), Paragraph("Indicator", style_header), Paragraph("Status", style_header), Paragraph("Justification / Comment", style_header)]]
                    for i in range(1, 10):
                        pdf_table_data.append([Paragraph(str(i), style_normal), Paragraph(str(row[f"Q{i}_Task_Text"]), style_normal), Paragraph(str(row[f"Q{i}_Status"]), style_normal), Paragraph(str(row[f"Q{i}_Justification"]) if pd.notna(row[f"Q{i}_Justification"]) else "", style_normal)])
                    t = Table(pdf_table_data, colWidths=[25, 275, 50, 200])
                    t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2D144B')), ('VALIGN', (0,0), (-1,-1), 'TOP'), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey)]))
                    story.append(t)
                    story.append(Spacer(1, 15))
                doc.build(story)
                st.download_button(label=T["dl_pdf"], data=pdf_buffer.getvalue(), file_name="Indigo_Compliance_Report.pdf", mime="application/pdf")

# Onglets Références en Pleine Page (Style Excel original)
with t_ref1:
    st.markdown("### 📋 Gestion des activités des clients de niveau 1")
    st.dataframe(get_sheet_data('Gestion des activités', 'Gestion des activités'), use_container_width=True)

with t_ref2:
    st.markdown("### 📋 Procédures Standards D'Opération — Aperçu des normes")
    st.dataframe(get_sheet_data('Aperçu des normes', 'Aperçu des normes'), use_container_width=True)

with t_ref3:
    st.markdown("### 📋 Meilleures pratiques opérationnelles — Niveaux minimaux d'activité locale")
    st.dataframe(get_sheet_data('Matrice des meilleures pratique', 'Matrice des meilleures pratique'), use_container_width=True)
