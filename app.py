import streamlit as st
import pandas as pd
import datetime
import os
import io

# Secure ReportLab Component Imports for PDF compilation
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
# 1. PAGE CONFIGURATION & BRAND STYLING
# ==========================================
st.set_page_config(page_title="Indigo Park City Matrix Portal", layout="wide", page_icon="🅿️")

st.markdown("""
    <style>
    .reportview-container .main .block-container{ max-width: 1200px; }
    h1, h2, h3 { color: #2D144B; } 
    .stButton>button { background-color: #E00073; color: white; border-radius: 6px; font-weight: bold; }
    div.stRadio > div{ flex-direction:row; gap: 20px; } 
    .reason-box { padding: 5px 10px; background-color: #FFF2F7; border-left: 3px solid #E00073; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

# Exact match with your Excel spreadsheet filename
master_excel_file = 'Montreal Lot List.xlsx'
excel_exists = os.path.isfile(master_excel_file)

# ==========================================
# 2. BILINGUAL TRANSLATION DICTIONARY
# ==========================================
LANG_DICT = {
    "English": {
        "title": "City Reporting Matrix Portal",
        "subtitle": "Montreal Region Compliance Tracker & Reference Hub",
        "sidebar_lang": "🌐 Language / Langue",
        "sidebar_ref_title": "📚 Reference Documents",
        "sidebar_ref_desc": "Expand to read on-screen or download operational standards sheets:",
        "tab_new": "📝 Fill City Reporting Matrix",
        "tab_history": "📜 View & Manage Historical Submissions",
        "select_cmo": "Select Target Lot ID (CMO):",
        "select_year": "Select Target Matrix Year:",
        "select_month": "Select Target Matrix Month:",
        "form_header": "City Reporting Matrix Log",
        "form_instruction": "Please choose a status for all 14 mandatory operational indicator items:",
        "comment_placeholder": "Provide specific justification details for this indicator...",
        "comment_warning": "⚠️ Please select a choice for all 14 indicators and fill out context comments for items marked NO or N/A.",
        "metrics_header": "📊 Performance Metrics Results",
        "m_status": "Form Status",
        "m_complete": "Form Completed",
        "m_incomplete": "Incomplete / Missing Selections",
        "m_score": "Operational Performance Score",
        "m_capped": "Penalized Max Capping (No Client Meeting)",
        "m_passed": "Validated Tasks (YES)",
        "sign_header": "🖋️ Digital Signature & Verification",
        "sign_label": "Type your Full Name to sign electronically:",
        "attest_label": "I verify that the answers provided above are accurate and true.",
        "submit_btn": "💾 Save & Sync Matrix to Database",
        "err_missing": "Submission Blocked: You must make a selection for every item, fill required comments, type your name, and check the attestation box.",
        "success_log": "🎉 Matrix report securely synchronized by",
        "no_history": "No submitted records found in the database yet.",
        "history_title": "🔎 Filter, Search & Manage Completed Matrices",
        "search_user": "Search by User Name:",
        "search_cmo": "Search by Lot ID (CMO):",
        "search_month": "Search by Month:",
        "search_year": "Search by Year:",
        "all": "All",
        "dl_excel": "📥 Download FULL Comprehensive Report (Excel)",
        "dl_pdf": "📥 Download FULL Comprehensive Report (PDF)",
        "delete_confirm": "Report deleted successfully from history.",
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
        "title": "Portail City Reporting Matrix",
        "subtitle": "Suivi de conformité Région de Montréal & Centre de Références",
        "sidebar_lang": "🌐 Langue / Language",
        "sidebar_ref_title": "📚 Documents de Référence",
        "sidebar_ref_desc": "Déroulez pour lire directement à l'écran ou téléchargez les guides officiels :",
        "tab_new": "📝 Remplir la City Reporting Matrix",
        "tab_history": "📜 Consulter & Gérer l'Historique",
        "select_cmo": "Sélectionner l'emplacement cible (CMO) :",
        "select_year": "Sélectionner l'année cible de la matrice :",
        "select_month": "Sélectionner le mois cible de la matrice :",
        "form_header": "Saisie de la City Reporting Matrix",
        "form_instruction": "Veuillez sélectionner un statut pour chacun des 14 indicateurs obligatoires :",
        "comment_placeholder": "Fournir les détails de justification pour cet indicateur...",
        "comment_warning": "⚠️ Veuillez répondre à toutes les questions et remplir les justifications obligatoires pour les choix 'NO' ou 'N/A'.",
        "metrics_header": "📊 Résultats Métriques de Performance",
        "m_status": "Statut de saisie",
        "m_complete": "Formulaire Complété",
        "m_incomplete": "Sélection Incomplète",
        "m_score": "Indice de Performance Réel",
        "m_capped": "Pénalité maximum appliquée (Réunion Client manquante)",
        "m_passed": "Tâches Validées (YES)",
        "sign_header": "🖋️ Signature et Validation Légale",
        "sign_label": "Saisissez votre Nom et Prénom complet pour la signature électronique :",
        "attest_label": "Je vérifie que les réponses fournies ci-dessus sont exactes.",
        "submit_btn": "💾 Valider et Enregistrer la Matrice",
        "err_missing": "Action Bloquée : Vous devez répondre à tout le formulaire, justifier les NO/NA, inscrire votre nom et cocher l'attestation.",
        "success_log": "🎉 Matrice synchronisée avec succès ! Enregistrée par",
        "no_history": "Aucun enregistrement de matrice trouvé dans la base pour le moment.",
        "history_title": "🔎 Filtrer, chercher et gérer les matrices enregistrées",
        "search_user": "Chercher par nom d'utilisateur :",
        "search_cmo": "Chercher par No de Lot (CMO) :",
        "search_month": "Chercher par Mois :",
        "search_year": "Chercher par Année :",
        "all": "Tous",
        "dl_excel": "📥 Télécharger le rapport COMPLET (Excel)",
        "dl_pdf": "📥 Télécharger le rapport COMPLET (PDF)",
        "delete_confirm": "Matrice supprimée avec succès de l'historique.",
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

# Language Selector
selected_lang = st.sidebar.selectbox(LANG_DICT["English"]["sidebar_lang"], ["Français", "English"])
T = LANG_DICT[selected_lang]

# ==========================================
# 3. SIDEBAR: READABLE & DOWNLOADABLE REFERENCES
# ==========================================
st.sidebar.markdown("---")
st.sidebar.subheader(T["sidebar_ref_title"])
st.sidebar.caption(T["sidebar_ref_desc"])

if excel_exists:
    try:
        xls = pd.ExcelFile(master_excel_file)
        reference_sheets = ['Gestion des activités', 'Aperçu des normes', 'Matrice des meilleures pratique']
        
        for sheet in reference_sheets:
            if sheet in xls.sheet_names:
                ref_df = pd.read_excel(master_excel_file, sheet_name=sheet)
                
                # On-Screen Reading Expander Container
                with st.sidebar.expander(f"🔍 Read: {sheet}", expanded=False):
                    st.dataframe(ref_df, use_container_width=True)
                
                # Download Buffer Setup
                ref_buffer = io.BytesIO()
                with pd.ExcelWriter(ref_buffer, engine='openpyxl') as writer:
                    ref_df.to_excel(writer, index=False, sheet_name=sheet)
                
                # Actionable Download Button Trigger
                st.sidebar.download_button(
                    label=f"📥 Download: {sheet}",
                    data=ref_buffer.getvalue(),
                    file_name=f"Indigo_Reference_{sheet.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"sidebar_dl_{sheet}"
                )
                st.sidebar.markdown("---")
    except Exception as e:
        st.sidebar.error(f"Error parsing reference sheets: {e}")
else:
    st.sidebar.error(f"⚠️ Master list file '{master_excel_file}' not found.")

# ==========================================
# 4. MAIN LAYOUT HEADER WITH INTEGRATED LOGO
# ==========================================
col_logo, col_title = st.columns([1.2, 4])
with col_logo:
    st.image("https://i.ibb.co/DHgswzDq/indigo-park-canada-logo.jpg", use_container_width=True)
with col_title:
    st.title(T["title"])
    st.caption(T["subtitle"])

st.markdown("---")

# ==========================================
# 5. DYNAMIC MATRIX CMO RETRIEVAL ENGINE
# ==========================================
@st.cache_data
def extract_real_cmo_data():
    if excel_exists:
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

cmo_options = extract_real_cmo_data()
years_options = [2026, 2025, 2024]
months_options = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"] if selected_lang == "Français" else ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

tab1, tab2 = st.tabs([T["tab_new"], T["tab_history"]])

# ==========================================
# TAB 1: OPERATIONAL ENTRY PANEL
# ==========================================
with tab1:
    col_cmo, col_yr, col_mnth = st.columns(3)
    with col_cmo:
        selected_cmo = st.selectbox(T["select_cmo"], cmo_options)
    with col_yr:
        chosen_year = st.selectbox(T["select_year"], years_options, index=0)
    with col_mnth:
        chosen_month = st.selectbox(T["select_month"], months_options, index=6)

    st.subheader(f"{T['form_header']} : {selected_cmo} — ({chosen_month} {chosen_year})")
    st.write(T["form_instruction"])
    
    responses = {}
    task_comments = {}
    incomplete_selections_flag = False
    missing_comments_flag = False
    
    for idx, task in enumerate(LANG_DICT["English"]["tasks"], start=1):
        st.markdown(f"**{idx}. {T['tasks'][idx-1]}**")
        
        responses[f"task_{idx}"] = st.radio(
            f"Statut {idx}", ["YES", "NO", "N/A"], index=None, key=f"task_radio_{idx}", label_visibility="collapsed"
        )
        
        if responses[f"task_{idx}"] is None:
            incomplete_selections_flag = True
            task_comments[f"comment_{idx}"] = ""
        elif responses[f"task_{idx}"] in ["NO", "N/A"]:
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

    st.markdown("---")
    st.subheader(T["metrics_header"])
    
    answers = [v for v in responses.values() if v is not None]
    yes_count = answers.count("YES")
    no_count = answers.count("NO")
    na_count = answers.count("N/A")
    
    if not incomplete_selections_flag:
        applicable_count = len(LANG_DICT["English"]["tasks"]) - na_count
        base_score = (yes_count / applicable_count * 100) if applicable_count > 0 else 100.0
        
        is_capped = False
        if responses["task_8"] == "NO" and base_score > 85.0:
            base_score = 85.0
            is_capped = True
        score_display = f"{base_score:.1f}%"
        status_display = T["m_complete"]
    else:
        base_score = 0.0
        is_capped = False
        score_display = "--"
        status_display = T["m_incomplete"]
        applicable_count = 0
        
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric(label=T["m_status"], value=status_display)
    with col_m2:
        if is_capped:
            st.metric(label=T["m_score"], value=score_display, delta=f"- {T['m_capped']}", delta_color="inverse")
        else:
            st.metric(label=T["m_score"], value=score_display)
    with col_m3:
        st.metric(label=T["m_passed"], value=f"{yes_count} / {14 - na_count if not incomplete_selections_flag else 14}")

    st.markdown("---")
    st.subheader(T["sign_header"])
    typed_signature = st.text_input(T["sign_label"])
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
            for i in range(1, 15):
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
                
            st.success(f"{T['success_log']} : {typed_signature} !")
            st.button("Refresh / Actualiser")

# ==========================================
# TAB 2: AUDIT LOG HISTORY MANAGEMENT
# ==========================================
with tab2:
    st.subheader(T["history_title"])
    save_path = "data/submissions.csv"
    
    if not os.path.isfile(save_path) or os.path.getsize(save_path) == 0:
        st.info(T["no_history"])
    else:
        history_df = pd.read_csv(save_path)
        
        h_col1, h_col2, h_col3, h_col4 = st.columns(4)
        with h_col1:
            s_user = st.text_input(T["search_user"], value="", key="search_user_input")
        with h_col2:
            s_cmo = st.text_input(T["search_cmo"], value="", key="search_cmo_input")
        with h_col3:
            s_month = st.selectbox(T["search_month"], [T["all"]] + months_options, key="search_month_select")
        with h_col4:
            s_year = st.selectbox(T["search_year"], [T["all"]] + [str(y) for y in years_options], key="search_year_select")
            
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
            st.warning("No matrix records found matching criteria.")
        else:
            display_columns = ["Report_ID", "Timestamp", "CMO_ID", "User", "Month", "Year", "Final_Score", "Capped_Flag"]
            st.dataframe(filtered_df[display_columns], use_container_width=True)
            
            st.markdown("### 🗑️ Record Management")
            selected_report_to_delete = st.selectbox("Select Report ID to Delete / Purge:", ["-- Select --"] + filtered_df["Report_ID"].tolist())
            
            if st.button("❌ Permanent Delete Selected Report"):
                if selected_report_to_delete != "-- Select --":
                    updated_master_df = history_df[history_df["Report_ID"] != selected_report_to_delete]
                    updated_master_df.to_csv(save_path, index=False)
                    st.success(T["delete_confirm"])
                    st.rerun()

            st.markdown("---")
            dl_col1, dl_col2 = st.columns(2)
            
            with dl_col1:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Comprehensive Data Log')
                
                st.download_button(
                    label=T["dl_excel"],
                    data=excel_buffer.getvalue(),
                    file_name=f"Indigo_Full_Matrix_Dump.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with dl_col2:
                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                styles = getSampleStyleSheet()
                
                style_normal = ParagraphStyle(name='WrapText', parent=styles['Normal'], fontSize=8, leading=10)
                style_header = ParagraphStyle(name='HeaderStyle', parent=styles['Normal'], fontSize=8, leading=10, fontName='Helvetica-Bold', textColor=colors.whitesmoke)
                
                story = []
                story.append(Paragraph(f"<b>INDIGO PARK CANADA — CITY REPORTING COMPREHENSIVE LOG</b>", styles['Heading2']))
                story.append(Paragraph(f"Exported On: {datetime.date.today().strftime('%Y-%m-%d')} | Records Found: {len(filtered_df)}", styles['Normal']))
                story.append(Spacer(1, 15))
                
                for idx, row in filtered_df.iterrows():
                    story.append(Paragraph(f"<b>Matrix Record Profile: {row['CMO_ID']} — {row['Month']} {row['Year']}</b> (Score: {row['Final_Score']})", styles['Heading3']))
                    story.append(Paragraph(f"Filed By: {row['User']} | Date Stamp: {row['Timestamp']} | Reference ID: {row['Report_ID']}", styles['Normal']))
                    story.append(Spacer(1, 5))
                    
                    pdf_table_data = [[Paragraph("No.", style_header), Paragraph("City Reporting Matrix Mandate Target Indicator", style_header), Paragraph("Status", style_header), Paragraph("Justification/Comment Context", style_header)]]
                    
                    for i in range(1, 15):
                        q_txt = str(row[f"Q{i}_Task_Text"])
                        q_stat = str(row[f"Q{i}_Status"])
                        q_comm = str(row[f"Q{i}_Justification"]) if pd.notna(row[f"Q{i}_Justification"]) else ""
                        
                        pdf_table_data.append([
                            Paragraph(str(i), style_normal),
                            Paragraph(q_txt, style_normal),
                            Paragraph(q_stat, style_normal),
                            Paragraph(q_comm, style_normal)
                        ])
                    
                    row_table = Table(pdf_table_data, colWidths=[25, 275, 50, 200])
                    row_table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2D144B')),
                        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                        ('VALIGN', (0,0), (-1,-1), 'TOP'),
                        ('BOTTOMPADDING', (0,0), (-1,0), 6),
                        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#FCFCFC')),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
                    ]))
                    story.append(row_table)
                    story.append(Spacer(1, 20))
                
                doc.build(story)
                
                st.download_button(
                    label=T["dl_pdf"],
                    data=pdf_buffer.getvalue(),
                    file_name=f"Indigo_Comprehensive_Matrix_Report.pdf",
                    mime="application/pdf"
                )
