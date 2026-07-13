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
    div.stRadio > div{ flex-direction:row; } /* Horizonal Radio buttons */
    </style>
""", unsafe_allow_html=True)

# App Header with Official Logo Layout
col_logo, col_title = st.columns([1, 4])
with col_logo:
    try:
        # Tries to read the logo you uploaded
        st.image(Image.open("indigo_logo_0.png"), width=180)
    except:
        st.write("🅿️ **INDIGO PARK**")
with col_title:
    st.title("Operations & Compliance Portal")
    st.caption("Montreal Region Lot Management & Best Practices Digital Tracker")

st.markdown("---")

# ==========================================
# 2. DYNAMIC DATA ENGINE (READS RAW EXCEL)
# ==========================================
@st.cache_data
def extract_real_cmo_data():
    try:
        # Parses the 'Montreal' sheet inside your raw uploaded spreadsheet
        df = pd.read_excel('Montreal Lot List.xlsx', sheet_name='Montreal', skiprows=1)
        df.columns = df.columns.str.strip()
        cmo_list = df['Lot #'].dropna().astype(str).str.strip().unique().tolist()
        # Grabs everything starting with CMO or VMO cleanly
        cmo_list = [x for x in cmo_list if x.startswith(('CMO', 'VMO'))]
        return sorted(cmo_list)
    except Exception as e:
        # Safe fallback if file access encounters directory shifting
        return [f"CMO{str(i).zfill(3)}" for i in [2, 4, 8, 9, 10, 15, 20, 25, 37, 101, 102]]

cmo_options = extract_real_cmo_data()

# ==========================================
# 3. SIDEBAR HISTORIC FILTERS
# ==========================================
st.sidebar.header("🔍 Search & Filter Database")
filter_user = st.sidebar.text_input("Filter by User Name", value="", placeholder="e.g. Francis")
filter_year = st.sidebar.selectbox("Select Year", [2026, 2025, 2024])
filter_month = st.sidebar.selectbox(
    "Select Month", 
    ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
)

# ==========================================
# 4. MAIN AUDIT FORM & MATRIX QUESTIONS
# ==========================================
st.header("📝 Grille d'évaluation des meilleures pratiques")
selected_cmo = st.selectbox("Sélectionner l'emplacement cible (CMO / Lot ID) :", cmo_options)

st.subheader(f"📋 Formulaire de validation : {selected_cmo} — ({filter_month} {filter_year})")
st.write("Remplissez le statut pour l'ensemble des 14 indicateurs obligatoires :")

# The 14 official row elements extracted directly from your matrix worksheet
indigo_tasks = [
    "Rapport opérationnel complet du site",
    "Visite du site par le responsable/superviseur du site - interne",
    "Visite du site par le responsable/superviseur du site - externe",
    "Appel du gestionnaire de compte au client",
    "Visite du gestionnaire de compte sur le site/examen interne",
    "Visite du gestionnaire de compte sur le site/examen externe",
    "Gestionnaire de compte - Contact avec l'équipe du site",
    "Gestionnaire de compte - Réunion avec le client (Face-to-Face Meeting)", # Crucial Task #8
    "Appel du directeur général au client",
    "Le directeur général contacte le client en personne",
    "Visite du site par le directeur général / examen interne et externe",
    "Le vice-président des opérations contacte le client en personne",
    "Comptes rendus et marketing des événements spéciaux/sportifs",
    "Proposer des discussions/propositions axées sur la valeur ajoutée (Génération de revenus, Recommandations tarifaires)"
]

# Render the matrix rows with dynamic radio components
responses = {}
for idx, task in enumerate(indigo_tasks, start=1):
    st.markdown(f"**{idx}. {task}**")
    responses[f"task_{idx}"] = st.radio(
        f"Statut {idx}", ["YES", "NO", "N/A"], key=f"task_radio_{idx}", label_visibility="collapsed"
    )
    st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

st.markdown("---")
st.subheader("💬 Remarques et Justifications")
comments = st.text_area("Notes et explications complémentaires (Obligatoire si des éléments sont cochés 'NO' ou 'N/A') :")

# ==========================================
# 5. BUSINESS LOGIC & SCORE CAPPING ENGINE
# ==========================================
st.markdown("---")
st.subheader("📊 Résultats Métriques de Performance")

answers = list(responses.values())
yes_count = answers.count("YES")
no_count = answers.count("NO")
na_count = answers.count("N/A")

# Calculate metrics mathematically based only on applicable items
applicable_count = len(indigo_tasks) - na_count
if applicable_count > 0:
    base_score = (yes_count / applicable_count) * 100
else:
    base_score = 100.0

# Critical Gatekeeper Rule: Capping Performance at 85% maximum if they failed to meet the client
is_capped = False
if responses["task_8"] == "NO" and base_score > 85.0:
    base_score = 85.0
    is_capped = True

# Display Metrics Block
col_m1, col_m2, col_m3 = st.columns(3)
with col_m1:
    st.metric(label="Statut de Saisie", value="Formulaire Complété" if (yes_count + no_count + na_count == 14) else "Saisie Incomplète")
with col_m2:
    if is_capped:
        st.metric(label="Indice de Performance Réel", value=f"{base_score:.1f}%", delta="- Penalized Max Capping (No Client Meeting)", delta_color="inverse")
    else:
        st.metric(label="Indice de Performance Réel", value=f"{base_score:.1f}%")
with col_m3:
    st.metric(label="Tâches Validées (YES)", value=f"{yes_count} / {applicable_count}")

# ==========================================
# 6. E-SIGNATURE SYSTEM & FILE STORAGE SYNC
# ==========================================
st.markdown("---")
st.subheader("🖋️ Signature et Validation Légale")

typed_signature = st.text_input("Veuillez saisir votre Nom et Prénom complet pour validation de signature électronique :")
attestation_check = st.checkbox("J'atteste sur l'honneur avoir complété cette évaluation conformément aux réalités du terrain observées.")

if st.button("💾 Valider et Enregistrer dans la Base Historique"):
    if typed_signature.strip() == "" or not attestation_check:
        st.error("Action Bloquée : Vous devez saisir votre signature nominative et cocher la case d'attestation pour soumettre.")
    else:
        # Package data row payload
        payload = {
            "Timestamp": [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            "Year": [filter_year],
            "Month": [filter_month],
            "User": [typed_signature],
            "CMO_ID": [selected_cmo],
            "Final_Score": [f"{base_score:.1f}%"],
            "Capped_Flag": ["YES" if is_capped else "NO"],
            "Comments": [comments]
        }
        
        # Add individual answers for complete diagnostic auditing tracking rows
        for i in range(1, 15):
            payload[f"Task_{i}_Status"] = [responses[f"task_{i}"]]
            
        new_row_df = pd.DataFrame(payload)
        
        # Dynamic storage execution path
        save_path = "data/submissions.csv"
        os.makedirs("data", exist_ok=True)
        
        if not os.path.isfile(save_path):
            new_row_df.to_csv(save_path, index=False)
        else:
            new_row_df.to_csv(save_path, mode='a', header=False, index=False)
            
        st.success(f"🎉 Rapport synchronisé avec succès ! Enregistré de manière sécurisée sous l'identifiant de signature : {typed_signature}.")
