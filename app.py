import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuration de la page
st.set_page_config(page_title="GMAO Biomédicale", layout="wide", page_icon="🏥")

# Noms de tes fichiers sur GitHub
EQUIP_FILE_DEFAULT = "Equipement 2026.xlsx"
INTERV_FILE_DEFAULT = "interventions.xlsx"

# 1. Chargement automatique des fichiers présents sur GitHub
@st.cache_data
def load_data():
    df_eq = None
    df_in = None
    if os.path.exists(EQUIP_FILE_DEFAULT):
        df_eq = pd.read_excel(EQUIP_FILE_DEFAULT)
        df_eq.columns = [str(c).strip() for c in df_eq.columns]
    if os.path.exists(INTERV_FILE_DEFAULT):
        df_in = pd.read_excel(INTERV_FILE_DEFAULT)
        df_in.columns = [str(c).strip() for c in df_in.columns]
    return df_eq, df_in

# Initialisation des données
default_equip, default_interv = load_data()

if "df_equip" not in st.session_state or st.session_state.df_equip is None:
    st.session_state.df_equip = default_equip
if "df_interv" not in st.session_state or st.session_state.df_interv is None:
    st.session_state.df_interv = default_interv

# Détection automatique de la colonne d'identifiant (N° inventaire ou série)
def find_id_column(df, candidates=["inventaire", "n° inventaire", "n°inventaire", "id", "numéro d'inventaire", "serie", "n° serie", "n° série", "code"]):
    if df is None:
        return None
    for col in df.columns:
        clean_col = str(col).strip().lower()
        for cand in candidates:
            if cand in clean_col:
                return col
    return df.columns[0] if len(df.columns) > 0 else None

# --- MENU LATÉRAL : ACCÈS ET PROFILS ---
st.sidebar.title("🔐 Espace Accès")
role = st.sidebar.radio("Profil d'accès :", ["Utilisateur (Consultation)", "Administrateur (Gestion)"])

is_admin = False
if role == "Administrateur (Gestion)":
    admin_password = st.sidebar.text_input("Mot de passe administrateur :", type="password")
    if admin_password == "biomed2026":
        is_admin = True
        st.sidebar.success("Mode Administrateur activé ✅")
    elif admin_password:
        st.sidebar.error("Mot de passe incorrect")

# Options réservées à l'administrateur
if is_admin:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📥 Mettre à jour les fichiers")
    new_equip = st.sidebar.file_uploader("Remplacer le fichier Équipements (.xlsx)", type=["xlsx", "xls"], key="admin_up_eq")
    if new_equip:
        df_temp = pd.read_excel(new_equip)
        df_temp.columns = [str(c).strip() for c in df_temp.columns]
        st.session_state.df_equip = df_temp
        st.sidebar.success("Équipements actualisés !")

    new_interv = st.sidebar.file_uploader("Charger le fichier Interventions (.xlsx)", type=["xlsx", "xls"], key="admin_up_in")
    if new_interv:
        df_itemp = pd.read_excel(new_interv)
        df_itemp.columns = [str(c).strip() for c in df_itemp.columns]
        st.session_state.df_interv = df_itemp
        st.sidebar.success("Interventions actualisées !")

# --- TITRE ET CONTENU ---
st.title("🏥 GMAO Biomédicale")

df_equip = st.session_state.df_equip
df_interv = st.session_state.df_interv

if df_equip is None:
    st.error(f"⚠️ Le fichier `{EQUIP_FILE_DEFAULT}` n'a pas été trouvé sur GitHub.")
    st.info("Vérifie bien que le fichier est déposé à la racine de ton projet GitHub avec l'orthographe exacte.")
else:
    equip_id_col = find_id_column(df_equip)
    interv_id_col = find_id_column(df_interv) if df_interv is not None else None

    # Indicateurs en haut
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Équipements au parc", len(df_equip))
    col_m2.metric("Interventions archivées", len(df_interv) if df_interv is not None else "En attente")
    col_m3.metric("Clé d'identification", equip_id_col)

    st.markdown("---")

    list_tabs = ["🔍 Consulter un Équipement", "📋 Parc Global"]
    if is_admin:
        list_tabs.append("➕ Ajouter un Équipement")

    tabs = st.tabs(list_tabs)

    # 1. Fiche équipement et historique des pannes
    with tabs[0]:
        st.subheader("Consultation d'un équipement")
        equip_ids = df_equip[equip_id_col].dropna().astype(str).unique().tolist()
        selected_id = st.selectbox("Rechercher par identifiant :", equip_ids)

        if selected_id:
            row = df_equip[df_equip[equip_id_col].astype(str) == selected_id].iloc[0]

            sub_tab_info, sub_tab_history = st.tabs(["📄 Fiche Signalétique", "🛠️ Historique des Interventions"])

            with sub_tab_info:
                st.markdown(f"#### Équipement : **{selected_id}**")
                cols = st.columns(3)
                idx = 0
                for col_name, val in row.items():
                    if pd.notna(val):
                        cols[idx % 3].write(f"**{col_name} :** {val}")
                        idx += 1

            with sub_tab_history:
                if df_interv is not None and interv_id_col is not None:
                    interv_match = df_interv[df_interv[interv_id_col].astype(str) == selected_id]
                    if not interv_match.empty:
                        st.success(f"{len(interv_match)} intervention(s) répertoriée(s) pour cet appareil.")
                        st.dataframe(interv_match, use_container_width=True)
                    else:
                        st.info("Aucune intervention enregistrée pour cet équipement.")
                else:
                    st.warning("Le registre des interventions n'est pas encore importé.")

    # 2. Vue du parc complet
    with tabs[1]:
        st.subheader("Inventaire général du parc")
        query = st.text_input("Filtrer par mot-clé (désignation, marque, service...) :")
        if query:
            filtered = df_equip[df_equip.apply(lambda r: r.astype(str).str.contains(query, case=False).any(), axis=1)]
            st.write(f"{len(filtered)} résultat(s)")
            st.dataframe(filtered, use_container_width=True)
        else:
            st.dataframe(df_equip, use_container_width=True)

    # 3. Ajout d'équipement (Admin)
    if is_admin:
        with tabs[2]:
            st.subheader("Ajouter un équipement")
            with st.form("add_equip_form", clear_on_submit=True):
                form_values = {}
                form_cols = st.columns(2)
                for i, col_name in enumerate(df_equip.columns):
                    form_values[col_name] = form_cols[i % 2].text_input(f"{col_name} :")
                
                submitted = st.form_submit_button("Enregistrer")
                if submitted:
                    new_row = pd.DataFrame([form_values])
                    st.session_state.df_equip = pd.concat([st.session_state.df_equip, new_row], ignore_index=True)
                    st.success("Équipement ajouté !")
                    st.rerun()
