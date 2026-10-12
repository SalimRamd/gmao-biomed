import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuration de la page
st.set_page_config(page_title="GMAO Biomédicale", layout="wide", page_icon="🏥")

EQUIP_FILE_DEFAULT = "Equipement 2026.xlsx"
INTERV_FILE_DEFAULT = "interventions.xlsx"

# 1. Chargement des données
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

default_equip, default_interv = load_data()

if "df_equip" not in st.session_state or st.session_state.df_equip is None:
    st.session_state.df_equip = default_equip
if "df_interv" not in st.session_state or st.session_state.df_interv is None:
    st.session_state.df_interv = default_interv

# Détection automatique de la colonne d'identifiant d'équipement
def find_equip_number_column(df):
    if df is None:
        return None
    priorities = [
        "n° equipement", "n° équipement", "numero equipement", "numéro équipement",
        "n° inventaire", "numero inventaire", "numéro d'inventaire", "n°inventaire",
        "code équipement", "code equipement", "equipement", "équipement",
        "n° serie", "n° série", "numero serie", "numéro de série"
    ]
    for prio in priorities:
        for col in df.columns:
            clean_col = str(col).strip().lower()
            if prio in clean_col and "famille" not in clean_col:
                return col
    for col in df.columns:
        if "famille" not in str(col).strip().lower():
            return col
    return df.columns[0] if len(df.columns) > 0 else None

# Fonction d'affichage réutilisable pour la fiche détaillée
def display_equip_details(row_data, equip_val, df_interv, interv_id_col):
    st.markdown(f"### 📋 Détails de l'équipement : **{equip_val}**")
    sub_info, sub_hist = st.tabs(["📄 Fiche Signalétique", "🛠️ Historique des Interventions"])
    
    with sub_info:
        cols = st.columns(3)
        idx = 0
        for col_name, val in row_data.items():
            if pd.notna(val):
                cols[idx % 3].write(f"**{col_name} :** {val}")
                idx += 1
                
    with sub_hist:
        if df_interv is not None and interv_id_col is not None:
            match = df_interv[df_interv[interv_id_col].astype(str) == str(equip_val)]
            if not match.empty:
                st.success(f"{len(match)} intervention(s) répertoriée(s) pour cet équipement.")
                st.dataframe(match, use_container_width=True)
            else:
                st.info("Aucune intervention enregistrée pour cet équipement.")
        else:
            st.warning("Le registre des interventions n'est pas encore importé.")

# --- MENU LATÉRAL ---
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

if is_admin:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📥 Mettre à jour les fichiers")
    new_equip = st.sidebar.file_uploader("Remplacer Équipements (.xlsx)", type=["xlsx", "xls"], key="admin_up_eq")
    if new_equip:
        df_temp = pd.read_excel(new_equip)
        df_temp.columns = [str(c).strip() for c in df_temp.columns]
        st.session_state.df_equip = df_temp
        st.sidebar.success("Équipements actualisés !")

    new_interv = st.sidebar.file_uploader("Charger Interventions (.xlsx)", type=["xlsx", "xls"], key="admin_up_in")
    if new_interv:
        df_itemp = pd.read_excel(new_interv)
        df_itemp.columns = [str(c).strip() for c in df_itemp.columns]
        st.session_state.df_interv = df_itemp
        st.sidebar.success("Interventions actualisées !")

# --- APPLICATION PRINCIPALE ---
st.title("🏥 GMAO Biomédicale")

df_equip = st.session_state.df_equip
df_interv = st.session_state.df_interv

if df_equip is None:
    st.error(f"⚠️ Le fichier `{EQUIP_FILE_DEFAULT}` n'a pas été trouvé sur GitHub.")
    st.info("Déposez le fichier sur GitHub ou importez-le en mode administrateur.")
else:
    auto_equip_col = find_equip_number_column(df_equip)
    interv_id_col = find_equip_number_column(df_interv) if df_interv is not None else None

    # Indicateurs
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Équipements au parc", len(df_equip))
    col_m2.metric("Interventions archivées", len(df_interv) if df_interv is not None else "En attente")
    col_m3.metric("Clé d'identification", auto_equip_col)

    st.markdown("---")

    list_tabs = ["📋 Parc Global (Sélection cliquable)", "🔍 Recherche Rapide"]
    if is_admin:
        list_tabs.append("➕ Ajouter un Équipement")

    tabs = st.tabs(list_tabs)

    # 1. PARC GLOBAL AVEC CLIC INTERACTIF
    with tabs[0]:
        st.subheader("Inventaire général du parc biomédical")
        st.caption("💡 Cliquez sur une ligne du tableau pour afficher automatiquement la fiche et les interventions.")
        
        query = st.text_input("Filtrer le tableau (mot-clé, marque, service...) :", key="filter_global")
        display_df = df_equip
        if query:
            display_df = df_equip[df_equip.apply(lambda r: r.astype(str).str.contains(query, case=False).any(), axis=1)]
            st.write(f"{len(display_df)} résultat(s)")

        # Tableau interactif avec sélection de ligne
        selection = st.dataframe(
            display_df,
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row"
        )

        # Si l'utilisateur clique sur une ligne
        selected_rows = selection.selection.rows if hasattr(selection, "selection") else []
        if len(selected_rows) > 0:
            selected_index = selected_rows[0]
            selected_row_data = display_df.iloc[selected_index]
            equip_val = selected_row_data[auto_equip_col]
            st.markdown("---")
            display_equip_details(selected_row_data, equip_val, df_interv, interv_id_col)

    # 2. RECHERCHE RAPIDE PAR MENU
    with tabs[1]:
        st.subheader("Recherche ciblée par numéro d'équipement")
        col_select, col_search = st.columns([1, 2])
        with col_select:
            col_list = list(df_equip.columns)
            default_index = col_list.index(auto_equip_col) if auto_equip_col in col_list else 0
            equip_id_col = st.selectbox("Champ de recherche :", col_list, index=default_index, key="search_col")

        with col_search:
            equip_numbers = df_equip[equip_id_col].dropna().astype(str).unique().tolist()
            selected_num = st.selectbox(f"Sélectionner / Saisir le {equip_id_col} :", equip_numbers, key="search_val")

        if selected_num:
            row_match = df_equip[df_equip[equip_id_col].astype(str) == selected_num].iloc[0]
            st.markdown("---")
            display_equip_details(row_match, selected_num, df_interv, interv_id_col)

    # 3. AJOUT D'ÉQUIPEMENT (ADMIN)
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
