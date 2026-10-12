import streamlit as st
import pandas as pd
import os
from datetime import datetime, date

# Configuration de la page
st.set_page_config(page_title="GMAO Biomédicale", layout="wide", page_icon="🏥")

EQUIP_FILE_DEFAULT = "Equipement 2026.xlsx"
INTERV_CANDIDATES = ["interventions.xlsx", "TEST.XLS", "test.xls", "TEST.xlsx"]

# 1. Chargement des données
@st.cache_data
def load_data():
    df_eq = None
    df_in = None
    
    # Équipements
    if os.path.exists(EQUIP_FILE_DEFAULT):
        try:
            df_eq = pd.read_excel(EQUIP_FILE_DEFAULT)
            df_eq.columns = [str(c).strip() for c in df_eq.columns]
        except Exception as e:
            st.error(f"Erreur de lecture de {EQUIP_FILE_DEFAULT} : {e}")
            
    # Interventions
    for fname in INTERV_CANDIDATES:
        if os.path.exists(fname):
            try:
                df_in = pd.read_excel(fname)
                df_in.columns = [str(c).strip() for c in df_in.columns]
                break
            except Exception as e:
                st.error(f"Erreur de lecture de {fname} : {e}")
                
    return df_eq, df_in

default_equip, default_interv = load_data()

if "df_equip" not in st.session_state or st.session_state.df_equip is None:
    st.session_state.df_equip = default_equip
if "df_interv" not in st.session_state or st.session_state.df_interv is None:
    st.session_state.df_interv = default_interv

# Détection de la colonne identifiant N° Équipement
def find_equip_number_column(df):
    if df is None:
        return None
    priorities = [
        "n° equipement", "n° équipement", "numero equipement", "numéro équipement",
        "n° inventaire", "numero inventaire", "numéro d'inventaire", "n°inventaire",
        "code équipement", "code equipement", "equipement", "équipement",
        "n° serie", "n° série"
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

# Mise en ordre des colonnes d'intervention pour un affichage propre
def format_interventions_table(df):
    if df is None or df.empty:
        return df
    target_cols = [
        "N° Interv.", "N° Equipement", "Début", "Fin", "Cadre", "Techn.", "Fournisseur", 
        "Observations lors de l'appel", "Observations technique", "N° BC"
    ]
    existing = [c for c in target_cols if c in df.columns]
    if existing:
        others = [c for c in df.columns if c not in existing and c != "R.I. Simple"]
        return df[existing + others]
    return df

# Affichage détaillé d'un équipement
def display_equip_view(row_data, equip_val, is_admin, auto_equip_col, interv_id_col):
    st.markdown(f"### 📋 Fiche Équipement : **{equip_val}**")
    tabs_eq = ["📄 Fiche Signalétique", "🛠️ Historique des Interventions"]
    if is_admin:
        tabs_eq.append("➕ Ajouter une Intervention")
    
    eq_subtabs = st.tabs(tabs_eq)

    # Signalétique
    with eq_subtabs[0]:
        cols = st.columns(3)
        idx = 0
        for col_name, val in row_data.items():
            if pd.notna(val) and str(val).strip() != "":
                cols[idx % 3].write(f"**{col_name} :** {val}")
                idx += 1

    # Historique lié
    with eq_subtabs[1]:
        current_interv = st.session_state.df_interv
        if current_interv is not None and not current_interv.empty and interv_id_col in current_interv.columns:
            match = current_interv[current_interv[interv_id_col].astype(str).str.strip() == str(equip_val).strip()]
            if not match.empty:
                st.success(f"{len(match)} intervention(s) archivée(s) pour cet appareil.")
                cleaned_match = format_interventions_table(match)
                st.dataframe(cleaned_match, use_container_width=True)
            else:
                st.info("Aucune intervention enregistrée pour cet équipement.")
        else:
            st.warning("Le registre des interventions n'est pas encore chargé.")

    # Ajout manuel (Admin)
    if is_admin:
        with eq_subtabs[2]:
            st.subheader("Enregistrer une intervention sur cet appareil")
            with st.form(f"form_interv_{equip_val}", clear_on_submit=True):
                f_c1, f_c2 = st.columns(2)
                int_date = f_c1.date_input("Date début", value=date.today())
                int_cadre = f_c2.selectbox("Cadre / Type", ["Corrective", "Préventive", "Contrôle Qualité / RSAC", "Mise en service"])
                
                int_tech = f_c1.text_input("Technicien")
                int_fourn = f_c2.text_input("Prestataire / Fournisseur")
                
                int_num = f_c1.text_input("N° Intervention")
                int_bc = f_c2.text_input("N° Devis / Commande / BC")
                
                obs
