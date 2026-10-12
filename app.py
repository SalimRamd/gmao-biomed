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
                
                obs_appel = st.text_area("Observations lors de l'appel (panne déclarée)")
                obs_tech = st.text_area("Observations technique (travaux réalisés)")
                
                submit_int = st.form_submit_button("Enregistrer l'intervention")
                if submit_int:
                    new_entry = {
                        interv_id_col: str(equip_val),
                        "N° Interv.": int_num,
                        "Début": str(int_date),
                        "Cadre": int_cadre,
                        "Techn.": int_tech,
                        "Fournisseur": int_fourn,
                        "N° BC": int_bc,
                        "Observations lors de l'appel": obs_appel,
                        "Observations technique": obs_tech
                    }
                    st.session_state.df_interv = pd.concat([st.session_state.df_interv, pd.DataFrame([new_entry])], ignore_index=True)
                    st.success("Intervention enregistrée avec succès !")
                    st.rerun()

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
    st.sidebar.subheader("📥 Remplacer les bases")
    new_eq_file = st.sidebar.file_uploader("Fichier Équipements (.xlsx)", type=["xlsx", "xls"], key="adm_eq")
    if new_eq_file:
        df_eq_up = pd.read_excel(new_eq_file)
        df_eq_up.columns = [str(c).strip() for c in df_eq_up.columns]
        st.session_state.df_equip = df_eq_up
        st.sidebar.success("Équipements actualisés !")

    new_in_file = st.sidebar.file_uploader("Fichier Interventions (.xlsx / .xls)", type=["xlsx", "xls"], key="adm_in")
    if new_in_file:
        df_in_up = pd.read_excel(new_in_file)
        df_in_up.columns = [str(c).strip() for c in df_in_up.columns]
        st.session_state.df_interv = df_in_up
        st.sidebar.success("Interventions actualisées !")

    st.sidebar.markdown("---")
    st.sidebar.subheader("💾 Exporter")
    if st.session_state.df_interv is not None and not st.session_state.df_interv.empty:
        import io
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state.df_interv.to_excel(writer, index=False, sheet_name='Interventions')
        st.sidebar.download_button(
            label="📥 Télécharger interventions.xlsx",
            data=buffer.getvalue(),
            file_name="interventions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- CORPS PRINCIPAL ---
st.title("🏥 GMAO Biomédicale")

df_equip = st.session_state.df_equip
df_interv = st.session_state.df_interv

if df_equip is None:
    st.error(f"⚠️ Le fichier `{EQUIP_FILE_DEFAULT}` n'a pas été trouvé sur GitHub.")
else:
    auto_equip_col = find_equip_number_column(df_equip)
    interv_id_col = find_equip_number_column(df_interv) if df_interv is not None and not df_interv.empty else "N° Equipement"

    # Indicateurs généraux
    m1, m2, m3 = st.columns(3)
    m1.metric("Équipements au parc", len(df_equip))
    m2.metric("Total Interventions", len(df_interv) if df_interv is not None else 0)
    m3.metric("Clé de liaison", f"{auto_equip_col} ↔ {interv_id_col}")

    st.markdown("---")

    # LES 3 GRANDES SECTIONS PRINCIPALES
    main_tabs = st.tabs([
        "📋 Parc des Équipements", 
        "🛠️ Registre des Interventions", 
        "🔍 Recherche Rapide Équipement"
    ])

    # ==========================================
    # SECTION 1 : PARC DES ÉQUIPEMENTS
    # ==========================================
    with main_tabs[0]:
        st.subheader("Inventaire général du parc biomédical")
        st.caption("💡 Cliquez sur une ligne pour afficher la fiche de l'équipement et son historique.")
        
        query_eq = st.text_input("Filtrer les équipements (nom, marque, modèle, service...) :", key="eq_search")
        display_eq = df_equip
        if query_eq:
            display_eq = df_equip[df_equip.apply(lambda r: r.astype(str).str.contains(query_eq, case=False).any(), axis=1)]
            st.write(f"{len(display_eq)} équipement(s) trouvé(s)")

        selection_eq = st.dataframe(
            display_eq,
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row"
        )

        selected_rows_eq = selection_eq.selection.rows if hasattr(selection_eq, "selection") else []
        if len(selected_rows_eq) > 0:
            row_data = display_eq.iloc[selected_rows_eq[0]]
            equip_val = row_data[auto_equip_col]
            st.markdown("---")
            display_equip_view(row_data, equip_val, is_admin, auto_equip_col, interv_id_col)

    # ==========================================
    # SECTION 2 : REGISTRE GLOBAL DES INTERVENTIONS
    # ==========================================
    with main_tabs[1]:
        st.subheader("Historique global de toutes les interventions")
        
        if df_interv is not None and not df_interv.empty:
            df_interv_clean = format_interventions_table(df_interv)
            
            # Filtres intelligents
            col_f1, col_f2, col_f3 = st.columns(3)
            
            # Filtre par technicien
            tech_list = ["Tous"]
            if "Techn." in df_interv_clean.columns:
                tech_list += [str(t) for t in df_interv_clean["Techn."].dropna().unique() if str(t).strip() != ""]
            selected_tech = col_f1.selectbox("Filtrer par technicien :", tech_list)
            
            # Filtre par type/cadre
            cadre_list = ["Tous"]
            if "Cadre" in df_interv_clean.columns:
                cadre_list += [str(c) for c in df_interv_clean["Cadre"].dropna().unique() if str(c).strip() != ""]
            selected_cadre = col_f2.selectbox("Filtrer par type / cadre :", cadre_list)

            # Recherche plein texte
            search_interv = col_f3.text_input("Recherche mot-clé (panne, n° BC, obs...) :")

            # Application des filtres
            filtered_interv = df_interv_clean
            if selected_tech != "Tous":
                filtered_interv = filtered_interv[filtered_interv["Techn."].astype(str) == selected_tech]
            if selected_cadre != "Tous":
                filtered_interv = filtered_interv[filtered_interv["Cadre"].astype(str) == selected_cadre]
            if search_interv:
                filtered_interv = filtered_interv[filtered_interv.apply(lambda r: r.astype(str).str.contains(search_interv, case=False).any(), axis=1)]

            st.write(f"Affichage de **{len(filtered_interv)}** intervention(s)")

            # Tableau cliquable des interventions
            selection_int = st.dataframe(
                filtered_interv,
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row"
            )

            # Affichage du rapport complet si on clique sur une ligne d'intervention
            selected_rows_int = selection_int.selection.rows if hasattr(selection_int, "selection") else []
            if len(selected_rows_int) > 0:
                int_row = filtered_interv.iloc[selected_rows_int[0]]
                st.markdown("---")
                st.markdown(f"### 🛠️ Détail de l'intervention : **{int_row.get('N° Interv.', 'Sans N°')}**")
                
                c_det1, c_det2 = st.columns(2)
                c_det1.write(f"**N° Équipement :** {int_row.get(interv_id_col, 'N/A')}")
                c_det1.write(f"**Date début :** {int_row.get('Début', 'N/A')}")
                c_det1.write(f"**Technicien :** {int_row.get('Techn.', 'N/A')}")
                c_det1.write(f"**Cadre :** {int_row.get('Cadre', 'N/A')}")

                c_det2.write(f"**Fournisseur :** {int_row.get('Fournisseur', 'N/A')}")
                c_det2.write(f"**N° Commande / BC :** {int_row.get('N° BC', 'N/A')}")
                c_det2.write(f"**Date fin :** {int_row.get('Fin', 'N/A')}")

                st.info(f"**Observations lors de l'appel (Panne) :**\n\n{int_row.get('Observations lors de l\'appel', 'Aucune note')}")
                st.success(f"**Observations technique (Réparation) :**\n\n{int_row.get('Observations technique', 'Aucune note')}")

        else:
            st.info("Aucun registre d'interventions n'est présent sur GitHub. Déposez `TEST.XLS` ou `interventions.xlsx`.")

    # ==========================================
    # SECTION 3 : RECHERCHE RAPIDE DIRECTE
    # ==========================================
    with main_tabs[2]:
        st.subheader("Recherche directe par numéro d'équipement")
        equip_numbers = df_equip[auto_equip_col].dropna().astype(str).unique().tolist()
        chosen_id = st.selectbox("Sélectionnez ou tapez le numéro :", equip_numbers, key="quick_search_box")
        if chosen_id:
            row_data = df_equip[df_equip[auto_equip_col].astype(str) == chosen_id].iloc[0]
            st.markdown("---")
            display_equip_view(row_data, chosen_id, is_admin, auto_equip_col, interv_id_col)
