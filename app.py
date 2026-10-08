import streamlit as st
import pandas as pd

# Configuration de la page
st.set_page_config(page_title="GMAO Biomédicale", layout="wide", page_icon="🏥")

st.title("🏥 GMAO Biomédicale - Gestion de Parc & Interventions")

# --- BARRE LATÉRALE : CHARGEMENT DES FICHIERS EXCEL ---
st.sidebar.header("📁 Données Excel")
st.sidebar.write("Importez vos fichiers d'inventaire et d'interventions.")

equip_file = st.sidebar.file_uploader("1. Fichier Équipements (.xlsx)", type=["xlsx", "xls"])
interv_file = st.sidebar.file_uploader("2. Fichier Interventions (.xlsx)", type=["xlsx", "xls"])

# Fonction utilitaire pour trouver une colonne d'identifiant commun
def find_id_column(df, candidates=["inventaire", "n° inventaire", "n°inventaire", "id", "numéro d'inventaire", "serie", "n° serie", "n° série"]):
    for col in df.columns:
        clean_col = str(col).strip().lower()
        for cand in candidates:
            if cand in clean_col:
                return col
    return df.columns[0] if len(df.columns) > 0 else None

# Vérification du chargement des équipements
if equip_file is not None:
    df_equip = pd.read_excel(equip_file)
    # Nettoyage des noms de colonnes
    df_equip.columns = [str(c).strip() for c in df_equip.columns]
    equip_id_col = find_id_column(df_equip)
    
    # Chargement éventuel des interventions
    df_interv = None
    interv_id_col = None
    if interv_file is not None:
        df_interv = pd.read_excel(interv_file)
        df_interv.columns = [str(c).strip() for c in df_interv.columns]
        interv_id_col = find_id_column(df_interv)

    # --- MÉTRIQUES EN HAUT DE PAGE ---
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Équipements", len(df_equip))
    if df_interv is not None:
        col2.metric("Total Interventions", len(df_interv))
    else:
        col2.metric("Total Interventions", "Fichier non chargé")
    col3.metric("Clé de liaison détectée", f"{equip_id_col}")

    st.markdown("---")

    # --- ONGLETS PRINCIPAUX ---
    tab_search, tab_table, tab_all_interv = st.tabs([
        "🔍 Recherche & Fiche Équipement", 
        "📋 Liste Complète des Équipements", 
        "🛠️ Historique Global des Interventions"
    ])

    # 1. Onglet Fiche Équipement
    with tab_search:
        st.subheader("Consulter un équipement et ses interventions")
        
        # Liste de recherche des équipements
        equip_list = df_equip[equip_id_col].dropna().astype(str).unique().tolist()
        selected_equip = st.selectbox("Sélectionnez ou recherchez un équipement par son identifiant :", equip_list)
        
        if selected_equip:
            # Récupération de la ligne équipement
            equip_data = df_equip[df_equip[equip_id_col].astype(str) == selected_equip].iloc[0]
            
            st.markdown("### 📋 Fiche Signalétique")
            # Affichage en 3 colonnes pour une vue aérée
            info_cols = st.columns(3)
            col_idx = 0
            for col_name, val in equip_data.items():
                if pd.notna(val):
                    info_cols[col_idx % 3].write(f"**{col_name} :** {val}")
                    col_idx += 1
            
            st.markdown("---")
            st.markdown("### 🛠️ Interventions associées")
            
            if df_interv is not None and interv_id_col is not None:
                # Filtrage des interventions pour cet équipement précis
                matched_interv = df_interv[df_interv[interv_id_col].astype(str) == selected_equip]
                
                if not matched_interv.empty:
                    st.success(f"{len(matched_interv)} intervention(s) trouvée(s) pour cet équipement.")
                    st.dataframe(matched_interv, use_container_width=True)
                else:
                    st.info("Aucune intervention enregistrée pour cet équipement dans le fichier importé.")
            else:
                st.warning("⚠️ Pour voir les pannes et maintenances, déposez le fichier des interventions dans le panneau latéral gauche.")

    # 2. Onglet Tableau Général Équipements
    with tab_table:
        st.subheader("Parc biomédical complet")
        # Champ de recherche plein texte
        search_kw = st.text_input("Filtrer par mot-clé (désignation, marque, service...) :")
        if search_kw:
            filtered_df = df_equip[df_equip.apply(lambda row: row.astype(str).str.contains(search_kw, case=False).any(), axis=1)]
            st.write(f"{len(filtered_df)} résultat(s) trouvé(s)")
            st.dataframe(filtered_df, use_container_width=True)
        else:
            st.dataframe(df_equip, use_container_width=True)

    # 3. Onglet Toutes les Interventions
    with tab_all_interv:
        st.subheader("Historique global des interventions")
        if df_interv is not None:
            st.dataframe(df_interv, use_container_width=True)
        else:
            st.info("Déposez le fichier Excel d'interventions dans la barre latérale pour afficher l'historique global.")

else:
    # Message d'accueil quand aucun fichier n'est encore chargé
    st.info("👋 Pour commencer, veuillez déposer votre fichier Excel d'équipements dans le menu à gauche.")
    st.markdown("""
    **Format recommandé :**
    - Un fichier Excel pour les **Équipements** (avec une colonne type *N° inventaire* ou *N° série*).
    - Un fichier Excel pour les **Interventions** (avec cette même colonne pour relier la panne à l'appareil).
    """)
