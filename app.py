import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="GMAO Biomédicale", layout="wide", page_icon="🏥")

# --- NOMS DES FICHIERS PRÉSENTS SUR GITHUB ---
FICHIER_EQUIPEMENT = "Equipement 2026.xlsx"
FICHIER_INTERVENTION = "interventions.xlsx"
FICHIER_INTERVENTION_SECOURS = "interventions.xls"

# --- 1. CHARGEMENT AUTOMATIQUE DES DEUX FICHIERS ---
@st.cache_data
def charger_les_donnees():
    df_eq = None
    df_in = None

    # Chargement du fichier des équipements
    if os.path.exists(FICHIER_EQUIPEMENT):
        try:
            df_eq = pd.read_excel(FICHIER_EQUIPEMENT)
            df_eq.columns = [str(c).strip() for c in df_eq.columns]
        except Exception as e:
            st.error(f"Erreur de lecture de {FICHIER_EQUIPEMENT} : {e}")

    # Chargement du fichier des interventions
    source_interv = None
    if os.path.exists(FICHIER_INTERVENTION):
        source_interv = FICHIER_INTERVENTION
    elif os.path.exists(FICHIER_INTERVENTION_SECOURS):
        source_interv = FICHIER_INTERVENTION_SECOURS

    if source_interv:
        try:
            df_in = pd.read_excel(source_interv)
            df_in.columns = [str(c).strip() for c in df_in.columns]
        except Exception as e:
            st.error(f"Erreur de lecture de {source_interv} : {e}")

    return df_eq, df_in

# Mise en mémoire
df_equip, df_interv = charger_les_donnees()

# Détection de la colonne N° Equipement
def trouver_colonne_id(df):
    if df is None:
        return None
    candidats = ["n° equipement", "n° équipement", "numero equipement", "numéro équipement", "n° inventaire", "code équipement"]
    for c in candidats:
        for col in df.columns:
            if c in str(col).lower() and "famille" not in str(col).lower():
                return col
    return df.columns[0] if len(df.columns) > 0 else None

col_id_equip = trouver_colonne_id(df_equip)
col_id_interv = trouver_colonne_id(df_interv) if df_interv is not None else "N° Equipement"

# Nettoyage des colonnes d'interventions pour un affichage lisible
def colonnes_utiles_interventions(df):
    if df is None or df.empty:
        return df
    colonnes_prioritaires = [
        "N° Interv.", "N° Equipement", "Début", "Fin", "Cadre", 
        "Techn.", "Fournisseur", "Observations lors de l'appel", 
        "Observations technique", "N° BC"
    ]
    presentes = [c for c in colonnes_prioritaires if c in df.columns]
    autres = [c for c in df.columns if c not in presentes and c != "R.I. Simple"]
    return df[presentes + autres] if presentes else df

# --- TITRE PRINCIPAL ---
st.title("🏥 GMAO Biomédicale")

# Compteurs
c1, c2, c3 = st.columns(3)
c1.metric("Équipements chargés", len(df_equip) if df_equip is not None else 0)
c2.metric("Interventions chargées", len(df_interv) if df_interv is not None else 0)
c3.metric("Liaison", f"{col_id_equip} ↔ {col_id_interv}")

st.markdown("---")

# --- LES DEUX SECTIONS ---
onglet_equipements, onglet_interventions = st.tabs([
    "📋 Parc des Équipements", 
    "🛠️ Registre des Interventions"
])

# =========================================================
# SECTION 1 : TOUS LES ÉQUIPEMENTS
# =========================================================
with onglet_equipements:
    if df_equip is None:
        st.error(f"Fichier `{FICHIER_EQUIPEMENT}` non trouvé sur GitHub.")
    else:
        st.subheader("Inventaire général du parc")
        st.caption("💡 Cliquez sur une ligne pour voir sa fiche et ses interventions passées.")

        recherche_eq = st.text_input("Rechercher un équipement :", key="rech_eq")
        vue_eq = df_equip
        if recherche_eq:
            vue_eq = df_equip[df_equip.apply(lambda r: r.astype(str).str.contains(recherche_eq, case=False).any(), axis=1)]

        table_eq = st.dataframe(
            vue_eq,
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row"
        )

        lignes_selectionnees = table_eq.selection.rows if hasattr(table_eq, "selection") else []
        if len(lignes_selectionnees) > 0:
            ligne_eq = vue_eq.iloc[lignes_selectionnees[0]]
            valeur_num = ligne_eq[col_id_equip]

            st.markdown("---")
            st.markdown(f"### 📋 Détail de l'équipement : **{valeur_num}**")
            
            sous_tab_info, sous_tab_pannes = st.tabs(["📄 Informations", "🛠️ Pannes de cet appareil"])
            
            with sous_tab_info:
                cols = st.columns(3)
                idx = 0
                for nom_col, val in ligne_eq.items():
                    if pd.notna(val) and str(val).strip() != "":
                        cols[idx % 3].write(f"**{nom_col} :** {val}")
                        idx += 1

            with sous_tab_pannes:
                if df_interv is not None and not df_interv.empty and col_id_interv in df_interv.columns:
                    pannes_trouvees = df_interv[df_interv[col_id_interv].astype(str).str.strip() == str(valeur_num).strip()]
                    if not pannes_trouvees.empty:
                        st.success(f"{len(pannes_trouvees)} intervention(s) trouvée(s) pour cet appareil.")
                        st.dataframe(colonnes_utiles_interventions(pannes_trouvees), use_container_width=True)
                    else:
                        st.info("Aucune intervention enregistrée pour cet équipement.")
                else:
                    st.info("Aucun fichier d'interventions chargé.")

# =========================================================
# SECTION 2 : TOUTES LES INTERVENTIONS (TEST_2)
# =========================================================
with onglet_interventions:
    if df_interv is None:
        st.error(f"Fichier `{FICHIER_INTERVENTION}` non trouvé sur GitHub.")
    else:
        st.subheader("Registre complet de toutes les interventions")
        st.caption("💡 Cliquez sur une intervention pour afficher le compte-rendu complet.")

        df_in_propre = colonnes_utiles_interventions(df_interv)

        col_t, col_c, col_r = st.columns(3)
        liste_techs = ["Tous"]
        if "Techn." in df_in_propre.columns:
            liste_techs += [str(t) for t in df_in_propre["Techn."].dropna().unique() if str(t).strip() != ""]
        tech_choisi = col_t.selectbox("Filtrer par technicien :", liste_techs)

        liste_cadres = ["Tous"]
        if "Cadre" in df_in_propre.columns:
            liste_cadres += [str(c) for c in df_in_propre["Cadre"].dropna().unique() if str(c).strip() != ""]
        cadre_choisi = col_c.selectbox("Filtrer par type (Cadre) :", liste_cadres)

        texte_recherche = col_r.text_input("Rechercher dans les remarques/pannes :")

        vue_interv = df_in_propre
        if tech_choisi != "Tous":
            vue_interv = vue_interv[vue_interv["Techn."].astype(str) == tech_choisi]
        if cadre_choisi != "Tous":
            vue_interv = vue_interv[vue_interv["Cadre"].astype(str) == cadre_choisi]
        if texte_recherche:
            vue_interv = vue_interv[vue_interv.apply(lambda r: r.astype(str).str.contains(texte_recherche, case=False).any(), axis=1)]

        table_in = st.dataframe(
            vue_interv,
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row"
        )

        lignes_sel_in = table_in.selection.rows if hasattr(table_in, "selection") else []
        if len(lignes_sel_in) > 0:
            ligne_panne = vue_interv.iloc[lignes_sel_in[0]]
            st.markdown("---")
            st.markdown(f"### 🛠️ Rapport complet : Interv. n° **{ligne_panne.get('N° Interv.', 'N/A')}**")

            r1, r2 = st.columns(2)
            r1.write(f"**Appareil concerné :** {ligne_panne.get(col_id_interv, 'N/A')}")
            r1.write(f"**Date début :** {ligne_panne.get('Début', 'N/A')}")
            r1.write(f"**Technicien :** {ligne_panne.get('Techn.', 'N/A')}")
            r1.write(f"**Cadre :** {ligne_panne.get('Cadre', 'N/A')}")

            r2.write(f"**Fournisseur :** {ligne_panne.get('Fournisseur', 'N/A')}")
            r2.write(f"**Date fin :** {ligne_panne.get('Fin', 'N/A')}")
            r2.write(f"**N° Commande / BC :** {ligne_panne.get('N° BC', 'N/A')}")

            st.info(f"**Motif de l'appel (Panne) :**\n\n{ligne_panne.get('Observations lors de l\'appel', 'Aucune observation')}")
            st.success(f"**Travaux réalisés (Compte-rendu) :**\n\n{ligne_panne.get('Observations technique', 'Aucun compte-rendu')}")
