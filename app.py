import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="GMAO Biomédicale", layout="wide", page_icon="🏥")

# --- INITIALISATION DE LA MÉMOIRE (SESSION STATE) ---
if "df_equip" not in st.session_state:
    st.session_state.df_equip = None
if "df_interv" not in st.session_state:
    st.session_state.df_interv = None

# --- FONCTION UTILITAIRE POUR DÉTECTER LA CLÉ DE LIAISON ---
def find_id_column(df, candidates=["inventaire", "n° inventaire", "n°inventaire", "id", "numéro d'inventaire", "serie", "n° serie", "n° série"]):
    for col in df.columns:
        clean_col = str(col).strip().lower()
        for cand in candidates:
            if cand in clean_col:
                return col
    return df.columns[0] if len(df.columns) > 0 else None

# --- BARRE LATÉRALE : GESTION DES PROFILS ---
st.sidebar.title("🔐 Espace Accès")

role = st.sidebar.radio("Profil d'accès :", ["Utilisateur (Consultation)", "Administrateur (Gestion)"])

is_admin = False
if role == "Administrateur (Gestion)":
    # Mot de passe par défaut : biomed2026 (tu peux le modifier ici)
    admin_password = st.sidebar.text_input("Mot de passe administrateur :", type="password")
    if admin_password == "biomed2026":
        is_admin = True
        st.sidebar.success("Mode Administrateur activé ✅")
    elif admin_password:
        st.sidebar.error("Mot de passe incorrect")

# --- ZONE ADMIN : CHARGEMENT DES FICHIERS EXCEL ---
if is_admin:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📥 Gestion des bases Excel")
    
    upload_equip = st.sidebar.file_uploader("Fichier Équipements (.xlsx)", type=["xlsx", "xls"], key="up_equip")
    if upload_equip is not None:
        df_eq = pd.read_excel(upload_equip)
        df_eq.columns = [str(c).strip() for c in df_eq.columns]
        st.session_state.df_equip = df_eq
        st.sidebar.success("Équipements chargés !")

    upload_interv = st.sidebar.file_uploader("Fichier Interventions (.xlsx)", type=["xlsx", "xls"], key="up_interv")
    if upload_interv is not None:
        df_in = pd.read_excel(upload_interv)
        df_in.columns = [str(c).strip() for c in df_in.columns]
        st.session_state.df_interv = df_in
        st.sidebar.success("Interventions chargées !")

st.title("🏥 GMAO Biomédicale")

# --- VÉRIFICATION DE PRÉSENCE DES DONNÉES ---
df_equip = st.session_state.df_equip
df_interv = st.session_state.df_interv

if df_equip is None:
    st.info("👋 Bienvenue sur la GMAO du service biomédical.")
    if is_admin:
        st.warning("⚠️ En tant qu'administrateur, veuillez déposer le fichier Excel des équipements dans le panneau à gauche.")
    else:
        st.warning("⚠️ La base de données n'est pas encore chargée. Veuillez contacter un administrateur du service.")
else:
    equip_id_col = find_id_column(df_equip)
    interv_id_col = find_id_column(df_interv) if df_interv is not None else None

    # Indicateurs clés
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Équipements au parc", len(df_equip))
    col_m2.metric("Interventions archivées", len(df_interv) if df_interv is not None else "Non chargées")
    col_m3.metric("Clé d'identification", equip_id_col)

    st.markdown("---")

    # Onglets de navigation principale
    list_tabs = ["🔍 Consulter un Équipement", "📋 Parc Global"]
    if is_admin:
        list_tabs.append("➕ Ajouter un Équipement")

    tabs = st.tabs(list_tabs)

    # --- ONGLET 1 : FICHE ÉQUIPEMENT & HISTORIQUE ---
    with tabs[0]:
        st.subheader("Sélection d'un dispositif médical")
        
        equip_ids = df_equip[equip_id_col].dropna().astype(str).unique().tolist()
        selected_id = st.selectbox("Rechercher par identifiant :", equip_ids)

        if selected_id:
            row = df_equip[df_equip[equip_id_col].astype(str) == selected_id].iloc[0]

            # Sous-onglets propres à l'équipement sélectionné
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
                    st.warning("Le registre des interventions n'a pas encore été importé.")

    # --- ONGLET 2 : VUE D'ENSEMBLE DU PARC ---
    with tabs[1]:
        st.subheader("Inventaire général du parc biomédical")
        query = st.text_input("Filtrer le tableau (marque, modèle, service...) :")
        if query:
            filtered = df_equip[df_equip.apply(lambda r: r.astype(str).str.contains(query, case=False).any(), axis=1)]
            st.write(f"{len(filtered)} résultat(s)")
            st.dataframe(filtered, use_container_width=True)
        else:
            st.dataframe(df_equip, use_container_width=True)

    # --- ONGLET 3 (ADMIN UNIQUEMENT) : AJOUT DE MATÉRIEL ---
    if is_admin:
        with tabs[2]:
            st.subheader("Ajouter un nouvel équipement au parc")
            st.write("Renseignez les champs ci-dessous pour intégrer une nouvelle fiche :")

            with st.form("add_equip_form", clear_on_submit=True):
                form_values = {}
                form_cols = st.columns(2)
                for i, col_name in enumerate(df_equip.columns):
                    form_values[col_name] = form_cols[i % 2].text_input(f"{col_name} :")
                
                submitted = st.form_submit_button("Enregistrer l'équipement")
                if submitted:
                    new_row = pd.DataFrame([form_values])
                    st.session_state.df_equip = pd.concat([st.session_state.df_equip, new_row], ignore_index=True)
                    st.success("Nouvel équipement ajouté avec succès à la base courante !")
                    st.rerun()

            st.markdown("---")
            st.subheader("Exporter l'inventaire mis à jour")
            # Préparation de l'export Excel
            import io
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                st.session_state.df_equip.to_excel(writer, index=False, sheet_name='Équipements')
            
            st.download_button(
                label="📥 Télécharger le fichier Excel mis à jour",
                data=buffer.getvalue(),
                file_name=f"inventaire_biomed_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
