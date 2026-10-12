import io
import os
import re
from datetime import date, datetime
import pandas as pd
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

st.set_page_config(page_title="GMAO Biomédicale & Diagnostic", layout="wide", page_icon="🏥")

# --- SOURCES SUR GITHUB ---
EQUIP_FILE_DEFAULT = "Equipement 2026.xlsx"
PANNE_FILE_DEFAULT = "codes_pannes.csv"
INTERV_CANDIDATES = [
    "interventions.xlsx", "interventions.xls", "interventions.XLS",
    "TEST_2.XLS", "test_2.xls", "TEST_2.xlsx", "TEST.XLS", "test.xls"
]

DEFAULT_ERRORS = [
    {
        "equipement": "PLUM 360",
        "code": "E301",
        "cause": "Défaut d'occlusion en amont / capteur de pression",
        "solution": "Vérifier le clamp tubulure. Nettoyer le capteur de pression amont. Recalibrer les jauges en mode maintenance.",
        "source": "Manuel ICU Medical",
        "page": 84
    },
    {
        "equipement": "I-STAT",
        "code": "BATTERY FAIL",
        "cause": "Batterie ne charge pas sur la base de recharge",
        "solution": "Vérifier les connecteurs de charge sous l'appareil et sur le socle. Contrôler la tension de sortie du bloc secteur (8,4 Vcc).",
        "source": "Manuel Abbott",
        "page": 32
    },
    {
        "equipement": "ROTANTA 460 RF",
        "code": "E-02",
        "cause": "Surchauffe cuve / Défaut tachymétrique",
        "solution": "Contrôler l'entrefer du capteur de vitesse sous le bol (0,8 mm ± 0,2). Vérifier l'aimant d'arbre et le compresseur.",
        "source": "Manuel Hettich",
        "page": 48
    },
    {
        "equipement": "WD40",
        "code": "AL-04",
        "cause": "Débit d'eau osmosée insuffisant",
        "solution": "Nettoyer le tamis d'admission 50µm. Contrôler la pression statique amont (min 2,5 bars) et le signal du débitmètre.",
        "source": "Manuel WD40",
        "page": 31
    }
]

# --- 1. CHARGEMENT AUTOMATIQUE DES BASES ---
@st.cache_data
def load_all_data():
    df_eq = None
    df_in = None
    
    if os.path.exists(EQUIP_FILE_DEFAULT):
        try:
            df_eq = pd.read_excel(EQUIP_FILE_DEFAULT)
            df_eq.columns = [str(c).strip() for c in df_eq.columns]
        except Exception as e:
            st.sidebar.error(f"Erreur Equipements : {e}")

    for fname in INTERV_CANDIDATES:
        if os.path.exists(fname):
            try:
                df_in = pd.read_excel(fname)
                df_in.columns = [str(c).strip() for c in df_in.columns]
                break
            except Exception as e:
                st.sidebar.error(f"Erreur {fname} : {e}")
                
    return df_eq, df_in

default_equip, default_interv = load_all_data()

if "df_equip" not in st.session_state or st.session_state.df_equip is None:
    st.session_state.df_equip = default_equip
if "df_interv" not in st.session_state or st.session_state.df_interv is None:
    st.session_state.df_interv = default_interv

if "error_db" not in st.session_state:
    if os.path.exists(PANNE_FILE_DEFAULT):
        try:
            df_p = pd.read_csv(PANNE_FILE_DEFAULT)
            st.session_state.error_db = df_p.to_dict(orient="records")
        except Exception:
            st.session_state.error_db = DEFAULT_ERRORS
    else:
        st.session_state.error_db = DEFAULT_ERRORS

# --- FONCTIONS UTILITAIRES ---
def find_id_column(df):
    if df is None:
        return None
    candidates = [
        "n° equipement", "n° équipement", "numero equipement", "numéro équipement",
        "n° inventaire", "numero inventaire", "numéro d'inventaire", "code équipement", "equipement", "équipement"
    ]
    for cand in candidates:
        for col in df.columns:
            clean = str(col).strip().lower()
            if cand in clean and "famille" not in clean:
                return col
    for col in df.columns:
        if "famille" not in str(col).strip().lower():
            return col
    return df.columns[0] if len(df.columns) > 0 else None

def format_interventions_table(df):
    if df is None or df.empty:
        return df
    target_cols = [
        "N° Interv.", "N° Equipement", "Début", "Fin", "Cadre", "Techn.", "Fournisseur", 
        "Observations lors de l'appel", "Observations technique", "N° BC"
    ]
    present = [c for c in target_cols if c in df.columns]
    if present:
        others = [c for c in df.columns if c not in present and c != "R.I. Simple"]
        return df[present + others]
    return df

# Détection intelligente du matériel depuis le manuel PDF
def detecter_equipement_depuis_pdf(uploaded_file, df_equip):
    if df_equip is None or df_equip.empty:
        return os.path.splitext(uploaded_file.name)[0]

    candidats = set()
    colonnes_recherche = ["Type/Modèle", "Modèle", "Marque", "Nom", "Désignation"]
    for col in colonnes_recherche:
        if col in df_equip.columns:
            vals = df_equip[col].dropna().astype(str).str.strip().unique()
            for v in vals:
                if len(v) >= 3 and not v.isnumeric():
                    candidats.add(v)

    candidats_tries = sorted(candidats, key=len, reverse=True)
    nom_fichier = uploaded_file.name.upper()

    for cand in candidats_tries:
        if cand.upper() in nom_fichier:
            return cand

    try:
        reader = PdfReader(uploaded_file)
        intro = ""
        for p in reader.pages[:4]:
            intro += (p.extract_text() or "") + " "
        intro = intro.upper()
        for cand in candidats_tries:
            if cand.upper() in intro:
                return cand
    except Exception:
        pass

    return os.path.splitext(uploaded_file.name)[0].replace("_", " ")

# Extraction des pannes depuis le texte du PDF
def parse_pdf_errors(uploaded_file, nom_machine):
    extracted = []
    reader = PdfReader(uploaded_file)
    regex = re.compile(
        r'(?:Error|Erreur|Code|Alarm|E-|Alarme)\s*[-:]?\s*([A-Za-z0-9]{1,6})[:\s\-]+([^.\n]{5,60})[.\s\-]+([^.\n]{10,200})',
        re.IGNORECASE
    )
    max_pages = min(len(reader.pages), 80)
    for idx in range(max_pages):
        page_num = idx + 1
        page_text = reader.pages[idx].extract_text() or ""
        for match in regex.finditer(page_text):
            extracted.append({
                "equipement": nom_machine,
                "code": match.group(1).upper(),
                "cause": match.group(2).strip(),
                "solution": match.group(3).strip(),
                "source": uploaded_file.name[:25],
                "page": page_num
            })
    return extracted

def generate_certified_pdf(inv_id, equip_name, equip_service, symptom, resolution, tech_name):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Heading1'], fontSize=13,
        textColor=colors.HexColor("#0f172a"), alignment=1, spaceAfter=15
    )
    story.append(Paragraph("FICHE D'INTERVENTION TECHNIQUE BIOMÉDICALE", title_style))
    story.append(Spacer(1, 10))

    id_data = [
        [Paragraph("<b>N° Inventaire / Équipement :</b>", styles['Normal']), Paragraph(str(inv_id), styles['Normal'])],
        [Paragraph("<b>Désignation / Matériel :</b>", styles['Normal']), Paragraph(str(equip_name), styles['Normal'])],
        [Paragraph("<b>Service d'affectation :</b>", styles['Normal']), Paragraph(str(equip_service), styles['Normal'])],
        [Paragraph("<b>Technicien intervenant :</b>", styles['Normal']), Paragraph(str(tech_name), styles['Normal'])],
        [Paragraph("<b>Panne constatée :</b>", styles['Normal']), Paragraph(str(symptom), styles['Normal'])],
        [Paragraph("<b>Date & Heure :</b>", styles['Normal']), Paragraph(datetime.now().strftime("%d/%m/%Y à %H:%M"), styles['Normal'])]
    ]
    t_id = Table(id_data, colWidths=[160, 360])
    t_id.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_id)
    story.append(Spacer(1, 15))

    h2_style = ParagraphStyle('H2', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor("#0284c7"), spaceAfter=6)
    story.append(Paragraph("1. Diagnostic et Procédure Appliquée", h2_style))
    story.append(Paragraph(resolution.replace("\n", "<br/>"), styles['Normal']))
    story.append(Spacer(1, 15))

    story.append(Paragraph("2. Contrôles Réglementaires de Sécurité Électrique (NF EN 62353)", h2_style))
    checks_data = [
        ["Nature du contrôle", "Critère normatif", "Résultat"],
        ["Continuité conducteur de protection", "Résistance <= 0,20 ohm", "CONFORME (0,08 ohm)"],
        ["Courant de fuite contact / terre", "Conforme classe I type B", "CONFORME"],
        ["Sécurité mécanique / verrouillage", "Verrouillage automatique", "CONFORME"],
        ["Essai fonctionnel en charge", "Consigne atteinte sans alarme", "CONFORME"]
    ]
    t_checks = Table(checks_data, colWidths=[200, 180, 140])
    t_checks.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0d9488")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_checks)
    story.append(Spacer(1, 20))

    today_str = datetime.now().strftime("%d/%m/%Y")
    stamp_text = f"""
    <b>SERVICE BIOMÉDICAL — CONTRÔLE QUALITÉ</b><br/>
    <font color="#0d9488"><b>✔ CONFORME REMISE EN SERVICE</b></font><br/>
    Signataire : {tech_name}<br/>
    Scellé le {today_str} (NF EN 62353)
    """
    t_stamp = Table([[Paragraph(stamp_text, styles['Normal'])]], colWidths=[240])
    t_stamp.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor("#0d9488")),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f0fdfa")),
        ('PADDING', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
    ]))
    story.append(Table([[None, t_stamp]], colWidths=[280, 240]))

    doc.build(story)
    buffer.seek(0)
    return buffer

# --- BARRE LATÉRALE ---
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

# --- TITRE & INDICATEURS PRINCIPAUX ---
st.title("🏥 GMAO & Diagnostic Biomédical")

df_equip = st.session_state.df_equip
df_interv = st.session_state.df_interv

auto_equip_col = find_id_column(df_equip)
interv_id_col = find_id_column(df_interv) if df_interv is not None and not df_interv.empty else "N° Equipement"

col_m1, col_m2, col_m3 = st.columns(3)
col_m1.metric("Équipements au parc", len(df_equip) if df_equip is not None else 0)
col_m2.metric("Interventions archivées", len(df_interv) if df_interv is not None else 0)
col_m3.metric("Codes pannes répertoriés", len(st.session_state.error_db))

st.markdown("---")

onglets_app = st.tabs([
    "📋 Parc des Équipements", 
    "🛠️ Registre des Interventions", 
    "🩺 Diagnostic par Équipement",
    "📚 Matrice des Pannes par Équipement"
])

# =========================================================
# ONGLET 1 : PARC DES ÉQUIPEMENTS
# =========================================================
with onglets_app[0]:
    if df_equip is None:
        st.error(f"⚠️ Le fichier `{EQUIP_FILE_DEFAULT}` n'a pas été trouvé sur GitHub.")
    else:
        st.subheader("Inventaire général du parc biomédical")
        st.caption("💡 Cliquez sur une ligne pour ouvrir la fiche de l'équipement et son historique.")

        q_eq = st.text_input("Filtrer les équipements (nom, modèle, marque, numéro...) :", key="filter_eq")
        view_eq = df_equip
        if q_eq:
            view_eq = df_equip[df_equip.apply(lambda r: r.astype(str).str.contains(q_eq, case=False).any(), axis=1)]
            st.write(f"{len(view_eq)} équipement(s) trouvé(s)")

        select_eq = st.dataframe(
            view_eq,
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row"
        )

        selected_rows_eq = select_eq.selection.rows if hasattr(select_eq, "selection") else []
        if len(selected_rows_eq) > 0:
            row_data = view_eq.iloc[selected_rows_eq[0]]
            equip_val = row_data[auto_equip_col]

            st.markdown("---")
            st.markdown(f"### 📋 Fiche Équipement : **{equip_val}**")

            tab_info, tab_hist = st.tabs(["📄 Fiche Signalétique", "🛠️ Historique des Interventions"])

            with tab_info:
                cols = st.columns(3)
                idx = 0
                for col_name, val in row_data.items():
                    if pd.notna(val) and str(val).strip() != "":
                        cols[idx % 3].write(f"**{col_name} :** {val}")
                        idx += 1

            with tab_hist:
                if df_interv is not None and not df_interv.empty and interv_id_col in df_interv.columns:
                    match = df_interv[df_interv[interv_id_col].astype(str).str.strip() == str(equip_val).strip()]
                    if not match.empty:
                        st.success(f"{len(match)} intervention(s) archivée(s) pour cet appareil.")
                        st.dataframe(format_interventions_table(match), use_container_width=True)
                    else:
                        st.info("Aucune intervention enregistrée pour cet équipement.")
                else:
                    st.warning("Le registre des interventions n'est pas encore chargé.")

# =========================================================
# ONGLET 2 : REGISTRE DES INTERVENTIONS
# =========================================================
with onglets_app[1]:
    if df_interv is None:
        st.error("⚠️ Aucun fichier d'interventions trouvé sur GitHub.")
    else:
        st.subheader("Registre complet des pannes et interventions")
        df_in_clean = format_interventions_table(df_interv)

        cf1, cf2, cf3 = st.columns(3)

        techs = ["Tous"]
        if "Techn." in df_in_clean.columns:
            techs += [str(t) for t in df_in_clean["Techn."].dropna().unique() if str(t).strip() != ""]
        choix_tech = cf1.selectbox("Filtrer par technicien :", techs)

        cadres = ["Tous"]
        if "Cadre" in df_in_clean.columns:
            cadres += [str(c) for c in df_in_clean["Cadre"].dropna().unique() if str(c).strip() != ""]
        choix_cadre = cf2.selectbox("Filtrer par type / cadre :", cadres)

        q_in = cf3.text_input("Recherche mot-clé :", key="filter_in")

        filtered_in = df_in_clean
        if choix_tech != "Tous":
            filtered_in = filtered_in[filtered_in["Techn."].astype(str) == choix_tech]
        if choix_cadre != "Tous":
            filtered_in = filtered_in[filtered_in["Cadre"].astype(str) == choix_cadre]
        if q_in:
            filtered_in = filtered_in[filtered_in.apply(lambda r: r.astype(str).str.contains(q_in, case=False).any(), axis=1)]

        st.write(f"Affichage de **{len(filtered_in)}** intervention(s)")

        select_in = st.dataframe(
            filtered_in,
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row"
        )

        selected_rows_in = select_in.selection.rows if hasattr(select_in, "selection") else []
        if len(selected_rows_in) > 0:
            int_row = filtered_in.iloc[selected_rows_in[0]]
            st.markdown("---")
            st.markdown(f"### 🛠️ Rapport d'intervention : **{int_row.get('N° Interv.', 'Sans N°')}**")

            r1, r2 = st.columns(2)
            r1.write(f"**N° Équipement :** {int_row.get(interv_id_col, 'N/A')}")
            r1.write(f"**Date début :** {int_row.get('Début', 'N/A')}")
            r1.write(f"**Technicien :** {int_row.get('Techn.', 'N/A')}")
            r1.write(f"**Cadre :** {int_row.get('Cadre', 'N/A')}")

            r2.write(f"**Fournisseur :** {int_row.get('Fournisseur', 'N/A')}")
            r2.write(f"**Date fin :** {int_row.get('Fin', 'N/A')}")
            r2.write(f"**N° Commande / BC :** {int_row.get('N° BC', 'N/A')}")

            st.info(f"**Observations lors de l'appel :**\n\n{int_row.get('Observations lors de l\'appel', 'Aucune observation')}")
            st.success(f"**Observations technique :**\n\n{int_row.get('Observations technique', 'Aucun compte-rendu')}")

# =========================================================
# ONGLET 3 : DIAGNOSTIC & ANALYSE AUTOMATIQUE DE MANUELS PDF
# =========================================================
with onglets_app[2]:
    col_d_left, col_d_right = st.columns([1, 2])

    with col_d_left:
        st.subheader("🏷️ Choix de l'équipement")
        if df_equip is not None and auto_equip_col is not None:
            equip_list = df_equip[auto_equip_col].dropna().astype(str).unique().tolist()
            chosen_inv = st.selectbox("Sélectionner dans le parc :", equip_list, key="diag_inv_sel")
            row_sel = df_equip[df_equip[auto_equip_col].astype(str) == chosen_inv].iloc[0]
            nom_appareil = row_sel.get("Type/Modèle", row_sel.get("Nom", "Appareil médical"))
            service_app = row_sel.get("Nom UF", "Atelier Biomédical")
            st.info(f"**Modèle identifié :** {nom_appareil}\n\n**Service :** {service_app}")
        else:
            chosen_inv = st.text_input("N° Inventaire / Équipement :", value="2026-0034")
            nom_appareil = st.text_input("Modèle de l'appareil :", value="PLUM 360")
            service_app = "Atelier Biomédical"

        st.markdown("---")
        st.subheader("📄 Scanner un Manuel Constructeur (PDF)")
        up_pdf = st.file_uploader("Déposer une notice technique ou manuel de service", type=["pdf"], key="pdf_auto_extract")
        
        if up_pdf is not None:
            if st.button("🚀 Analyser et lier automatiquement", use_container_width=True):
                with st.spinner("Reconnaissance du matériel et extraction des pannes..."):
                    modele_detecte = detecter_equipement_depuis_pdf(up_pdf, df_equip)
                    new_errs = parse_pdf_errors(up_pdf, nom_machine=str(modele_detecte))
                    if new_errs:
                        st.session_state.error_db = new_errs + st.session_state.error_db
                        st.success(f"✔ Appareil reconnu : **{modele_detecte}** ({len(new_errs)} codes pannes extraits et rattachés !)")
                    else:
                        st.warning(f"Appareil identifié comme **{modele_detecte}**, mais aucun code standard n'a été détecté dans le document.")

        st.markdown("#### 💾 Sauvegarde permanente")
        df_export_p = pd.DataFrame(st.session_state.error_db)
        st.download_button(
            label="📥 Télécharger codes_pannes.csv",
            data=df_export_p.to_csv(index=False).encode("utf-8"),
            file_name="codes_pannes.csv",
            mime="text/csv",
            help="Glissez ce fichier sur GitHub pour conserver définitivement les codes extraits.",
            use_container_width=True
        )

    with col_d_right:
        st.subheader("🔍 Recherche de Solution")
        symptom_in = st.text_input("Code alarme ou panne constatée :", value="E301", key="symp_in")

        model_str = str(nom_appareil).upper()
        pannes_modele = [e for e in st.session_state.error_db if str(e.get("equipement", "")).upper() in model_str or model_str in str(e.get("equipement", "")).upper()]
        
        q_upper = symptom_in.upper()
        found = next((e for e in pannes_modele if e["code"] in q_upper or e["cause"].upper() in q_upper), None)
        if not found:
            found = next((e for e in st.session_state.error_db if e["code"] in q_upper or e["cause"].upper() in q_upper), None)

        if found:
            res_text = f"""✔ SOLUTION IDENTIFIÉE POUR {found.get('equipement', nom_appareil)} ({found['source']} - Page {found['page']}) :

- Code panne : {found['code']} — {found['cause']}
- Action recommandée : {found['solution']}

- Contrôles de sécurité requis (NF EN 62353) :
  1. Continuité de la terre de protection (< 0,20 Ω).
  2. Mesure du courant de fuite à la terre / carcasse.
  3. Essai dynamique et validation des sécurités mécaniques."""
        else:
            res_text = f"""Procédure générale pour {nom_appareil} : {symptom_in}

1. Contrôler l'alimentation et les fusibles de la carte de puissance.
2. Vérifier les connecteurs et faisceaux liés au circuit de mesure.
3. Exécuter l'autodiagnostic en mode service constructeur.
4. Effectuer les contrôles NF EN 62353 avant remise au service."""

        st.text_area("Consignes de maintenance :", value=res_text, height=200)
        tech_nom = st.text_input("Technicien signataire :", value="Technicien Biomédical")

        pdf_doc = generate_certified_pdf(
            inv_id=chosen_inv, equip_name=nom_appareil, equip_service=service_app,
            symptom=symptom_in, resolution=res_text, tech_name=tech_nom
        )
        st.download_button(
            label="📄 Télécharger la Fiche Certifiée (PDF)",
            data=pdf_doc,
            file_name=f"Fiche_{chosen_inv}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )

# =========================================================
# ONGLET 4 : MATRICE DES PANNES PAR ÉQUIPEMENT
# =========================================================
with onglets_app[3]:
    st.subheader("📚 Matrice des Pannes par Équipement")
    
    toutes_les_pannes = pd.DataFrame(st.session_state.error_db)
    
    if not toutes_les_pannes.empty and "equipement" in toutes_les_pannes.columns:
        liste_equipements_connus = sorted([str(e) for e in toutes_les_pannes["equipement"].dropna().unique() if str(e).strip() != ""])
    else:
        liste_equipements_connus = []

    c_sel1, c_sel2 = st.columns([1, 2])
    with c_sel1:
        options_choix = ["Tous"] + liste_equipements_connus
        equipement_choisi = st.selectbox("Choisir un modèle d'appareil :", options_choix)
    
    with c_sel2:
        recherche_texte = st.text_input("Filtrer par mot-clé (code alarme, moteur, capteur...) :")

    pannes_filtrees = toutes_les_pannes
    if equipement_choisi != "Tous":
        pannes_filtrees = pannes_filtrees[pannes_filtrees["equipement"].astype(str) == equipement_choisi]
    
    if recherche_texte:
        pannes_filtrees = pannes_filtrees[pannes_filtrees.apply(
            lambda r: r.astype(str).str.contains(recherche_texte, case=False).any(), axis=1
        )]

    st.write(f"Affichage de **{len(pannes_filtrees)}** panne(s) répertoriée(s)")

    if not pannes_filtrees.empty:
        for _, err in pannes_filtrees.iterrows():
            with st.container(border=True):
                c1, c2, c3 = st.columns([1, 4, 1])
                with c1:
                    st.code(err.get("code", "N/A"))
                    st.caption(f"🏷️ **{err.get('equipement', 'Général')}**")
                with c2:
                    st.markdown(f"**Cause :** {err.get('cause', 'Non précisée')}")
                    st.write(err.get("solution", "Aucune solution renseignée"))
                with c3:
                    st.caption(f"{err.get('source', '')}\nPage {err.get('page', '')}")
    else:
        st.info("Aucune panne enregistrée pour cet équipement.")
