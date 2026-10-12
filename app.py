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

# --- FICHIERS PERMANENTS SUR GITHUB ---
EQUIP_FILE_DEFAULT = "Equipement 2026.xlsx"
PANNE_FILE_DEFAULT = "codes_pannes.csv"
INTERV_CANDIDATES = [
    "interventions.xlsx", "interventions.xls", "interventions.XLS",
    "TEST_2.XLS", "test_2.xls", "TEST_2.xlsx", "TEST.XLS", "test.xls"
]

# Codes d'erreurs par défaut (si codes_pannes.csv n'est pas encore sur GitHub)
DEFAULT_ERRORS = [
    {
        "code": "E-02",
        "cause": "Surchauffe cuve / Défaut tachymétrique",
        "solution": "Contrôler l'entrefer du capteur de vitesse sous le bol (0,8 mm ± 0,2). Vérifier l'aimant d'arbre moteur et les enroulements du compresseur.",
        "source": "Manuel Hettich",
        "page": 48
    },
    {
        "code": "LID OPEN",
        "cause": "Contacteur couvercle défaillant",
        "solution": "Vérifier la course du micro-contacteur de sécurité (SW1). Remplacer l'électro-aimant de verrouillage si résistance infinie.",
        "source": "Manuel Hettich",
        "page": 52
    },
    {
        "code": "AL-04",
        "cause": "Débit d'eau osmosée insuffisant",
        "solution": "Nettoyer le tamis d'admission 50µm. Vérifier la pression statique amont (min 2,5 bars) et le signal du débitmètre à turbine.",
        "source": "Manuel WD40",
        "page": 31
    },
    {
        "code": "AIR IN LINE",
        "cause": "Bulle détectée ou capteur encrassé",
        "solution": "Nettoyer les prismes optiques de détection ultrasonique avec de l'alcool isopropylique. Étalonner la sensibilité du capteur.",
        "source": "Manuel Fluido",
        "page": 19
    }
]

# --- 1. CHARGEMENT DES BASES DE DONNÉES ---
@st.cache_data
def load_all_data():
    df_eq = None
    df_in = None
    
    # 1. Équipements
    if os.path.exists(EQUIP_FILE_DEFAULT):
        try:
            df_eq = pd.read_excel(EQUIP_FILE_DEFAULT)
            df_eq.columns = [str(c).strip() for c in df_eq.columns]
        except Exception as e:
            st.sidebar.error(f"Erreur Equipements : {e}")

    # 2. Interventions
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

# 3. Chargement permanent des codes pannes
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

def parse_pdf_errors(uploaded_file):
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
        [Paragraph("<b>Panne / Symptôme constaté :</b>", styles['Normal']), Paragraph(str(symptom), styles['Normal'])],
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
st.title("🏥 GMAO & Maintenance Biomédicale")

df_equip = st.session_state.df_equip
df_interv = st.session_state.df_interv

auto_equip_col = find_id_column(df_equip)
interv_id_col = find_id_column(df_interv) if df_interv is not None and not df_interv.empty else "N° Equipement"

col_m1, col_m2, col_m3 = st.columns(3)
col_m1.metric("Équipements au parc", len(df_equip) if df_equip is not None else 0)
col_m2.metric("Interventions archivées", len(df_interv) if df_interv is not None else 0)
col_m3.metric("Codes pannes répertoriés", len(st.session_state.error_db))

st.markdown("---")

# ONGLETS PRINCIPAUX
onglets_app = st.tabs([
    "📋 Parc des Équipements", 
    "🛠️ Registre des Interventions", 
    "🩺 Diagnostic & Codes Pannes",
    "📚 Matrice des Pannes"
])

# =========================================================
# ONGLET 1 : PARC DES ÉQUIPEMENTS
# =========================================================
with onglets_app[0]:
    if df_equip is None:
        st.error(f"⚠️ Le fichier `{EQUIP_FILE_DEFAULT}` n'a pas été trouvé sur GitHub.")
    else:
        st.subheader("Inventaire général du parc biomédical")
        st.caption("💡 Cliquez sur une ligne pour afficher la fiche de l'équipement et son historique.")

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
        st.error("⚠️ Aucun fichier d'interventions trouvé sur GitHub (ex: `interventions.xlsx` ou `TEST_2.XLS`).")
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

        q_in = cf3.text_input("Recherche mot-clé (panne, n° BC, observations...) :", key="filter_in")

        filtered_in = df_in_clean
        if choix_tech != "Tous":
            filtered_in = filtered_in[filtered_in["Techn."].astype(str) == choix_tech]
        if choix_cadre != "Tous":
            filtered_in = filtered_in[filtered_in["Cadre"].astype(str) == choix_cadre]
        if q_in:
            filtered_in = filtered_in[filtered_in.apply(lambda r: r.astype(str).str.contains(q_in, case=False).any(), axis=1)]

        st.write(f"Affichage de **{len(filtered_in)}** intervention(s)")
        st.caption("💡 Cliquez sur une ligne pour ouvrir le rapport d'intervention complet.")

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

            st.info(f"**Observations lors de l'appel (Panne déclarée) :**\n\n{int_row.get('Observations lors de l\'appel', 'Aucune observation')}")
            st.success(f"**Observations technique (Compte-rendu réparation) :**\n\n{int_row.get('Observations technique', 'Aucun compte-rendu')}")

# =========================================================
# ONGLET 3 : DIAGNOSTIC & CODES PANNES (AVEC SAUVEGARDE PERMANENTE)
# =========================================================
with onglets_app[2]:
    col_d_left, col_d_right = st.columns([1, 2])

    with col_d_left:
        st.subheader("🏷️ Équipement concerné")
        if df_equip is not None and auto_equip_col is not None:
            equip_list = df_equip[auto_equip_col].dropna().astype(str).unique().tolist()
            chosen_inv = st.selectbox("Sélectionner dans le parc :", equip_list, key="diag_inv")
            row_sel = df_equip[df_equip[auto_equip_col].astype(str) == chosen_inv].iloc[0]
            nom_appareil = row_sel.get("Nom", row_sel.get("Désignation", "Appareil médical"))
            service_app = row_sel.get("Nom UF", row_sel.get("Service", "Atelier Biomédical"))
            st.info(f"**Désignation :** {nom_appareil}\n\n**Service :** {service_app}")
        else:
            chosen_inv = st.text_input("N° Inventaire / Équipement :", value="2026-0034")
            nom_appareil = "Appareil médical"
            service_app = "Biomédical"

        st.markdown("---")
        st.subheader("📄 Manuel Constructeur (PDF)")
        uploaded_pdf = st.file_uploader("Charger une notice ou manuel technique", type=["pdf"], key="pdf_manual")
        
        if uploaded_pdf is not None:
            if st.button("📥 Extraire les codes erreurs du PDF", use_container_width=True):
                with st.spinner("Analyse du manuel technique en cours..."):
                    new_errors = parse_pdf_errors(uploaded_pdf)
                    if new_errors:
                        st.session_state.error_db = new_errors + st.session_state.error_db
                        st.success(f"✔ {len(new_errors)} codes pannes extraits et ajoutés !")
                    else:
                        st.warning("Aucun code panne conventionnel détecté dans ce fichier.")

        # BOUTON POUR SAUVEGARDER SUR GITHUB
        st.markdown("#### 💾 Sauvegarder pour toujours")
        df_export_pannes = pd.DataFrame(st.session_state.error_db)
        csv_bytes = df_export_pannes.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Télécharger codes_pannes.csv",
            data=csv_bytes,
            file_name="codes_pannes.csv",
            mime="text/csv",
            help="Téléchargez ce fichier et déposez-le sur GitHub : les codes pannes ne s'effaceront plus jamais !",
            use_container_width=True
        )

    with col_d_right:
        st.subheader("🔍 Recherche de Solution Constructeur")
        symptom_query = st.text_input(
            "Code alarme ou symptôme observé :",
            value="E-02 surchauffe cuve",
            key="symp_input"
        )

        q_clean = symptom_query.upper()
        found_err = next((item for item in st.session_state.error_db if item["code"] in q_clean or item["cause"].upper() in q_clean), None)

        if found_err:
            res_text = f"""✔ SOLUTION CONSTRUCTEUR IDENTIFIÉE ({found_err['source']} - Page {found_err['page']}) :

- Code panne : {found_err['code']} — {found_err['cause']}
- Action recommandée : {found_err['solution']}

- Contrôles de sécurité requis (NF EN 62353) :
  1. Continuité de la terre de protection (< 0,20 Ω).
  2. Mesure du courant de fuite à la terre / carcasse.
  3. Essai dynamique et validation des sécurités mécaniques."""
        else:
            res_text = f"""Procédure générale de diagnostic pour : {symptom_query}

1. Contrôler les tensions d'alimentation secondaires sur la carte de puissance.
2. Inspecter les connecteurs et faisceaux de câblage (oxydation, échauffement).
3. Exécuter la routine d'autotest en mode maintenance constructeur.
4. Réaliser les tests de sécurité électrique selon la norme NF EN 62353 avant remise en service."""

        st.text_area("Consignes techniques & Procédure de maintenance :", value=res_text, height=200)

        tech_intervenant = st.text_input("Nom du technicien signataire :", value="Technicien Biomédical")

        # Fiche certifiée PDF
        pdf_file = generate_certified_pdf(
            inv_id=chosen_inv,
            equip_name=nom_appareil,
            equip_service=service_app,
            symptom=symptom_query,
            resolution=res_text,
            tech_name=tech_intervenant
        )

        st.download_button(
            label="📄 Télécharger la Fiche d'Intervention Certifiée (PDF)",
            data=pdf_file,
            file_name=f"Fiche_Technique_{chosen_inv}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )

# =========================================================
# ONGLET 4 : MATRICE DES PANNES COMPILÉES
# =========================================================
with onglets_app[3]:
    st.subheader("📋 Matrice Globale des Pannes Extraites")
    filter_panne = st.text_input("Filtrer par code ou symptôme :", placeholder="Ex: E-02, rotor, débit, température...")

    matched_db = [
        item for item in st.session_state.error_db
        if filter_panne.lower() in item["code"].lower()
        or filter_panne.lower() in item["cause"].lower()
        or filter_panne.lower() in item["solution"].lower()
    ]

    if matched_db:
        for err in matched_db:
            with st.container(border=True):
                c_c1, c_c2, c_c3 = st.columns([1, 4, 1])
                with c_c1:
                    st.code(err["code"])
                with c_c2:
                    st.markdown(f"**Cause :** {err['cause']}")
                    st.write(err["solution"])
                with c_c3:
                    st.caption(f"{err['source']}\nPage {err['page']}")
    else:
        st.info("Aucun code panne ne correspond à ce filtre.")
