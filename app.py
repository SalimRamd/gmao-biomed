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

# --- 1. DOSSIERS & FICHIERS PERMANENTS SUR DISQUE ---
DOSSIER_MANUELS = "manuels"
os.makedirs(DOSSIER_MANUELS, exist_ok=True)

PANNE_FILE_DEFAULT = "codes_pannes.csv"

EQUIP_CANDIDATES = [
    "Equipements.xlsx", "Equipements.xls",
    "equipements.xlsx", "equipements.xls",
    "Équipements.xlsx", "Équipements.xls",
    "équipements.xlsx", "équipements.xls",
    "Equipement 2026.xlsx", "Les Equipements.xlsx"
]

INTERV_CANDIDATES = [
    "Interventions.xlsx", "Interventions.xls",
    "interventions.xlsx", "interventions.xls",
    "INTERVENTIONS.xlsx", "INTERVENTIONS.xls",
    "TEST_2.XLS", "TEST.XLS", "Les Interventions.xlsx"
]

# Codes de secours initiaux si le fichier CSV n'existe pas encore
DEFAULT_ERRORS = [
    {
        "marque": "ICU MEDICAL",
        "equipement": "PLUM 360",
        "code": "E301",
        "cause": "Défaut d'occlusion en amont / capteur de pression",
        "solution": "Vérifier le clamp tubulure. Nettoyer le capteur de pression amont. Recalibrer les jauges en mode maintenance.",
        "source": "Manuel ICU Medical",
        "page": 84
    },
    {
        "marque": "ABBOTT",
        "equipement": "I-STAT",
        "code": "BATTERY FAIL",
        "cause": "Batterie ne charge pas sur le socle",
        "solution": "Contrôler les connecteurs de charge sous l'appareil et sur le socle. Vérifier la tension (8,4 Vcc).",
        "source": "Manuel Abbott",
        "page": 32
    },
    {
        "marque": "HETTICH",
        "equipement": "ROTANTA 460 RF",
        "code": "E-02",
        "cause": "Surchauffe cuve / Défaut tachymétrique",
        "solution": "Contrôler l'entrefer du capteur de vitesse (0,8 mm ± 0,2). Vérifier l'aimant d'arbre et le compresseur.",
        "source": "Manuel Hettich",
        "page": 48
    },
    {
        "marque": "WASSENBURG",
        "equipement": "WD40",
        "code": "AL-04",
        "cause": "Débit d'eau osmosée insuffisant",
        "solution": "Nettoyer le tamis d'admission 50µm. Contrôler la pression statique amont (min 2,5 bars) et le débitmètre.",
        "source": "Manuel WD40",
        "page": 31
    }
]

# --- 2. CHARGEMENT DES BASES EXCEL ---
@st.cache_data
def load_all_data():
    df_eq, df_in = None, None
    for fname in EQUIP_CANDIDATES:
        if os.path.exists(fname):
            try:
                df_eq = pd.read_excel(fname)
                df_eq.columns = [str(c).strip() for c in df_eq.columns]
                break
            except Exception as e:
                st.sidebar.error(f"Erreur lecture {fname} : {e}")

    for fname in INTERV_CANDIDATES:
        if os.path.exists(fname):
            try:
                df_in = pd.read_excel(fname)
                df_in.columns = [str(c).strip() for c in df_in.columns]
                break
            except Exception as e:
                st.sidebar.error(f"Erreur lecture {fname} : {e}")
                
    return df_eq, df_in

default_equip, default_interv = load_all_data()

if "df_equip" not in st.session_state or st.session_state.df_equip is None:
    st.session_state.df_equip = default_equip
if "df_interv" not in st.session_state or st.session_state.df_interv is None:
    st.session_state.df_interv = default_interv


# --- 3. PERSISTANCE ANTI-F5 (ÉCRITURE/LECTURE DISQUE DIRECTE) ---
def charger_ou_initialiser_pannes():
    """Lit le fichier CSV s'il existe, sinon l'initialise et l'enregistre sur le disque."""
    if os.path.exists(PANNE_FILE_DEFAULT):
        try:
            df_p = pd.read_csv(PANNE_FILE_DEFAULT)
            if "marque" not in df_p.columns:
                df_p["marque"] = "Général"
            return df_p.to_dict(orient="records")
        except Exception:
            pass
    # Création du fichier sur le disque au 1er lancement
    df_init = pd.DataFrame(DEFAULT_ERRORS)
    df_init.to_csv(PANNE_FILE_DEFAULT, index=False, encoding="utf-8")
    return DEFAULT_ERRORS

# On charge les pannes depuis le disque dans la session state
if "error_db" not in st.session_state:
    st.session_state.error_db = charger_ou_initialiser_pannes()

def sauvegarder_pannes_sur_disque(nouvelles_pannes_list):
    """Fusionne les nouvelles pannes, dédoublonne, et écrit sur le disque immédiatement."""
    # Met à jour la mémoire vive
    st.session_state.error_db = nouvelles_pannes_list + st.session_state.error_db
    # Écrit tout sur le disque de façon définitive
    df_sauvegarde = pd.DataFrame(st.session_state.error_db)
    df_sauvegarde = df_sauvegarde.drop_duplicates(subset=["equipement", "code"], keep="first")
    df_sauvegarde.to_csv(PANNE_FILE_DEFAULT, index=False, encoding="utf-8")


# --- 4. IDENTIFICATION DES COLONNES ---
def find_id_column(df):
    if df is None:
        return None
    candidates = [
        "n° equipement", "n° équipement", "numero equipement", "numéro équipement",
        "n° inventaire", "numero inventaire", "numéro inventaire", "code equipement", 
        "code équipement", "equipement", "équipement"
    ]
    for cand in candidates:
        for col in df.columns:
            clean = str(col).strip().lower()
            if cand in clean and "famille" not in clean:
                return col
    return df.columns[0] if len(df.columns) > 0 else None

def find_name_column(df):
    if df is None:
        return None
    for cand in ["type/modèle", "modèle", "modele", "nom", "désignation", "designation"]:
        for col in df.columns:
            if cand in str(col).strip().lower():
                return col
    return df.columns[1] if len(df.columns) > 1 else None

def find_brand_column(df):
    if df is None:
        return None
    for cand in ["marque", "constructeur", "fabricant"]:
        for col in df.columns:
            if cand in str(col).strip().lower():
                return col
    return None

def find_uf_column(df):
    if df is None:
        return None
    for cand in ["nom uf", "code uf", "uf", "service", "affectation", "localisation", "pôle", "pole"]:
        for col in df.columns:
            clean = str(col).strip().lower()
            if cand == clean or cand in clean:
                return col
    return None

def find_sn_column(df):
    if df is None:
        return None
    for cand in ["n° serie", "n° série", "numero serie", "numéro de série", "sn", "s/n", "serie", "série"]:
        for col in df.columns:
            clean = str(col).strip().lower()
            if cand == clean or cand in clean:
                return col
    return None

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


# --- 5. ANALYSE ET DÉTECTION PDF ---
def detecter_marque_et_modele_pdf(uploaded_file, df_equip):
    texte_intro = ""
    try:
        reader = PdfReader(uploaded_file)
        for p in reader.pages[:4]:
            texte_intro += (p.extract_text() or "") + " "
    except Exception:
        pass
    
    texte_complet = (uploaded_file.name + " " + texte_intro).upper()
    
    marque_trouvee = "Autre Marque"
    marques_existantes = set()
    col_marque = find_brand_column(df_equip)
    if df_equip is not None and col_marque and col_marque in df_equip.columns:
        marques_existantes.update(df_equip[col_marque].dropna().astype(str).str.strip().unique())
    
    for err in st.session_state.error_db:
        if err.get("marque"):
            marques_existantes.add(err["marque"])
            
    marques_existantes.update(["ICU MEDICAL", "HETTICH", "ABBOTT", "WASSENBURG", "THE SURGICAL CO", "DRAEGER", "PHILIPS", "GE HEALTHCARE"])
    marques_triees = sorted([m for m in marques_existantes if len(m) >= 2], key=len, reverse=True)
    
    for m in marques_triees:
        if m.upper() in texte_complet:
            marque_trouvee = m
            break

    modele_trouve = os.path.splitext(uploaded_file.name)[0].replace("_", " ").replace("-", " ")
    col_nom = find_name_column(df_equip)
    modeles_existants = set()
    if df_equip is not None and col_nom and col_nom in df_equip.columns:
        for val in df_equip[col_nom].dropna().astype(str).str.strip().unique():
            if len(val) >= 3 and not val.isnumeric():
                modeles_existants.add(val)

    modeles_tries = sorted(modeles_existants, key=len, reverse=True)
    for mod in modeles_tries:
        if mod.upper() in texte_complet:
            modele_trouve = mod
            break

    return marque_trouvee, modele_trouve

def parse_pdf_errors(uploaded_file, marque_machine, nom_machine):
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
                "marque": str(marque_machine).strip().upper(),
                "equipement": str(nom_machine).strip(),
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
        [Paragraph("<b>Identification / N° :</b>", styles['Normal']), Paragraph(str(inv_id), styles['Normal'])],
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

# --- 6. HEADER DE L'APPLICATION ---
st.title("🏥 GMAO & Diagnostic Biomédical")

df_equip = st.session_state.df_equip
df_interv = st.session_state.df_interv

auto_equip_col = find_id_column(df_equip)
auto_name_col = find_name_column(df_equip)
auto_brand_col = find_brand_column(df_equip)
auto_uf_col = find_uf_column(df_equip)
auto_sn_col = find_sn_column(df_equip)
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
    "📚 Matrice des Pannes (Marques & Modèles)"
])

# =========================================================
# ONGLET 1 : PARC (FILTRE STRICT SUR N° ÉQUIPEMENT, UF, SN)
# =========================================================
with onglets_app[0]:
    st.subheader("Recherche ciblée dans le parc")
    
    if df_equip is None:
        st.error("⚠️ Fichier des équipements introuvable (`Equipements.xlsx`).")
    else:
        cols_trouvees = []
        if auto_equip_col:
            cols_trouvees.append(f"**N° Équipement** (`{auto_equip_col}`)")
        if auto_uf_col:
            cols_trouvees.append(f"**UF** (`{auto_uf_col}`)")
        if auto_sn_col:
            cols_trouvees.append(f"**N° Série** (`{auto_sn_col}`)")
            
        st.caption(f"🎯 Recherche stricte sur les colonnes : {' • '.join(cols_trouvees)}")

        q_eq = st.text_input(
            "Entrez un N° d'équipement, un code/nom d'UF ou un N° de série (SN) :",
            placeholder="Ex : 2026-0042, PEDIATRIE, SN-9941...",
            key="search_bar_eq"
        ).strip()

        if not q_eq:
            st.info("💡 Saisissez un N° d'équipement, une UF ou un SN pour afficher les fiches correspondantes.")
        else:
            candidats_cibles = [c for c in [auto_equip_col, auto_uf_col, auto_sn_col] if c and c in df_equip.columns]
            
            if not candidats_cibles:
                st.warning("Aucune colonne correspondante n'a été trouvée.")
                resultats_eq = pd.DataFrame()
            else:
                masque = pd.Series(False, index=df_equip.index)
                for col in candidats_cibles:
                    masque |= df_equip[col].astype(str).str.contains(q_eq, case=False, na=False)
                resultats_eq = df_equip[masque]

            nb_res = len(resultats_eq)
            if nb_res == 0:
                st.warning(f"Aucun équipement trouvé avec « {q_eq} » dans le N° d'équipement, l'UF ou le SN.")
            else:
                st.write(f"**{nb_res}** appareil(s) trouvé(s) (affichage limité aux 30 premiers) :")
                vue_reduite = resultats_eq.head(30)
                
                sel_eq = st.dataframe(
                    vue_reduite,
                    use_container_width=True,
                    on_select="rerun",
                    selection_mode="single-row"
                )

                selected_rows_eq = sel_eq.selection.rows if hasattr(sel_eq, "selection") else []
                if len(selected_rows_eq) > 0:
                    row_data = vue_reduite.iloc[selected_rows_eq[0]]
                    equip_val = row_data[auto_equip_col]

                    st.markdown("---")
                    st.markdown(f"### 📋 Fiche Équipement : **{equip_val}**")

                    tab_info, tab_hist = st.tabs(["📄 Signalétique", "🛠️ Historique des pannes"])

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
        st.error("⚠️ Fichier des interventions introuvable (`Interventions.xlsx`).")
    else:
        st.subheader("Registre complet des interventions")
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

        st.write(f"Affichage de **{len(filtered_in)}** intervention(s) :")

        select_in = st.dataframe(
            filtered_in.head(50),
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
# ONGLET 3 : DIAGNOSTIC (SÉLECTION PAR NOM & SAUVEGARDE PERMANENTE SUR DISQUE)
# =========================================================
with onglets_app[2]:
    col_d_left, col_d_right = st.columns([1, 2])

    with col_d_left:
        st.subheader("🏷️ Sélection de l'appareil par Modèle / Nom")
        
        if df_equip is not None and auto_name_col:
            modeles_uniques = sorted(df_equip[auto_name_col].dropna().astype(str).str.strip().unique().tolist())
        else:
            modeles_uniques = sorted(list({err.get("equipement", "") for err in st.session_state.error_db if err.get("equipement")}))

        recherche_nom = st.text_input("Filtrer la liste des équipements :", placeholder="Tapez le modèle ou nom...")
        
        if recherche_nom:
            modeles_affiches = [m for m in modeles_uniques if recherche_nom.lower() in m.lower()]
        else:
            modeles_affiches = modeles_uniques

        nom_choisi = st.selectbox(
            "Nom de l'équipement :", 
            modeles_affiches if modeles_affiches else ["Autre / Inconnu"],
            key="select_nom_equip"
        )

        marque_appareil = "Constructeur"
        service_app = "Atelier Biomédical"
        if df_equip is not None and auto_name_col:
            matching_rows = df_equip[df_equip[auto_name_col].astype(str).str.strip() == str(nom_choisi).strip()]
            if not matching_rows.empty:
                r_first = matching_rows.iloc[0]
                if auto_brand_col and auto_brand_col in r_first:
                    marque_appareil = str(r_first[auto_brand_col])
                if auto_uf_col and auto_uf_col in r_first:
                    service_app = str(r_first[auto_uf_col])

        st.info(f"**Modèle :** {nom_choisi}\n\n**Marque :** {marque_appareil}\n\n**Service :** {service_app}")

        st.markdown("---")
        st.subheader("📄 Scanner un Manuel Technique (PDF)")
        up_pdf = st.file_uploader("Déposer une notice ou manuel constructeur", type=["pdf"], key="pdf_auto_diag")
        
        if up_pdf is not None:
            # 1. Sauvegarde automatique du PDF localement
            chemin_sauvegarde = os.path.join(DOSSIER_MANUELS, up_pdf.name)
            with open(chemin_sauvegarde, "wb") as f_out:
                f_out.write(up_pdf.getbuffer())
            st.caption(f"💾 *Fichier PDF archivé dans `{chemin_sauvegarde}`*")

            # 2. Analyse et écriture directe sur disque dur
            if st.button("🚀 Analyser et sauvegarder définitivement", use_container_width=True):
                with st.spinner("Détection de la marque, du modèle et extraction des codes pannes..."):
                    marque_det, modele_det = detecter_marque_et_modele_pdf(up_pdf, df_equip)
                    
                    if nom_choisi and nom_choisi not in ["Autre / Inconnu"]:
                        modele_final = nom_choisi
                        marque_final = marque_appareil if marque_appareil != "Constructeur" else marque_det
                    else:
                        modele_final = modele_det
                        marque_final = marque_det

                    nouvelles_erreurs = parse_pdf_errors(up_pdf, marque_machine=marque_final, nom_machine=modele_final)
                    
                    if nouvelles_erreurs:
                        # Sauvegarde immédiate et définitive sur le disque (Anti-F5 Refresh)
                        sauvegarder_pannes_sur_disque(nouvelles_erreurs)
                        st.success(f"✔ **{len(nouvelles_erreurs)} codes** enregistrés de façon permanente pour **{marque_final} - {modele_final}** !")
                        st.rerun() # Rafraichit l'interface sans vider les données
                    else:
                        st.warning(f"Manuel enregistré pour **{marque_final} - {modele_final}**, mais aucun code standard détecté.")

        # Bouton export de secours vers GitHub
        df_export_p = pd.DataFrame(st.session_state.error_db)
        st.download_button(
            label="💾 Télécharger codes_pannes.csv",
            data=df_export_p.to_csv(index=False).encode("utf-8"),
            file_name="codes_pannes.csv",
            mime="text/csv",
            help="Optionnel: exportez pour publier sur GitHub",
            use_container_width=True
        )

    with col_d_right:
        st.subheader("🔍 Recherche de Solution Constructeur")
        symptom_in = st.text_input("Code alarme ou panne constatée :", value="E301", key="symp_diag")

        nom_upper = str(nom_choisi).upper()
        pannes_modele = [e for e in st.session_state.error_db if str(e.get("equipement", "")).upper() in nom_upper or nom_upper in str(e.get("equipement", "")).upper()]
        
        q_upper = symptom_in.upper()
        found = next((e for e in pannes_modele if e["code"] in q_upper or e["cause"].upper() in q_upper), None)
        if not found:
            found = next((e for e in st.session_state.error_db if e["code"] in q_upper or e["cause"].upper() in q_upper), None)

        if found:
            res_text = f"""✔ SOLUTION IDENTIFIÉE POUR {found.get('marque', '')} - {found.get('equipement', nom_choisi)} ({found['source']} - Page {found['page']}) :

- Code alarme : {found['code']} — {found['cause']}
- Action recommandée : {found['solution']}

- Contrôles de sécurité requis (NF EN 62353) :
  1. Continuité de la terre de protection (< 0,20 Ω).
  2. Mesure du courant de fuite à la terre / carcasse.
  3. Essai dynamique et validation des sécurités mécaniques."""
        else:
            res_text = f"""Procédure générale pour {nom_choisi} : {symptom_in}

1. Contrôler l'alimentation et les fusibles de la carte de puissance.
2. Vérifier les connecteurs et faisceaux liés au circuit de mesure.
3. Exécuter l'autodiagnostic en mode service constructeur.
4. Effectuer les contrôles NF EN 62353 avant remise au service."""

        st.text_area("Consignes de maintenance :", value=res_text, height=190)
        tech_nom = st.text_input("Technicien signataire :", value="Technicien Biomédical")

        pdf_doc = generate_certified_pdf(
            inv_id=nom_choisi, 
            equip_name=f"{marque_appareil} {nom_choisi}", 
            equip_service=service_app,
            symptom=symptom_in, 
            resolution=res_text, 
            tech_name=tech_nom
        )
        st.download_button(
            label="📄 Télécharger la Fiche Certifiée (PDF)",
            data=pdf_doc,
            file_name=f"Fiche_{nom_choisi}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )

# =========================================================
# ONGLET 4 : MATRICE (MARQUES & MODÈLES)
# =========================================================
with onglets_app[3]:
    st.subheader("📚 Matrice des Pannes (Classée par Marques & Modèles)")
    
    toutes_les_pannes = pd.DataFrame(st.session_state.error_db)
    
    if not toutes_les_pannes.empty:
        if "marque" not in toutes_les_pannes.columns:
            toutes_les_pannes["marque"] = "Général"

        marques_dispos = sorted([str(m) for m in toutes_les_pannes["marque"].dropna().unique() if str(m).strip() != ""])

        c_f1, c_f2, c_f3 = st.columns(3)
        with c_f1:
            choix_marque = st.selectbox("1. Choisir la Marque :", ["Toutes les marques"] + marques_dispos)

        if choix_marque != "Toutes les marques":
            sous_table_marque = toutes_les_pannes[toutes_les_pannes["marque"].astype(str) == choix_marque]
            modeles_dispos = sorted([str(e) for e in sous_table_marque["equipement"].dropna().unique() if str(e).strip() != ""])
        else:
            modeles_dispos = sorted([str(e) for e in toutes_les_pannes["equipement"].dropna().unique() if str(e).strip() != ""])

        with c_f2:
            choix_modele = st.selectbox("2. Choisir le Modèle :", ["Tous les modèles"] + modeles_dispos)

        with c_f3:
            recherche_texte = st.text_input("3. Mot-clé (code, alarme, pièce...) :")

        pannes_filtrees = toutes_les_pannes
        if choix_marque != "Toutes les marques":
            pannes_filtrees = pannes_filtrees[pannes_filtrees["marque"].astype(str) == choix_marque]
        if choix_modele != "Tous les modèles":
            pannes_filtrees = pannes_filtrees[pannes_filtrees["equipement"].astype(str) == choix_modele]
        if recherche_texte:
            pannes_filtrees = pannes_filtrees[pannes_filtrees.apply(
                lambda r: r.astype(str).str.contains(recherche_texte, case=False).any(), axis=1
            )]

        st.write(f"Affichage de **{len(pannes_filtrees)}** panne(s) répertoriée(s)")

        for _, err in pannes_filtrees.iterrows():
            with st.container(border=True):
                c1, c2, c3 = st.columns([1, 4, 1])
                with c1:
                    st.code(err.get("code", "N/A"))
                    marque_tag = err.get("marque", "Général")
                    equip_tag = err.get("equipement", "")
                    st.caption(f"🏭 **{marque_tag}**\n\n🏷️ **{equip_tag}**")
                with c2:
                    st.markdown(f"**Cause :** {err.get('cause', 'Non précisée')}")
                    st.write(err.get("solution", "Aucune solution renseignée"))
                with c3:
                    st.caption(f"{err.get('source', '')}\nPage {err.get('page', '')}")
    else:
        st.info("Aucune panne enregistrée.")
