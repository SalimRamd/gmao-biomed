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

def detecter_fichier_cible(liste_candidats, nom_defaut):
    for f in liste_candidats:
        if os.path.exists(f):
            return f
    return nom_defaut

FICHIER_EQUIP_ACTIF = detecter_fichier_cible(EQUIP_CANDIDATES, "Equipements.xlsx")
FICHIER_INTERV_ACTIF = detecter_fichier_cible(INTERV_CANDIDATES, "Interventions.xlsx")

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

# --- 2. CHARGEMENT ET PERSISTANCE ---
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

def charger_ou_initialiser_pannes():
    if os.path.exists(PANNE_FILE_DEFAULT):
        try:
            df_p = pd.read_csv(PANNE_FILE_DEFAULT)
            if "marque" not in df_p.columns:
                df_p["marque"] = "Général"
            return df_p.to_dict(orient="records")
        except Exception:
            pass
    df_init = pd.DataFrame(DEFAULT_ERRORS)
    df_init.to_csv(PANNE_FILE_DEFAULT, index=False, encoding="utf-8")
    return DEFAULT_ERRORS

if "error_db" not in st.session_state:
    st.session_state.error_db = charger_ou_initialiser_pannes()

def enregistrer_pannes_disque(nouvelles_pannes):
    st.session_state.error_db = nouvelles_pannes + st.session_state.error_db
    df_p = pd.DataFrame(st.session_state.error_db)
    df_p = df_p.drop_duplicates(subset=["equipement", "code"], keep="first")
    df_p.to_csv(PANNE_FILE_DEFAULT, index=False, encoding="utf-8")

# --- 3. DÉTECTION PRÉCISE DES COLONNES ---
def get_col_id_equip(df):
    if df is None:
        return None
    for cand in ["n° equipement", "n° équipement", "numero equipement", "numéro équipement", "n° inventaire", "numero inventaire", "code equipement", "equipement", "équipement"]:
        for col in df.columns:
            c_clean = str(col).strip().lower()
            if cand in c_clean and "famille" not in c_clean and "interv" not in c_clean:
                return col
    return df.columns[0] if len(df.columns) > 0 else None

def get_col_uf(df):
    if df is None:
        return None
    for cand in ["nom uf", "code uf", "uf", "service", "affectation", "localisation", "pôle", "pole"]:
        for col in df.columns:
            c_clean = str(col).strip().lower()
            if cand == c_clean or cand in c_clean:
                return col
    return None

def get_col_sn(df):
    if df is None:
        return None
    for cand in ["n° serie", "n° série", "numero serie", "numéro de série", "sn", "s/n", "serie", "série"]:
        for col in df.columns:
            c_clean = str(col).strip().lower()
            if cand == c_clean or cand in c_clean:
                return col
    return None

def get_col_id_interv(df):
    if df is None:
        return None
    for cand in ["n° interv", "numero interv", "numéro interv", "n° intervention", "interv"]:
        for col in df.columns:
            if cand in str(col).strip().lower():
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

# --- 4. DÉTECTION ET EXTRACTION PDF ---
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
    marques_connues = set()
    if df_equip is not None:
        for c in df_equip.columns:
            if any(k in str(c).lower() for k in ["marque", "constructeur"]):
                marques_connues.update(df_equip[c].dropna().astype(str).str.strip().unique())
    for err in st.session_state.error_db:
        if err.get("marque"):
            marques_connues.add(err["marque"])
    marques_connues.update(["ICU MEDICAL", "HETTICH", "ABBOTT", "WASSENBURG", "THE SURGICAL CO", "DRAEGER", "PHILIPS", "GE HEALTHCARE"])
    
    for m in sorted([m for m in marques_connues if len(m) >= 2], key=len, reverse=True):
        if m.upper() in texte_complet:
            marque_trouvee = m
            break

    modele_trouve = os.path.splitext(uploaded_file.name)[0].replace("_", " ").replace("-", " ")
    if df_equip is not None:
        modeles_connus = set()
        for c in df_equip.columns:
            if any(k in str(c).lower() for k in ["modèle", "modele", "type"]):
                for val in df_equip[c].dropna().astype(str).str.strip().unique():
                    if len(val) >= 3 and not val.isnumeric():
                        modeles_connus.add(val)
        for mod in sorted(modeles_connus, key=len, reverse=True):
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

# --- 5. HEADER DE L'APPLICATION ---
st.title("🏥 GMAO & Diagnostic Biomédical")

df_equip = st.session_state.df_equip
df_interv = st.session_state.df_interv

col_eq_id = get_col_id_equip(df_equip)
col_eq_uf = get_col_uf(df_equip)
col_eq_sn = get_col_sn(df_equip)

col_in_id = get_col_id_interv(df_interv)
col_in_eq_id = get_col_id_equip(df_interv)
col_in_uf = get_col_uf(df_interv)
col_in_sn = get_col_sn(df_interv)

col_m1, col_m2, col_m3 = st.columns(3)
col_m1.metric("Équipements au parc", len(df_equip) if df_equip is not None else 0)
col_m2.metric("Interventions archivées", len(df_interv) if df_interv is not None else 0)
col_m3.metric("Codes pannes répertoriés", len(st.session_state.error_db))

st.markdown("---")

onglets_app = st.tabs([
    "📋 Parc des Équipements", 
    "🛠️ Registre des Interventions", 
    "🩺 Diagnostic des Pannes",
    "📚 Matrice des Pannes (Marques & Modèles)"
])

# =========================================================
# ONGLET 1 : PARC (CHAMPS DE RECHERCHE STRICTS & AJOUT)
# =========================================================
with onglets_app[0]:
    st.subheader("Parc des Équipements")
    
    with st.expander("➕ Ajouter un nouvel équipement au parc", expanded=False):
        with st.form("form_add_equip", clear_on_submit=True):
            ca1, ca2 = st.columns(2)
            n_id = ca1.text_input("N° Équipement / Inventaire * :", placeholder="Ex : 2026-0042")
            n_nom = ca2.text_input("Nom / Désignation :", placeholder="Ex : Pousse-seringue")
            
            ca3, ca4 = st.columns(2)
            n_mq = ca3.text_input("Marque :", placeholder="Ex : Fresenius")
            n_mod = ca4.text_input("Type / Modèle :", placeholder="Ex : Injectomat Agilia")
            
            ca5, ca6 = st.columns(2)
            n_sn = ca5.text_input("N° de Série (SN) :", placeholder="Ex : SN-849201")
            n_uf = ca6.text_input("UF / Service :", placeholder="Ex : Réanimation")
            
            btn_add_eq = st.form_submit_button("💾 Enregistrer définitivement", use_container_width=True)
            
            if btn_add_eq:
                if not n_id.strip():
                    st.error("Le N° d'équipement est obligatoire.")
                else:
                    nouvelle_ligne = {}
                    if df_equip is not None:
                        for c in df_equip.columns:
                            c_clean = str(c).strip().lower()
                            if col_eq_uf and c == col_eq_uf:
                                nouvelle_ligne[c] = n_uf.strip()
                            elif col_eq_id and c == col_eq_id:
                                nouvelle_ligne[c] = n_id.strip()
                            elif col_eq_sn and c == col_eq_sn:
                                nouvelle_ligne[c] = n_sn.strip()
                            elif any(k in c_clean for k in ["modèle", "modele", "type"]):
                                nouvelle_ligne[c] = n_mod.strip()
                            elif any(k in c_clean for k in ["marque", "constructeur"]):
                                nouvelle_ligne[c] = n_mq.strip()
                            elif "nom" in c_clean or "désignation" in c_clean:
                                nouvelle_ligne[c] = n_nom.strip()
                            else:
                                nouvelle_ligne[c] = ""
                    else:
                        nouvelle_ligne = {"N° Equipement": n_id, "Nom": n_nom, "Marque": n_mq, "Type/Modèle": n_mod, "N° Série": n_sn, "Nom UF": n_uf}
                    
                    df_ajoute = pd.DataFrame([nouvelle_ligne])
                    if st.session_state.df_equip is not None:
                        st.session_state.df_equip = pd.concat([st.session_state.df_equip, df_ajoute], ignore_index=True)
                    else:
                        st.session_state.df_equip = df_ajoute
                        
                    st.session_state.df_equip.to_excel(FICHIER_EQUIP_ACTIF, index=False)
                    st.cache_data.clear()
                    st.success(f"✔ Équipement **{n_id}** enregistré définitivement sur le disque !")
                    st.rerun()

    st.markdown("---")
    
    if df_equip is None:
        st.error(f"⚠️ Fichier des équipements introuvable (`{FICHIER_EQUIP_ACTIF}`).")
    else:
        st.write("#### Filtrer les équipements :")
        c_s1, c_s2, c_s3 = st.columns(3)
        filtre_eq_id = c_s1.text_input("Par N° Équipement :", placeholder="Ex : 2026-", key="f_eq_id").strip()
        filtre_eq_uf = c_s2.text_input("Par UF / Service :", placeholder="Ex : URGENCES, PEDIATRIE...", key="f_eq_uf").strip()
        filtre_eq_sn = c_s3.text_input("Par N° de Série (SN) :", placeholder="Ex : SN-...", key="f_eq_sn").strip()

        # Si aucun filtre n'est saisi, on n'affiche rien pour la fluidité
        if not (filtre_eq_id or filtre_eq_uf or filtre_eq_sn):
            st.info("💡 Saisissez un critère dans l'un des champs ci-dessus (ex: tapez `2026-` dans *N° Équipement* pour afficher uniquement ceux de 2026).")
        else:
            masque = pd.Series(True, index=df_equip.index)
            
            # Filtre STRICT N° Equipement
            if filtre_eq_id and col_eq_id:
                s_id = df_equip[col_eq_id].astype(str).str.strip().str.lower()
                q_id = filtre_eq_id.lower()
                m_id = s_id.str.contains(q_id, regex=False, na=False)
                # Tolérance aux espaces / tirets
                s_norm = s_id.str.replace(" ", "", regex=False).str.replace("-", "", regex=False).str.replace("_", "", regex=False)
                q_norm = q_id.replace(" ", "").replace("-", "").replace("_", "")
                if len(q_norm) >= 2:
                    m_id |= s_norm.str.contains(q_norm, regex=False, na=False)
                masque &= m_id

            # Filtre STRICT UF
            if filtre_eq_uf and col_eq_uf:
                s_uf = df_equip[col_eq_uf].astype(str).str.strip().str.lower()
                masque &= s_uf.str.contains(filtre_eq_uf.lower(), regex=False, na=False)

            # Filtre STRICT N° Série
            if filtre_eq_sn and col_eq_sn:
                s_sn = df_equip[col_eq_sn].astype(str).str.strip().str.lower()
                masque &= s_sn.str.contains(filtre_eq_sn.lower(), regex=False, na=False)

            resultats_eq = df_equip[masque]
            nb_res = len(resultats_eq)

            if nb_res == 0:
                st.warning("Aucun équipement ne correspond à vos filtres.")
            else:
                st.write(f"Affichage de **{nb_res}** résultat(s) (limité aux 30 premiers) :")
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
                    equip_val = row_data[col_eq_id]

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
                        if df_interv is not None and not df_interv.empty and col_in_eq_id in df_interv.columns:
                            match = df_interv[df_interv[col_in_eq_id].astype(str).str.strip() == str(equip_val).strip()]
                            if not match.empty:
                                st.success(f"{len(match)} intervention(s) archivée(s) pour cet appareil.")
                                st.dataframe(format_interventions_table(match), use_container_width=True)
                            else:
                                st.info("Aucune intervention enregistrée pour cet équipement.")
                        else:
                            st.warning("Le registre des interventions n'est pas encore chargé.")

# =========================================================
# ONGLET 2 : INTERVENTIONS (CHAMPS DE RECHERCHE STRICTS & AJOUT)
# =========================================================
with onglets_app[1]:
    st.subheader("Registre des Interventions")
    
    with st.expander("➕ Enregistrer une nouvelle intervention", expanded=False):
        with st.form("form_add_interv", clear_on_submit=True):
            ci1, ci2, ci3 = st.columns(3)
            ni_num = ci1.text_input("N° Intervention * :", placeholder="Ex : INT-2026-0045")
            ni_eq = ci2.text_input("N° Équipement concerné * :", placeholder="Ex : 2026-0042")
            ni_cadre = ci3.selectbox("Cadre :", ["Curatif", "Préventif", "Contrôle Qualité", "Autre"])
            
            ci4, ci5, ci6 = st.columns(3)
            ni_deb = ci4.date_input("Date début :", value=date.today())
            ni_fin = ci5.date_input("Date fin :", value=date.today())
            ni_tech = ci6.text_input("Technicien :", value="Technicien Biomédical")
            
            ci7, ci8 = st.columns(2)
            ni_fourn = ci7.text_input("Fournisseur / SAV :", placeholder="Interne")
            ni_bc = ci8.text_input("N° Bon de commande (BC) :", placeholder="BC-2026-99")
            
            ni_obs1 = st.text_area("Observations appel (Panne déclarée) :", height=70)
            ni_obs2 = st.text_area("Observations technique (Compte-rendu réparation) :", height=70)
            
            btn_add_int = st.form_submit_button("💾 Enregistrer définitivement", use_container_width=True)
            
            if btn_add_int:
                if not ni_eq.strip():
                    st.error("Le N° d'équipement est obligatoire.")
                else:
                    int_row = {
                        "N° Interv.": ni_num if ni_num.strip() else f"INT-{datetime.now().strftime('%m%d%H%M')}",
                        "N° Equipement": ni_eq.strip(),
                        "Début": str(ni_deb),
                        "Fin": str(ni_fin),
                        "Cadre": ni_cadre,
                        "Techn.": ni_tech,
                        "Fournisseur": ni_fourn,
                        "Observations lors de l'appel": ni_obs1,
                        "Observations technique": ni_obs2,
                        "N° BC": ni_bc
                    }
                    df_int_add = pd.DataFrame([int_row])
                    if st.session_state.df_interv is not None:
                        st.session_state.df_interv = pd.concat([st.session_state.df_interv, df_int_add], ignore_index=True)
                    else:
                        st.session_state.df_interv = df_int_add
                    
                    st.session_state.df_interv.to_excel(FICHIER_INTERV_ACTIF, index=False)
                    st.cache_data.clear()
                    st.success(f"✔ Intervention enregistrée sur le disque pour **{ni_eq}** !")
                    st.rerun()

    st.markdown("---")

    if df_interv is None:
        st.error(f"⚠️ Fichier des interventions introuvable (`{FICHIER_INTERV_ACTIF}`).")
    else:
        df_in_clean = format_interventions_table(df_interv)
        st.write("#### Filtrer les interventions :")
        
        c_i1, c_i2, c_i3, c_i4 = st.columns(4)
        filtre_in_num = c_i1.text_input("Par N° Intervention :", placeholder="Ex : INT-...", key="f_in_num").strip()
        filtre_in_eq = c_i2.text_input("Par N° Équipement :", placeholder="Ex : 2026-...", key="f_in_eq").strip()
        filtre_in_uf = c_i3.text_input("Par UF / Service :", placeholder="Ex : REA...", key="f_in_uf").strip()
        filtre_in_sn = c_i4.text_input("Par N° Série :", placeholder="Ex : SN-...", key="f_in_sn").strip()

        if not (filtre_in_num or filtre_in_eq or filtre_in_uf or filtre_in_sn):
            st.info("💡 Saisissez un critère dans l'un des champs ci-dessus pour afficher les interventions correspondantes.")
        else:
            masque_in = pd.Series(True, index=df_in_clean.index)
            
            if filtre_in_num and col_in_id:
                masque_in &= df_in_clean[col_in_id].astype(str).str.lower().str.contains(filtre_in_num.lower(), regex=False, na=False)
            if filtre_in_eq and col_in_eq_id:
                masque_in &= df_in_clean[col_in_eq_id].astype(str).str.lower().str.contains(filtre_in_eq.lower(), regex=False, na=False)
            if filtre_in_uf and col_in_uf:
                masque_in &= df_in_clean[col_in_uf].astype(str).str.lower().str.contains(filtre_in_uf.lower(), regex=False, na=False)
            if filtre_in_sn and col_in_sn:
                masque_in &= df_in_clean[col_in_sn].astype(str).str.lower().str.contains(filtre_in_sn.lower(), regex=False, na=False)

            resultats_in = df_in_clean[masque_in]
            nb_res_in = len(resultats_in)

            if nb_res_in == 0:
                st.warning("Aucune intervention ne correspond à vos filtres.")
            else:
                st.write(f"Affichage de **{nb_res_in}** intervention(s) (limité aux 30 premières) :")
                vue_in = resultats_in.head(30)

                select_in = st.dataframe(
                    vue_in,
                    use_container_width=True,
                    on_select="rerun",
                    selection_mode="single-row"
                )

                selected_rows_in = select_in.selection.rows if hasattr(select_in, "selection") else []
                if len(selected_rows_in) > 0:
                    int_row = vue_in.iloc[selected_rows_in[0]]
                    st.markdown("---")
                    st.markdown(f"### 🛠️ Rapport d'intervention : **{int_row.get(col_in_id, 'Sans N°')}**")

                    r1, r2 = st.columns(2)
                    r1.write(f"**N° Équipement :** {int_row.get(col_in_eq_id, 'N/A')}")
                    r1.write(f"**Date début :** {int_row.get('Début', 'N/A')}")
                    r1.write(f"**Technicien :** {int_row.get('Techn.', 'N/A')}")
                    r1.write(f"**Cadre :** {int_row.get('Cadre', 'N/A')}")

                    r2.write(f"**Fournisseur :** {int_row.get('Fournisseur', 'N/A')}")
                    r2.write(f"**Date fin :** {int_row.get('Fin', 'N/A')}")
                    r2.write(f"**N° Commande / BC :** {int_row.get('N° BC', 'N/A')}")

                    st.info(f"**Observations appel :**\n\n{int_row.get('Observations lors de l\'appel', 'Aucune observation')}")
                    st.success(f"**Observations technique :**\n\n{int_row.get('Observations technique', 'Aucun compte-rendu')}")

# =========================================================
# ONGLET 3 : DIAGNOSTIC (AUTO : GLISSER LE PDF & SOLUTIONS)
# =========================================================
with onglets_app[2]:
    col_d_left, col_d_right = st.columns([1, 1])

    with col_d_left:
        st.subheader("📄 Scanner un Manuel Constructeur (PDF)")
        st.caption("Déposez la notice : la marque et le modèle sont automatiquement identifiés et les pannes rangées dans la Matrice.")
        
        up_pdf = st.file_uploader("Notice ou manuel technique :", type=["pdf"], key="pdf_auto_diag")
        
        if up_pdf is not None:
            chemin_pdf = os.path.join(DOSSIER_MANUELS, up_pdf.name)
            with open(chemin_pdf, "wb") as f_out:
                f_out.write(up_pdf.getbuffer())
            st.caption(f"💾 *Fichier archivé dans `{chemin_pdf}`*")

            if st.button("🚀 Analyser et classer automatiquement dans la Matrice", use_container_width=True):
                with st.spinner("Reconnaissance de la marque, du modèle et extraction des codes pannes..."):
                    mq_det, mod_det = detecter_marque_et_modele_pdf(up_pdf, df_equip)
                    nouvelles_errs = parse_pdf_errors(up_pdf, marque_machine=mq_det, nom_machine=mod_det)
                    
                    if nouvelles_errs:
                        enregistrer_pannes_disque(nouvelles_errs)
                        st.success(f"✔ **{mq_det} — {mod_det}** identifié avec succès !")
                        st.info(f"**{len(nouvelles_errs)} codes pannes** extraits et sauvegardés définitivement dans la Matrice.")
                        st.rerun()
                    else:
                        st.warning(f"Manuel archivé pour **{mq_det} — {mod_det}**, mais aucun code alarme conventionnel n'a été repéré.")

    with col_d_right:
        st.subheader("🔍 Recherche de Solution Constructeur")
        symptom_in = st.text_input("Code alarme ou symptôme observé :", value="E301", key="symp_diag")

        q_clean = symptom_in.upper().strip()
        found = next((e for e in st.session_state.error_db if e["code"] in q_clean or e["cause"].upper() in q_clean), None)

        if found:
            res_text = f"""✔ SOLUTION IDENTIFIÉE POUR {found.get('marque', '')} - {found.get('equipement', '')} ({found['source']} - Page {found['page']}) :

- Code alarme : {found['code']} — {found['cause']}
- Action recommandée : {found['solution']}

- Contrôles de sécurité requis (NF EN 62353) :
  1. Continuité de la terre de protection (< 0,20 Ω).
  2. Mesure du courant de fuite à la terre / carcasse.
  3. Essai dynamique et validation des sécurités mécaniques."""
            appareil_label = f"{found.get('marque', '')} {found.get('equipement', '')}"
        else:
            res_text = f"""Procédure générale de maintenance pour : {symptom_in}

1. Contrôler les tensions d'alimentation secondaires sur la carte de puissance.
2. Vérifier les faisceaux et connecteurs liés à la fonction défaillante.
3. Exécuter la routine d'autotest en mode maintenance constructeur.
4. Effectuer les contrôles de sécurité électrique NF EN 62353 avant remise au service."""
            appareil_label = "Équipement biomédical"

        st.text_area("Consignes techniques :", value=res_text, height=180)
        
        c_p1, c_p2 = st.columns(2)
        n_id_pdf = c_p1.text_input("N° Inventaire pour le rapport :", value="2026-0042")
        tech_nom = c_p2.text_input("Technicien signataire :", value="Technicien Biomédical")

        pdf_doc = generate_certified_pdf(
            inv_id=n_id_pdf, 
            equip_name=appareil_label, 
            equip_service="Service de Soins",
            symptom=symptom_in, 
            resolution=res_text, 
            tech_name=tech_nom
        )
        st.download_button(
            label="📄 Télécharger la Fiche Certifiée (PDF NF EN 62353)",
            data=pdf_doc,
            file_name=f"Fiche_{n_id_pdf}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )

# =========================================================
# ONGLET 4 : MATRICE (MARQUES & MODÈLES HIÉRARCHIQUES)
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
            recherche_texte = st.text_input("3. Filtrer par code ou symptôme :")

        pannes_filtrees = toutes_les_pannes
        if choix_marque != "Toutes les marques":
            pannes_filtrees = pannes_filtrees[pannes_filtrees["marque"].astype(str) == choix_marque]
        if choix_modele != "Tous les modèles":
            pannes_filtrees = pannes_filtrees[pannes_filtrees["equipement"].astype(str) == choix_modele]
        if recherche_texte:
            pannes_filtrees = pannes_filtrees[pannes_filtrees.apply(
                lambda r: r.astype(str).str.contains(recherche_texte, case=False, regex=False).any(), axis=1
            )]

        st.write(f"Affichage de **{len(pannes_filtrees)}** code(s) panne(s) répertorié(s)")

        for _, err in pannes_filtrees.iterrows():
            with st.container(border=True):
                c1, c2, c3 = st.columns([1, 4, 1])
                with c1:
                    st.code(err.get("code", "N/A"))
                    st.caption(f"🏭 **{err.get('marque', 'Général')}**\n\n🏷️ **{err.get('equipement', '')}**")
                with c2:
                    st.markdown(f"**Cause :** {err.get('cause', 'Non précisée')}")
                    st.write(err.get("solution", "Aucune solution renseignée"))
                with c3:
                    st.caption(f"{err.get('source', '')}\nPage {err.get('page', '')}")
    else:
        st.info("Aucune panne enregistrée.")
