import streamlit as st
from io import BytesIO
import html
import re
import base64
from datetime import date

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="CV Pro Studio",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# TEXTES
# ============================================================

TEXTES = {
    "Français": {
        "nom": "Nom complet",
        "titre": "Titre professionnel",
        "email": "E-mail",
        "telephone": "Téléphone",
        "adresse": "Adresse",
        "linkedin": "LinkedIn",
        "site": "Site web",
        "profil": "Profil professionnel",
        "formation": "Formations",
        "experience": "Expériences professionnelles",
        "competence": "Compétences",
        "langue": "Langues",
        "certification": "Certifications",
        "projet": "Projets",
        "association": "Associations et engagements",
        "interet": "Intérêts et loisirs",
    },
    "English": {
        "nom": "Full name",
        "titre": "Professional title",
        "email": "Email",
        "telephone": "Phone",
        "adresse": "Address",
        "linkedin": "LinkedIn",
        "site": "Website",
        "profil": "Professional profile",
        "formation": "Education",
        "experience": "Professional experience",
        "competence": "Skills",
        "langue": "Languages",
        "certification": "Certifications",
        "projet": "Projects",
        "association": "Associations and volunteering",
        "interet": "Interests and hobbies",
    },
}


# ============================================================
# MODÈLES
# ============================================================

MODELES = [
    "ATS",
    "Professionnel",
    "Européen",
    "International",
    "Académique",
    "Moderne",
    "Créatif",
    "Élégant",
    "Corporate",
    "Minimal",
]

COULEURS = {
    "Bleu professionnel": "#1F4E79",
    "Bleu moderne": "#2563EB",
    "Vert": "#16805B",
    "Rouge": "#B42318",
    "Violet": "#6D4AFF",
    "Orange": "#C65D00",
    "Gris anthracite": "#374151",
    "Noir": "#111111",
}

NIVEAUX = {
    "Français": [
        "Débutant",
        "Académique",
        "Intermédiaire",
        "Avancé",
    ],
    "English": [
        "Beginner",
        "Academic",
        "Intermediate",
        "Advanced",
    ],
}

MOIS_FR = [
    "Janvier",
    "Février",
    "Mars",
    "Avril",
    "Mai",
    "Juin",
    "Juillet",
    "Août",
    "Septembre",
    "Octobre",
    "Novembre",
    "Décembre",
]


# ============================================================
# SESSION STATE
# ============================================================

DEFAULTS = {
    "nom": "",
    "titre": "",
    "email": "",
    "telephone": "",
    "adresse": "",
    "linkedin": "",
    "site": "",
    "profil": "",
    "photo": None,
    "formation": [],
    "experiences": [],
    "competences": [],
    "langues": [],
    "certifications": [],
    "projets": [],
    "associations": [],
    "interets": [],
    "offre": "",
    "modele": "Professionnel",
    "couleur": "#1F4E79",
    "langue_cv": "Français",
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# OUTILS
# ============================================================

def esc(value):
    """Échappe le HTML."""
    if value is None:
        return ""
    return html.escape(str(value))


def couleur_rgb(hex_color):
    hex_color = hex_color.replace("#", "")
    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16),
    )


def format_mois_annee(mois, annee):
    return f"{mois} {annee}"


def choisir_mois_annee(label, key_prefix):
    """
    Sélection de date limitée au mois et à l'année.
    Aucun jour n'est demandé.
    """

    col1, col2 = st.columns(2)

    annee_actuelle = date.today().year

    with col1:
        mois = st.selectbox(
            f"{label} – Mois",
            MOIS_FR,
            index=0,
            key=f"{key_prefix}_mois",
        )

    with col2:
        annees = list(
            range(
                annee_actuelle + 10,
                1950 - 1,
                -1,
            )
        )

        annee = st.selectbox(
            f"{label} – Année",
            annees,
            index=10,
            key=f"{key_prefix}_annee",
        )

    return mois, annee


def section_card(numero, titre, couleur):
    """
    Affiche un titre de section sans faire apparaître le HTML brut.
    """

    st.html(
        f"""
        <div class="section-card"
             style="
                border-left: 6px solid {couleur};
                border-top: 1px solid {couleur}30;
                border-right: 1px solid {couleur}30;
                border-bottom: 1px solid {couleur}30;
             ">

            <div class="section-number"
                 style="background:{couleur};">
                {numero}
            </div>

            <div>
                <div class="section-title">
                    {esc(titre)}
                </div>

                <div class="section-line"
                     style="background:{couleur};">
                </div>
            </div>

        </div>
        """
    )


# ============================================================
# CSS
# ============================================================

COULEUR = st.session_state.couleur

st.markdown(
    f"""
    <style>

    .stApp {{
        background: #f5f7fb;
    }}

    .main .block-container {{
        max-width: 1450px;
        padding-top: 2rem;
        padding-bottom: 4rem;
    }}

    /* HEADER */

    .app-header {{
        background: linear-gradient(
            135deg,
            {COULEUR},
            #111827
        );
        padding: 30px;
        border-radius: 20px;
        margin-bottom: 25px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.12);
        color: white;
    }}

    .app-header h1 {{
        margin: 0;
        font-size: 34px;
        font-weight: 800;
    }}

    .app-header p {{
        margin-top: 8px;
        margin-bottom: 0;
        opacity: 0.9;
        font-size: 16px;
    }}

    /* CARTES SECTIONS */

    .section-card {{
        background: white;
        border-radius: 15px;
        padding: 18px 20px;
        margin-top: 10px;
        margin-bottom: 22px;
        display: flex;
        align-items: center;
        gap: 15px;
        box-shadow: 0 5px 18px rgba(0,0,0,0.06);
    }}

    .section-number {{
        width: 44px;
        height: 44px;
        min-width: 44px;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-size: 18px;
        font-weight: 800;
    }}

    .section-title {{
        font-size: 22px;
        font-weight: 800;
        color: #111827;
    }}

    .section-line {{
        width: 50px;
        height: 3px;
        border-radius: 10px;
        margin-top: 5px;
    }}

    /* CARTES INFO */

    .info-card {{
        background: white;
        padding: 20px;
        border-radius: 15px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 15px rgba(0,0,0,0.04);
    }}

    /* SIDEBAR */

    section[data-testid="stSidebar"] {{
        background: #111827;
    }}

    section[data-testid="stSidebar"] * {{
        color: white;
    }}

    /* INPUTS */

    div[data-baseweb="input"],
    div[data-baseweb="textarea"],
    div[data-baseweb="select"] {{
        border-radius: 10px;
    }}

    /* BOUTONS */

    .stButton > button {{
        border-radius: 10px;
        font-weight: 700;
        min-height: 42px;
    }}

    /* APERÇU */

    .cv-paper {{
        background: white;
        padding: 42px;
        border-radius: 8px;
        box-shadow: 0 10px 35px rgba(0,0,0,0.12);
        margin-top: 15px;
    }}

    .cv-name {{
        font-size: 31px;
        font-weight: 800;
        margin-bottom: 4px;
    }}

    .cv-title {{
        font-size: 17px;
        margin-bottom: 15px;
    }}

    .cv-section-title {{
        font-size: 17px;
        font-weight: 800;
        margin-top: 22px;
        margin-bottom: 7px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}

    .cv-line {{
        height: 2px;
        width: 45px;
        margin-bottom: 10px;
    }}

    .cv-item {{
        margin-bottom: 13px;
    }}

    .cv-date {{
        font-weight: 700;
        color: #4b5563;
        font-size: 13px;
    }}

    .cv-small {{
        color: #4b5563;
        font-size: 13px;
    }}

    .cv-contact {{
        color: #374151;
        font-size: 13px;
        line-height: 1.8;
    }}

    .skill {{
        display: inline-block;
        padding: 5px 10px;
        margin: 3px;
        border-radius: 15px;
        background: #f1f5f9;
        font-size: 12px;
        font-weight: 600;
    }}

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown("## 📄 CV PRO STUDIO")
    st.caption("Créateur de CV professionnel")

    st.markdown("---")

    modele = st.selectbox(
        "🎨 Modèle de CV",
        MODELES,
        index=MODELES.index(st.session_state.modele),
    )

    st.session_state.modele = modele

    couleur_nom = st.selectbox(
        "🎨 Couleur principale",
        list(COULEURS.keys()),
        index=list(COULEURS.values()).index(
            st.session_state.couleur
        ),
    )

    st.session_state.couleur = COULEURS[couleur_nom]

    langue_cv = st.selectbox(
        "🌐 Langue du CV",
        ["Français", "English"],
        index=0 if st.session_state.langue_cv == "Français" else 1,
    )

    st.session_state.langue_cv = langue_cv

    st.markdown("---")

    st.markdown("### Sections")

    st.markdown(
        """
        ① Informations  
        ② Profil  
        ③ Formations  
        ④ Expériences  
        ⑤ Compétences  
        ⑥ Langues  
        ⑦ Certifications  
        ⑧ Projets  
        ⑨ Associations  
        ⑩ Intérêts  
        ⑪ Offre  
        ⑫ Aperçu & Export
        """
    )


# ============================================================
# HEADER
# ============================================================

st.html(
    f"""
    <div class="app-header">

        <h1>CV Pro Studio</h1>

        <p>
            Créez un CV professionnel, moderne et prêt à envoyer.
        </p>

    </div>
    """
)


# ============================================================
# TABS
# ============================================================

tabs = st.tabs(
    [
        "① Informations",
        "② Profil",
        "③ Formations",
        "④ Expériences",
        "⑤ Compétences",
        "⑥ Langues",
        "⑦ Certifications",
        "⑧ Projets",
        "⑨ Associations",
        "⑩ Intérêts",
        "⑪ Offre",
        "⑫ Aperçu & Export",
    ]
)


# ============================================================
# 1. INFORMATIONS
# ============================================================

with tabs[0]:

    section_card(
        "1",
        "Informations personnelles",
        COULEUR,
    )

    col1, col2 = st.columns([2, 1])

    with col1:

        st.session_state.nom = st.text_input(
            "Nom complet",
            value=st.session_state.nom,
            placeholder="Ex. Amani Francis Kanga",
        )

        st.session_state.titre = st.text_input(
            "Titre professionnel",
            value=st.session_state.titre,
            placeholder="Ex. Chargé de projet environnemental",
        )

        st.session_state.email = st.text_input(
            "E-mail",
            value=st.session_state.email,
            placeholder="exemple@email.com",
        )

        st.session_state.telephone = st.text_input(
            "Téléphone",
            value=st.session_state.telephone,
            placeholder="+225 XX XX XX XX XX",
        )

        st.session_state.adresse = st.text_input(
            "Adresse",
            value=st.session_state.adresse,
            placeholder="Ville, Pays",
        )

        st.session_state.linkedin = st.text_input(
            "LinkedIn",
            value=st.session_state.linkedin,
            placeholder="linkedin.com/in/...",
        )

        st.session_state.site = st.text_input(
            "Site web",
            value=st.session_state.site,
            placeholder="www.monsite.com",
        )

    with col2:

        st.markdown("### Photo professionnelle")

        photo = st.file_uploader(
            "Importer une photo",
            type=["png", "jpg", "jpeg"],
        )

        if photo is not None:
            st.session_state.photo = photo.getvalue()
            st.image(
                st.session_state.photo,
                width=180,
            )


# ============================================================
# 2. PROFIL
# ============================================================

with tabs[1]:

    section_card(
        "2",
        "Profil professionnel",
        COULEUR,
    )

    st.session_state.profil = st.text_area(
        "Présentez votre profil professionnel",
        value=st.session_state.profil,
        height=180,
        placeholder=(
            "Présentez brièvement votre formation, "
            "vos principales compétences, votre expérience "
            "et votre objectif professionnel."
        ),
    )


# ============================================================
# 3. FORMATIONS
# ============================================================

with tabs[2]:

    section_card(
        "3",
        "Formations",
        COULEUR,
    )

    nombre = st.number_input(
        "Nombre de formations",
        min_value=0,
        max_value=10,
        value=max(1, len(st.session_state.formation)),
        step=1,
    )

    formations = []

    for i in range(int(nombre)):

        with st.expander(
            f"Formation n°{i + 1}",
            expanded=(i == 0),
        ):

            etablissement = st.text_input(
                "Établissement",
                key=f"formation_etablissement_{i}",
                placeholder="Université, école, institut...",
            )

            diplome = st.text_input(
                "Diplôme / Formation",
                key=f"formation_diplome_{i}",
                placeholder="Licence, Master, BTS...",
            )

            domaine = st.text_input(
                "Domaine / Spécialité",
                key=f"formation_domaine_{i}",
                placeholder="Ex. Études germaniques",
            )

            col1, col2 = st.columns(2)

            with col1:
                mois_debut, annee_debut = choisir_mois_annee(
                    "Date de début",
                    f"formation_{i}_debut",
                )

            en_cours = st.checkbox(
                "Formation en cours",
                key=f"formation_{i}_en_cours",
            )

            if en_cours:

                date_fin = "En cours"

            else:

                with col2:
                    mois_fin, annee_fin = choisir_mois_annee(
                        "Date de fin",
                        f"formation_{i}_fin",
                    )

                date_fin = format_mois_annee(
                    mois_fin,
                    annee_fin,
                )

            date_debut = format_mois_annee(
                mois_debut,
                annee_debut,
            )

            formations.append(
                {
                    "etablissement": etablissement,
                    "diplome": diplome,
                    "domaine": domaine,
                    "debut": date_debut,
                    "fin": date_fin,
                }
            )

    st.session_state.formation = formations


# ============================================================
# 4. EXPÉRIENCES
# ============================================================

with tabs[3]:

    section_card(
        "4",
        "Expériences professionnelles",
        COULEUR,
    )

    nombre = st.number_input(
        "Nombre d'expériences",
        min_value=0,
        max_value=10,
        value=max(1, len(st.session_state.experiences)),
        step=1,
    )

    experiences = []

    for i in range(int(nombre)):

        with st.expander(
            f"Expérience professionnelle n°{i + 1}",
            expanded=(i == 0),
        ):

            poste = st.text_input(
                "Poste occupé",
                key=f"exp_poste_{i}",
                placeholder="Ex. Assistant chargé de projet",
            )

            entreprise = st.text_input(
                "Entreprise / Organisation",
                key=f"exp_entreprise_{i}",
                placeholder="Nom de l'organisation",
            )

            lieu = st.text_input(
                "Lieu",
                key=f"exp_lieu_{i}",
                placeholder="Ville, Pays",
            )

            col1, col2 = st.columns(2)

            with col1:
                mois_debut, annee_debut = choisir_mois_annee(
                    "Date de début",
                    f"experience_{i}_debut",
                )

            en_cours = st.checkbox(
                "Poste actuel",
                key=f"experience_{i}_en_cours",
            )

            if en_cours:

                date_fin = "En cours"

            else:

                with col2:
                    mois_fin, annee_fin = choisir_mois_annee(
                        "Date de fin",
                        f"experience_{i}_fin",
                    )

                date_fin = format_mois_annee(
                    mois_fin,
                    annee_fin,
                )

            date_debut = format_mois_annee(
                mois_debut,
                annee_debut,
            )

            description = st.text_area(
                "Description / Missions",
                key=f"exp_description_{i}",
                height=130,
                placeholder=(
                    "Décrivez vos principales missions "
                    "et réalisations."
                ),
            )

            experiences.append(
                {
                    "poste": poste,
                    "entreprise": entreprise,
                    "lieu": lieu,
                    "debut": date_debut,
                    "fin": date_fin,
                    "description": description,
                }
            )

    st.session_state.experiences = experiences


# ============================================================
# 5. COMPÉTENCES
# ============================================================

with tabs[4]:

    section_card(
        "5",
        "Compétences",
        COULEUR,
    )

    nombre = st.number_input(
        "Nombre de compétences",
        min_value=0,
        max_value=20,
        value=max(1, len(st.session_state.competences)),
        step=1,
    )

    competences = []

    for i in range(int(nombre)):

        with st.expander(
            f"Compétence n°{i + 1}",
            expanded=(i == 0),
        ):

            nom = st.text_input(
                "Compétence",
                key=f"competence_nom_{i}",
                placeholder="Ex. Gestion de projet",
            )

            niveau = st.selectbox(
                "Niveau",
                (
                    [
                        "Débutant",
                        "Intermédiaire",
                        "Avancé",
                        "Expert",
                    ]
                    if st.session_state.langue_cv == "Français"
                    else [
                        "Beginner",
                        "Intermediate",
                        "Advanced",
                        "Expert",
                    ]
                ),
                key=f"competence_niveau_{i}",
            )

            competences.append(
                {
                    "nom": nom,
                    "niveau": niveau,
                }
            )

    st.session_state.competences = competences


# ============================================================
# 6. LANGUES
# ============================================================

with tabs[5]:

    section_card(
        "6",
        "Langues",
        COULEUR,
    )

    nombre = st.number_input(
        "Nombre de langues",
        min_value=0,
        max_value=10,
        value=max(1, len(st.session_state.langues)),
        step=1,
    )

    langues = []

    for i in range(int(nombre)):

        with st.expander(
            f"Langue n°{i + 1}",
            expanded=(i == 0),
        ):

            langue = st.text_input(
                "Langue",
                key=f"langue_nom_{i}",
                placeholder="Ex. Allemand",
            )

            niveau = st.selectbox(
                "Niveau",
                NIVEAUX[st.session_state.langue_cv],
                key=f"langue_niveau_{i}",
            )

            langues.append(
                {
                    "langue": langue,
                    "niveau": niveau,
                }
            )

    st.session_state.langues = langues


# ============================================================
# 7. CERTIFICATIONS
# ============================================================

with tabs[6]:

    section_card(
        "7",
        "Certifications",
        COULEUR,
    )

    nombre = st.number_input(
        "Nombre de certifications",
        min_value=0,
        max_value=10,
        value=max(1, len(st.session_state.certifications)),
        step=1,
    )

    certifications = []

    for i in range(int(nombre)):

        with st.expander(
            f"Certification n°{i + 1}",
            expanded=(i == 0),
        ):

            nom = st.text_input(
                "Nom de la certification",
                key=f"cert_nom_{i}",
                placeholder="Ex. Certification en intelligence artificielle",
            )

            organisme = st.text_input(
                "Organisme",
                key=f"cert_organisme_{i}",
                placeholder="Nom de l'organisme",
            )

            mois, annee = choisir_mois_annee(
                "Date d'obtention",
                f"certification_{i}",
            )

            date_certification = format_mois_annee(
                mois,
                annee,
            )

            certifications.append(
                {
                    "nom": nom,
                    "organisme": organisme,
                    "date": date_certification,
                }
            )

    st.session_state.certifications = certifications


# ============================================================
# 8. PROJETS
# ============================================================

with tabs[7]:

    section_card(
        "8",
        "Projets",
        COULEUR,
    )

    st.info(
        "Cette section est facultative. "
        "Si vous ne renseignez aucun projet, "
        "la section n'apparaîtra pas dans votre CV."
    )

    nombre = st.number_input(
        "Nombre de projets",
        min_value=0,
        max_value=10,
        value=0,
        step=1,
    )

    projets = []

    for i in range(int(nombre)):

        with st.expander(
            f"Projet n°{i + 1}",
            expanded=(i == 0),
        ):

            nom = st.text_input(
                "Nom du projet",
                key=f"projet_nom_{i}",
                placeholder="Ex. Eco-cir Bouaké",
            )

            role = st.text_input(
                "Votre rôle",
                key=f"projet_role_{i}",
                placeholder="Ex. Responsable du projet",
            )

            col1, col2 = st.columns(2)

            with col1:

                mois_debut, annee_debut = choisir_mois_annee(
                    "Date de début",
                    f"projet_{i}_debut",
                )

            en_cours = st.checkbox(
                "Projet en cours",
                key=f"projet_{i}_en_cours",
            )

            if en_cours:

                date_fin = "En cours"

            else:

                with col2:

                    mois_fin, annee_fin = choisir_mois_annee(
                        "Date de fin",
                        f"projet_{i}_fin",
                    )

                date_fin = format_mois_annee(
                    mois_fin,
                    annee_fin,
                )

            date_debut = format_mois_annee(
                mois_debut,
                annee_debut,
            )

            description = st.text_area(
                "Description",
                key=f"projet_description_{i}",
                height=130,
                placeholder=(
                    "Présentez brièvement le projet, "
                    "son objectif et vos réalisations."
                ),
            )

            projets.append(
                {
                    "nom": nom,
                    "role": role,
                    "debut": date_debut,
                    "fin": date_fin,
                    "description": description,
                }
            )

    st.session_state.projets = projets


# ============================================================
# 9. ASSOCIATIONS
# ============================================================

with tabs[8]:

    section_card(
        "9",
        "Associations et engagements",
        COULEUR,
    )

    nombre = st.number_input(
        "Nombre d'associations / engagements",
        min_value=0,
        max_value=10,
        value=0,
        step=1,
    )

    associations = []

    for i in range(int(nombre)):

        with st.expander(
            f"Association / engagement n°{i + 1}",
            expanded=(i == 0),
        ):

            organisation = st.text_input(
                "Organisation / Association",
                key=f"association_nom_{i}",
                placeholder="Nom de l'organisation",
            )

            role = st.text_input(
                "Fonction / Rôle",
                key=f"association_role_{i}",
                placeholder="Ex. Volontaire",
            )

            description = st.text_area(
                "Description",
                key=f"association_description_{i}",
                height=100,
                placeholder="Décrivez votre engagement.",
            )

            associations.append(
                {
                    "organisation": organisation,
                    "role": role,
                    "description": description,
                }
            )

    st.session_state.associations = associations


# ============================================================
# 10. INTÉRÊTS
# ============================================================

with tabs[9]:

    section_card(
        "10",
        "Intérêts et loisirs",
        COULEUR,
    )

    nombre = st.number_input(
        "Nombre d'intérêts",
        min_value=0,
        max_value=15,
        value=0,
        step=1,
    )

    interets = []

    for i in range(int(nombre)):

        with st.expander(
            f"Intérêt / loisir n°{i + 1}",
            expanded=(i == 0),
        ):

            interet = st.text_input(
                "Intérêt / loisir",
                key=f"interet_{i}",
                placeholder="Ex. Lecture",
            )

            interets.append(interet)

    st.session_state.interets = interets


# ============================================================
# 11. OFFRE D'EMPLOI
# ============================================================

with tabs[10]:

    section_card(
        "11",
        "Offre d'emploi",
        COULEUR,
    )

    st.write(
        "Collez ici une offre d'emploi pour analyser "
        "les mots-clés présents dans votre CV."
    )

    st.session_state.offre = st.text_area(
        "Texte de l'offre",
        value=st.session_state.offre,
        height=300,
        placeholder=(
            "Collez ici le texte de l'offre d'emploi..."
        ),
    )

    if st.button(
        "🔎 Analyser l'offre",
        use_container_width=True,
    ):

        offre = st.session_state.offre.lower()

        mots = re.findall(
            r"\b[a-zA-ZÀ-ÿ]{5,}\b",
            offre,
        )

        mots_frequents = {}

        for mot in mots:
            mots_frequents[mot] = mots_frequents.get(mot, 0) + 1

        mots_tries = sorted(
            mots_frequents.items(),
            key=lambda x: x[1],
            reverse=True,
        )

        mots_cles = [
            mot
            for mot, nombre in mots_tries[:30]
            if nombre >= 2
        ]

        st.markdown("### Mots-clés détectés")

        if mots_cles:
            st.write(", ".join(mots_cles))
        else:
            st.info(
                "Aucun mot-clé significatif détecté."
            )


# ============================================================
# CONSTRUCTION DU CV HTML
# ============================================================

def build_cv_html():

    couleur = st.session_state.couleur

    nom = esc(st.session_state.nom)
    titre = esc(st.session_state.titre)

    contact = []

    if st.session_state.email:
        contact.append(
            f"✉ {esc(st.session_state.email)}"
        )

    if st.session_state.telephone:
        contact.append(
            f"☎ {esc(st.session_state.telephone)}"
        )

    if st.session_state.adresse:
        contact.append(
            f"⌖ {esc(st.session_state.adresse)}"
        )

    if st.session_state.linkedin:
        contact.append(
            f"LinkedIn : {esc(st.session_state.linkedin)}"
        )

    if st.session_state.site:
        contact.append(
            f"Web : {esc(st.session_state.site)}"
        )

    photo_html = ""

    if st.session_state.photo:

        image_base64 = base64.b64encode(
            st.session_state.photo
        ).decode("utf-8")

        photo_html = f"""
        <img
            src="data:image/jpeg;base64,{image_base64}"
            style="
                width:110px;
                height:110px;
                object-fit:cover;
                border-radius:50%;
                border:4px solid {couleur};
                float:right;
                margin-left:20px;
            "
        />
        """

    html_cv = f"""
    <div class="cv-paper">

        {photo_html}

        <div class="cv-name"
             style="color:{couleur};">
            {nom or "Votre nom"}
        </div>

        <div class="cv-title">
            {titre or "Titre professionnel"}
        </div>

        <div class="cv-contact">
            {" &nbsp; | &nbsp; ".join(contact)}
        </div>

        <div style="clear:both;"></div>
    """

    # PROFIL

    if st.session_state.profil.strip():

        html_cv += f"""
        <div class="cv-section-title"
             style="color:{couleur};">
            {esc(TEXTES[st.session_state.langue_cv]["profil"])}
        </div>

        <div class="cv-line"
             style="background:{couleur};">
        </div>

        <div>
            {esc(st.session_state.profil)}
        </div>
        """

    # FORMATIONS

    formations_valides = [
        x for x in st.session_state.formation
        if x.get("diplome") or x.get("etablissement")
    ]

    if formations_valides:

        html_cv += f"""
        <div class="cv-section-title"
             style="color:{couleur};">
            {esc(TEXTES[st.session_state.langue_cv]["formation"])}
        </div>

        <div class="cv-line"
             style="background:{couleur};">
        </div>
        """

        for item in formations_valides:

            html_cv += f"""
            <div class="cv-item">

                <div class="cv-date">
                    {esc(item["debut"])} – {esc(item["fin"])}
                </div>

                <strong>
                    {esc(item["diplome"])}
                </strong>

                <div class="cv-small">
                    {esc(item["etablissement"])}
                    {(" – " + esc(item["domaine"])) if item["domaine"] else ""}
                </div>

            </div>
            """

    # EXPERIENCES

    experiences_valides = [
        x for x in st.session_state.experiences
        if x.get("poste") or x.get("entreprise")
    ]

    if experiences_valides:

        html_cv += f"""
        <div class="cv-section-title"
             style="color:{couleur};">
            {esc(TEXTES[st.session_state.langue_cv]["experience"])}
        </div>

        <div class="cv-line"
             style="background:{couleur};">
        </div>
        """

        for item in experiences_valides:

            description = esc(
                item["description"]
            ).replace("\n", "<br>")

            html_cv += f"""
            <div class="cv-item">

                <div class="cv-date">
                    {esc(item["debut"])} – {esc(item["fin"])}
                </div>

                <strong>
                    {esc(item["poste"])}
                </strong>

                <div class="cv-small">
                    {esc(item["entreprise"])}
                    {(" – " + esc(item["lieu"])) if item["lieu"] else ""}
                </div>

                <div style="margin-top:5px;">
                    {description}
                </div>

            </div>
            """

    # COMPÉTENCES

    competences_valides = [
        x for x in st.session_state.competences
        if x.get("nom")
    ]

    if competences_valides:

        html_cv += f"""
        <div class="cv-section-title"
             style="color:{couleur};">
            {esc(TEXTES[st.session_state.langue_cv]["competence"])}
        </div>

        <div class="cv-line"
             style="background:{couleur};">
        </div>
        """

        for item in competences_valides:

            html_cv += f"""
            <span class="skill">
                {esc(item["nom"])}
                {" – " + esc(item["niveau"]) if item["niveau"] else ""}
            </span>
            """

    # LANGUES

    langues_valides = [
        x for x in st.session_state.langues
        if x.get("langue")
    ]

    if langues_valides:

        html_cv += f"""
        <div class="cv-section-title"
             style="color:{couleur};">
            {esc(TEXTES[st.session_state.langue_cv]["langue"])}
        </div>

        <div class="cv-line"
             style="background:{couleur};">
        </div>
        """

        for item in langues_valides:

            html_cv += f"""
            <div class="cv-item">
                <strong>{esc(item["langue"])}</strong>
                ({esc(item["niveau"])})
            </div>
            """

    # CERTIFICATIONS

    certifications_valides = [
        x for x in st.session_state.certifications
        if x.get("nom")
    ]

    if certifications_valides:

        html_cv += f"""
        <div class="cv-section-title"
             style="color:{couleur};">
            {esc(TEXTES[st.session_state.langue_cv]["certification"])}
        </div>

        <div class="cv-line"
             style="background:{couleur};">
        </div>
        """

        for item in certifications_valides:

            html_cv += f"""
            <div class="cv-item">

                <div class="cv-date">
                    {esc(item["date"])}
                </div>

                <strong>
                    {esc(item["nom"])}
                </strong>

                <div class="cv-small">
                    {esc(item["organisme"])}
                </div>

            </div>
            """

    # PROJETS

    projets_valides = [
        x for x in st.session_state.projets
        if x.get("nom")
    ]

    if projets_valides:

        html_cv += f"""
        <div class="cv-section-title"
             style="color:{couleur};">
            {esc(TEXTES[st.session_state.langue_cv]["projet"])}
        </div>

        <div class="cv-line"
             style="background:{couleur};">
        </div>
        """

        for item in projets_valides:

            description = esc(
                item["description"]
            ).replace("\n", "<br>")

            html_cv += f"""
            <div class="cv-item">

                <div class="cv-date">
                    {esc(item["debut"])} – {esc(item["fin"])}
                </div>

                <strong>
                    {esc(item["nom"])}
                </strong>

                <div class="cv-small">
                    {esc(item["role"])}
                </div>

                <div style="margin-top:5px;">
                    {description}
                </div>

            </div>
            """

    # ASSOCIATIONS

    associations_valides = [
        x for x in st.session_state.associations
        if x.get("organisation")
    ]

    if associations_valides:

        html_cv += f"""
        <div class="cv-section-title"
             style="color:{couleur};">
            {esc(TEXTES[st.session_state.langue_cv]["association"])}
        </div>

        <div class="cv-line"
             style="background:{couleur};">
        </div>
        """

        for item in associations_valides:

            html_cv += f"""
            <div class="cv-item">

                <strong>
                    {esc(item["organisation"])}
                </strong>

                <div class="cv-small">
                    {esc(item["role"])}
                </div>

                <div>
                    {esc(item["description"])}
                </div>

            </div>
            """

    # INTÉRÊTS

    interets_valides = [
        x for x in st.session_state.interets
        if x
    ]

    if interets_valides:

        html_cv += f"""
        <div class="cv-section-title"
             style="color:{couleur};">
            {esc(TEXTES[st.session_state.langue_cv]["interet"])}
        </div>

        <div class="cv-line"
             style="background:{couleur};">
        </div>

        <div>
            {" • ".join(esc(x) for x in interets_valides)}
        </div>
        """

    html_cv += "</div>"

    return html_cv


# ============================================================
# WORD
# ============================================================

def generate_word():

    doc = Document()

    couleur = couleur_rgb(
        st.session_state.couleur
    )

    # Marges

    section = doc.sections[0]

    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    # Nom

    p = doc.add_paragraph()

    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run(
        st.session_state.nom or "Votre nom"
    )

    run.bold = True
    run.font.size = Pt(25)
    run.font.color.rgb = couleur

    # Titre

    p = doc.add_paragraph()

    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run(
        st.session_state.titre or "Titre professionnel"
    )

    run.bold = True
    run.font.size = Pt(14)

    # Contact

    contacts = []

    if st.session_state.email:
        contacts.append(
            st.session_state.email
        )

    if st.session_state.telephone:
        contacts.append(
            st.session_state.telephone
        )

    if st.session_state.adresse:
        contacts.append(
            st.session_state.adresse
        )

    if st.session_state.linkedin:
        contacts.append(
            st.session_state.linkedin
        )

    if contacts:

        p = doc.add_paragraph()

        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = p.add_run(
            " | ".join(contacts)
        )

        run.font.size = Pt(9)

    def word_section(titre):

        p = doc.add_paragraph()

        run = p.add_run(titre.upper())

        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = couleur

    # Profil

    if st.session_state.profil.strip():

        word_section(
            TEXTES[
                st.session_state.langue_cv
            ]["profil"]
        )

        p = doc.add_paragraph(
            st.session_state.profil
        )

        p.paragraph_format.space_after = Pt(6)

    # Formations

    formations = [
        x for x in st.session_state.formation
        if x.get("diplome") or x.get("etablissement")
    ]

    if formations:

        word_section(
            TEXTES[
                st.session_state.langue_cv
            ]["formation"]
        )

        for item in formations:

            p = doc.add_paragraph()

            r = p.add_run(
                f'{item["debut"]} – {item["fin"]}'
            )

            r.bold = True

            p.add_run(
                f'\n{item["diplome"]}'
            )

            if item["etablissement"]:
                p.add_run(
                    f'\n{item["etablissement"]}'
                )

            if item["domaine"]:
                p.add_run(
                    f' – {item["domaine"]}'
                )

    # Expériences

    experiences = [
        x for x in st.session_state.experiences
        if x.get("poste") or x.get("entreprise")
    ]

    if experiences:

        word_section(
            TEXTES[
                st.session_state.langue_cv
            ]["experience"]
        )

        for item in experiences:

            p = doc.add_paragraph()

            r = p.add_run(
                f'{item["debut"]} – {item["fin"]}'
            )

            r.bold = True

            p.add_run(
                f'\n{item["poste"]}'
            )

            if item["entreprise"]:
                p.add_run(
                    f' – {item["entreprise"]}'
                )

            if item["lieu"]:
                p.add_run(
                    f' – {item["lieu"]}'
                )

            if item["description"]:
                p.add_run(
                    f'\n{item["description"]}'
                )

    # Compétences

    competences = [
        x for x in st.session_state.competences
        if x.get("nom")
    ]

    if competences:

        word_section(
            TEXTES[
                st.session_state.langue_cv
            ]["competence"]
        )

        for item in competences:

            p = doc.add_paragraph(
                style="List Bullet"
            )

            p.add_run(
                item["nom"]
            ).bold = True

            if item["niveau"]:
                p.add_run(
                    f' – {item["niveau"]}'
                )

    # Langues

    langues = [
        x for x in st.session_state.langues
        if x.get("langue")
    ]

    if langues:

        word_section(
            TEXTES[
                st.session_state.langue_cv
            ]["langue"]
        )

        for item in langues:

            p = doc.add_paragraph(
                style="List Bullet"
            )

            p.add_run(
                item["langue"]
            ).bold = True

            p.add_run(
                f' ({item["niveau"]})'
            )

    # Certifications

    certifications = [
        x for x in st.session_state.certifications
        if x.get("nom")
    ]

    if certifications:

        word_section(
            TEXTES[
                st.session_state.langue_cv
            ]["certification"]
        )

        for item in certifications:

            p = doc.add_paragraph()

            p.add_run(
                f'{item["date"]} – '
            ).bold = True

            p.add_run(
                item["nom"]
            ).bold = True

            if item["organisme"]:
                p.add_run(
                    f' – {item["organisme"]}'
                )

    # Projets

    projets = [
        x for x in st.session_state.projets
        if x.get("nom")
    ]

    if projets:

        word_section(
            TEXTES[
                st.session_state.langue_cv
            ]["projet"]
        )

        for item in projets:

            p = doc.add_paragraph()

            p.add_run(
                f'{item["debut"]} – {item["fin"]}'
            ).bold = True

            p.add_run(
                f'\n{item["nom"]}'
            ).bold = True

            if item["role"]:
                p.add_run(
                    f' – {item["role"]}'
                )

            if item["description"]:
                p.add_run(
                    f'\n{item["description"]}'
                )

    # Associations

    associations = [
        x for x in st.session_state.associations
        if x.get("organisation")
    ]

    if associations:

        word_section(
            TEXTES[
                st.session_state.langue_cv
            ]["association"]
        )

        for item in associations:

            p = doc.add_paragraph()

            p.add_run(
                item["organisation"]
            ).bold = True

            if item["role"]:
                p.add_run(
                    f' – {item["role"]}'
                )

            if item["description"]:
                p.add_run(
                    f'\n{item["description"]}'
                )

    # Intérêts

    interets = [
        x for x in st.session_state.interets
        if x
    ]

    if interets:

        word_section(
            TEXTES[
                st.session_state.langue_cv
            ]["interet"]
        )

        p = doc.add_paragraph(
            " • ".join(interets)
        )

    output = BytesIO()

    doc.save(output)

    output.seek(0)

    return output


# ============================================================
# PDF
# ============================================================

def generate_pdf():

    output = BytesIO()

    couleur = HexColor(
        st.session_state.couleur
    )

    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=35,
        leftMargin=35,
        topMargin=35,
        bottomMargin=35,
    )

    styles = getSampleStyleSheet()

    name_style = ParagraphStyle(
        "CVName",
        parent=styles["Heading1"],
        fontSize=24,
        leading=28,
        textColor=couleur,
        spaceAfter=4,
    )

    title_style = ParagraphStyle(
        "CVTitle",
        parent=styles["Normal"],
        fontSize=11,
        leading=14,
        spaceAfter=7,
    )

    contact_style = ParagraphStyle(
        "Contact",
        parent=styles["Normal"],
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#374151"),
    )

    section_style = ParagraphStyle(
        "Section",
        parent=styles["Heading2"],
        fontSize=12,
        leading=15,
        textColor=couleur,
        spaceBefore=12,
        spaceAfter=5,
    )

    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontSize=9,
        leading=13,
        spaceAfter=5,
    )

    small_style = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#4B5563"),
    )

    story = []

    story.append(
        Paragraph(
            esc(
                st.session_state.nom
                or "Votre nom"
            ),
            name_style,
        )
    )

    story.append(
        Paragraph(
            esc(
                st.session_state.titre
                or "Titre professionnel"
            ),
            title_style,
        )
    )

    contacts = []

    if st.session_state.email:
        contacts.append(
            f"Email : {esc(st.session_state.email)}"
        )

    if st.session_state.telephone:
        contacts.append(
            f"Tél. : {esc(st.session_state.telephone)}"
        )

    if st.session_state.adresse:
        contacts.append(
            f"Adresse : {esc(st.session_state.adresse)}"
        )

    if st.session_state.linkedin:
        contacts.append(
            f"LinkedIn : {esc(st.session_state.linkedin)}"
        )

    if st.session_state.site:
        contacts.append(
            f"Web : {esc(st.session_state.site)}"
        )

    if contacts:

        story.append(
            Paragraph(
                " | ".join(contacts),
                contact_style,
            )
        )

    # Profil

    if st.session_state.profil.strip():

        story.append(
            Paragraph(
                esc(
                    TEXTES[
                        st.session_state.langue_cv
                    ]["profil"]
                ),
                section_style,
            )
        )

        story.append(
            Paragraph(
                esc(
                    st.session_state.profil
                ).replace("\n", "<br/>"),
                body_style,
            )
        )

    # Formations

    formations = [
        x for x in st.session_state.formation
        if x.get("diplome") or x.get("etablissement")
    ]

    if formations:

        story.append(
            Paragraph(
                esc(
                    TEXTES[
                        st.session_state.langue_cv
                    ]["formation"]
                ),
                section_style,
            )
        )

        for item in formations:

            story.append(
                Paragraph(
                    f"<b>{esc(item['debut'])} – "
                    f"{esc(item['fin'])}</b>",
                    small_style,
                )
            )

            story.append(
                Paragraph(
                    f"<b>{esc(item['diplome'])}</b>",
                    body_style,
                )
            )

            texte = esc(
                item["etablissement"]
            )

            if item["domaine"]:
                texte += (
                    " – "
                    + esc(item["domaine"])
                )

            story.append(
                Paragraph(
                    texte,
                    small_style,
                )
            )

    # Expériences

    experiences = [
        x for x in st.session_state.experiences
        if x.get("poste") or x.get("entreprise")
    ]

    if experiences:

        story.append(
            Paragraph(
                esc(
                    TEXTES[
                        st.session_state.langue_cv
                    ]["experience"]
                ),
                section_style,
            )
        )

        for item in experiences:

            story.append(
                Paragraph(
                    f"<b>{esc(item['debut'])} – "
                    f"{esc(item['fin'])}</b>",
                    small_style,
                )
            )

            story.append(
                Paragraph(
                    f"<b>{esc(item['poste'])}</b>",
                    body_style,
                )
            )

            entreprise = esc(
                item["entreprise"]
            )

            if item["lieu"]:
                entreprise += (
                    " – "
                    + esc(item["lieu"])
                )

            story.append(
                Paragraph(
                    entreprise,
                    small_style,
                )
            )

            if item["description"]:

                story.append(
                    Paragraph(
                        esc(
                            item["description"]
                        ).replace(
                            "\n",
                            "<br/>",
                        ),
                        body_style,
                    )
                )

    # Compétences

    competences = [
        x for x in st.session_state.competences
        if x.get("nom")
    ]

    if competences:

        story.append(
            Paragraph(
                esc(
                    TEXTES[
                        st.session_state.langue_cv
                    ]["competence"]
                ),
                section_style,
            )
        )

        texte = " • ".join(
            f"<b>{esc(x['nom'])}</b>"
            + (
                f" ({esc(x['niveau'])})"
                if x["niveau"]
                else ""
            )
            for x in competences
        )

        story.append(
            Paragraph(
                texte,
                body_style,
            )
        )

    # Langues

    langues = [
        x for x in st.session_state.langues
        if x.get("langue")
    ]

    if langues:

        story.append(
            Paragraph(
                esc(
                    TEXTES[
                        st.session_state.langue_cv
                    ]["langue"]
                ),
                section_style,
            )
        )

        for item in langues:

            story.append(
                Paragraph(
                    f"<b>{esc(item['langue'])}</b> "
                    f"({esc(item['niveau'])})",
                    body_style,
                )
            )

    # Certifications

    certifications = [
        x for x in st.session_state.certifications
        if x.get("nom")
    ]

    if certifications:

        story.append(
            Paragraph(
                esc(
                    TEXTES[
                        st.session_state.langue_cv
                    ]["certification"]
                ),
                section_style,
            )
        )

        for item in certifications:

            story.append(
                Paragraph(
                    f"<b>{esc(item['date'])}</b> – "
                    f"<b>{esc(item['nom'])}</b>"
                    + (
                        f" – {esc(item['organisme'])}"
                        if item["organisme"]
                        else ""
                    ),
                    body_style,
                )
            )

    # Projets

    projets = [
        x for x in st.session_state.projets
        if x.get("nom")
    ]

    if projets:

        story.append(
            Paragraph(
                esc(
                    TEXTES[
                        st.session_state.langue_cv
                    ]["projet"]
                ),
                section_style,
            )
        )

        for item in projets:

            story.append(
                Paragraph(
                    f"<b>{esc(item['debut'])} – "
                    f"{esc(item['fin'])}</b>",
                    small_style,
                )
            )

            story.append(
                Paragraph(
                    f"<b>{esc(item['nom'])}</b>"
                    + (
                        f" – {esc(item['role'])}"
                        if item["role"]
                        else ""
                    ),
                    body_style,
                )
            )

            if item["description"]:

                story.append(
                    Paragraph(
                        esc(
                            item["description"]
                        ).replace(
                            "\n",
                            "<br/>",
                        ),
                        body_style,
                    )
                )

    # Associations

    associations = [
        x for x in st.session_state.associations
        if x.get("organisation")
    ]

    if associations:

        story.append(
            Paragraph(
                esc(
                    TEXTES[
                        st.session_state.langue_cv
                    ]["association"]
                ),
                section_style,
            )
        )

        for item in associations:

            story.append(
                Paragraph(
                    f"<b>{esc(item['organisation'])}</b>"
                    + (
                        f" – {esc(item['role'])}"
                        if item["role"]
                        else ""
                    ),
                    body_style,
                )
            )

            if item["description"]:

                story.append(
                    Paragraph(
                        esc(
                            item["description"]
                        ).replace(
                            "\n",
                            "<br/>",
                        ),
                        small_style,
                    )
                )

    # Intérêts

    interets = [
        x for x in st.session_state.interets
        if x
    ]

    if interets:

        story.append(
            Paragraph(
                esc(
                    TEXTES[
                        st.session_state.langue_cv
                    ]["interet"]
                ),
                section_style,
            )
        )

        story.append(
            Paragraph(
                " • ".join(
                    esc(x)
                    for x in interets
                ),
                body_style,
            )
        )

    doc.build(story)

    output.seek(0)

    return output


# ============================================================
# SCORE ATS
# ============================================================

def calcul_ats():

    score = 0
    details = []

    if st.session_state.nom.strip():
        score += 10
        details.append("Nom renseigné")

    if st.session_state.titre.strip():
        score += 10
        details.append("Titre professionnel renseigné")

    if st.session_state.email.strip():
        score += 10
        details.append("E-mail renseigné")

    if st.session_state.telephone.strip():
        score += 10
        details.append("Téléphone renseigné")

    if st.session_state.profil.strip():
        score += 10
        details.append("Profil professionnel présent")

    if any(
        x.get("diplome")
        for x in st.session_state.formation
    ):
        score += 10
        details.append("Formation renseignée")

    if any(
        x.get("poste")
        for x in st.session_state.experiences
    ):
        score += 10
        details.append("Expérience renseignée")

    if any(
        x.get("nom")
        for x in st.session_state.competences
    ):
        score += 10
        details.append("Compétences renseignées")

    if any(
        x.get("langue")
        for x in st.session_state.langues
    ):
        score += 10
        details.append("Langues renseignées")

    if (
        st.session_state.certifications
        or st.session_state.projets
    ):
        score += 10
        details.append("Certifications ou projets présents")

    return score, details


# ============================================================
# 12. APERÇU ET EXPORT
# ============================================================

with tabs[11]:

    section_card(
        "12",
        "Aperçu & Export",
        COULEUR,
    )

    col1, col2 = st.columns([3, 1])

    with col1:

        st.markdown("### Aperçu du CV")

    with col2:

        score, details = calcul_ats()

        st.metric(
            "Score CV",
            f"{score}/100",
        )

    st.markdown(
        build_cv_html(),
        unsafe_allow_html=True,
    )

    st.markdown("---")

    st.markdown("### Exporter votre CV")

    col1, col2 = st.columns(2)

    with col1:

        word_file = generate_word()

        st.download_button(
            "📄 Télécharger en Word",
            data=word_file.getvalue(),
            file_name="CV_Professionnel.docx",
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with col2:

        pdf_file = generate_pdf()

        st.download_button(
            "📑 Télécharger en PDF",
            data=pdf_file.getvalue(),
            file_name="CV_Professionnel.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    st.markdown("---")

    st.markdown("### Analyse rapide du CV")

    if score >= 80:

        st.success(
            f"Votre CV contient actuellement "
            f"les principales informations attendues. "
            f"Score : {score}/100."
        )

    elif score >= 60:

        st.warning(
            f"Votre CV est bien avancé, mais certaines "
            f"sections peuvent encore être complétées. "
            f"Score : {score}/100."
        )

    else:

        st.info(
            f"Votre CV nécessite encore plusieurs "
            f"informations. Score actuel : {score}/100."
        )

    if details:

        st.write("Éléments détectés :")

        for detail in details:
            st.write(f"✓ {detail}")