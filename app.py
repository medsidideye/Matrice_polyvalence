import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Recherche OF / Moules / Machines",
    page_icon="🏭",
    layout="wide"
)

# =========================
# Fonctions utiles
# =========================

def nettoyer_texte(serie):
    return serie.astype(str).str.strip()

def convertir_excel_en_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultats")
    return output.getvalue()

def verifier_colonnes(df, colonnes_attendues):
    colonnes_manquantes = [col for col in colonnes_attendues if col not in df.columns]
    return colonnes_manquantes

# =========================
# Titre et introduction
# =========================

st.title("Plateforme de recherche OF / Moules / Machines")
st.markdown(
    """
Cette application permet de :

- charger les deux fichiers Excel,
- relier les données à partir du numéro d'OF,
- rechercher par **code article**,
- rechercher par **OF**,
- rechercher par **moule**,
- rechercher par **machine**,
- exporter les résultats en Excel.
"""
)

# =========================
# Upload des fichiers
# =========================

col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    fichier1 = st.file_uploader(
        "Importer le fichier 1 : moule par OF",
        type=["xlsx"]
    )

with col_upload2:
    fichier2 = st.file_uploader(
        "Importer le fichier 2 : liste OF",
        type=["xlsx"]
    )

if fichier1 is None or fichier2 is None:
    st.info("Veuillez importer les deux fichiers Excel pour commencer.")
    st.stop()

# =========================
# Lecture des fichiers
# =========================

try:
    df1 = pd.read_excel(fichier1)
    df2 = pd.read_excel(fichier2)
except Exception as e:
    st.error(f"Erreur lors de la lecture des fichiers : {e}")
    st.stop()

# =========================
# Vérification des colonnes
# =========================

colonnes_fichier1 = ["Ordre", "Désignation article", "libellé"]
colonnes_fichier2 = ["N_OF_GPAO", "REF_ARTICLE", "LIB_ARTICLE", "Date début réel", "N_POSTE"]

manquantes_f1 = verifier_colonnes(df1, colonnes_fichier1)
manquantes_f2 = verifier_colonnes(df2, colonnes_fichier2)

if manquantes_f1:
    st.error(f"Colonnes manquantes dans le fichier 1 : {manquantes_f1}")
    st.stop()

if manquantes_f2:
    st.error(f"Colonnes manquantes dans le fichier 2 : {manquantes_f2}")
    st.stop()

# =========================
# Préparation des données
# =========================

df1["Ordre_str"] = nettoyer_texte(df1["Ordre"])
df2["N_OF_GPAO_str"] = nettoyer_texte(df2["N_OF_GPAO"])

base = df2.merge(
    df1,
    left_on="N_OF_GPAO_str",
    right_on="Ordre_str",
    how="left"
)

base_recherche = base[
    [
        "N_OF_GPAO",
        "REF_ARTICLE",
        "LIB_ARTICLE",
        "Date début réel",
        "N_POSTE",
        "libellé"
    ]
].copy()

base_recherche.columns = [
    "OF",
    "Code article",
    "Libellé article",
    "Date",
    "Machine",
    "Moule"
]

base_recherche["OF"] = nettoyer_texte(base_recherche["OF"])
base_recherche["Code article"] = nettoyer_texte(base_recherche["Code article"])
base_recherche["Libellé article"] = nettoyer_texte(base_recherche["Libellé article"])
base_recherche["Machine"] = nettoyer_texte(base_recherche["Machine"])
base_recherche["Moule"] = nettoyer_texte(base_recherche["Moule"])
base_recherche["Date"] = pd.to_datetime(base_recherche["Date"], errors="coerce")

base_recherche = base_recherche.drop_duplicates()

# =========================
# Statistiques
# =========================

st.subheader("Indicateurs")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("Nombre de lignes", len(base_recherche))

with col2:
    st.metric("OF uniques", base_recherche["OF"].nunique())

with col3:
    st.metric("Articles uniques", base_recherche["Code article"].nunique())

with col4:
    st.metric("Machines uniques", base_recherche["Machine"].nunique())

# =========================
# Filtres de recherche
# =========================

st.subheader("Recherche")

col_rech1, col_rech2 = st.columns([1, 2])

with col_rech1:
    type_recherche = st.selectbox(
        "Choisir le type de recherche",
        ["Code article", "OF", "Moule", "Machine"]
    )

with col_rech2:
    valeur_recherche = st.text_input("Entrer une valeur")

resultat = base_recherche.copy()

if valeur_recherche:
    valeur = valeur_recherche.strip()

    if type_recherche == "Code article":
        resultat = resultat[
            resultat["Code article"].str.contains(valeur, case=False, na=False)
        ]

    elif type_recherche == "OF":
        resultat = resultat[
            resultat["OF"].str.contains(valeur, case=False, na=False)
        ]

    elif type_recherche == "Moule":
        resultat = resultat[
            resultat["Moule"].str.contains(valeur, case=False, na=False)
        ]

    elif type_recherche == "Machine":
        resultat = resultat[
            resultat["Machine"].str.contains(valeur, case=False, na=False)
        ]

# =========================
# Affichage des résultats
# =========================

st.subheader("Résultats")
st.write(f"Nombre de résultats : {len(resultat)}")

st.dataframe(resultat, use_container_width=True)

# =========================
# Export Excel
# =========================

excel_data = convertir_excel_en_bytes(resultat)

st.download_button(
    label="Télécharger les résultats en Excel",
    data=excel_data,
    file_name="resultats_recherche.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# =========================
# Aperçu de la base complète
# =========================

with st.expander("Afficher la base fusionnée complète"):
    st.dataframe(base_recherche, use_container_width=True)
