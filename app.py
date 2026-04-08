import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Traçabilité OF / Articles / Moules / Machines",
    page_icon="🏭",
    layout="wide"
)

# =========================
# Style
# =========================
st.markdown("""
<style>
.main-title {
    font-size: 2.4rem;
    font-weight: 700;
    margin-bottom: 0.2rem;
}
.sub-title {
    color: #666;
    margin-bottom: 1.2rem;
}
.block-container {
    padding-top: 1.5rem;
}
div[data-testid="stMetricValue"] {
    font-size: 1.7rem;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">Traçabilité OF / Articles / Moules / Machines</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Application de recherche par article, OF, moule et machine.</div>',
    unsafe_allow_html=True
)

# =========================
# Fonctions utilitaires
# =========================
def nettoyer_texte(serie):
    return serie.astype(str).str.strip()

@st.cache_data
def charger_donnees():
    # Chargement des fichiers Excel
    df1 = pd.read_excel("moule par OF.XLSX")
    df2 = pd.read_excel("Liste OF 01.04.24 au 03.04.26.xlsx")

    # Clés de jointure
    df1["Ordre_str"] = nettoyer_texte(df1["Ordre"])
    df2["N_OF_GPAO_str"] = nettoyer_texte(df2["N_OF_GPAO"])

    # Jointure
    base = df2.merge(
        df1,
        left_on="N_OF_GPAO_str",
        right_on="Ordre_str",
        how="inner"
    )

    # Colonnes utiles
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

    # Renommage
    base_recherche.columns = [
        "OF",
        "Code article",
        "Libellé article",
        "Date",
        "Machine",
        "Moule"
    ]

    # Nettoyage texte
    for col in ["OF", "Code article", "Libellé article", "Machine", "Moule"]:
        base_recherche[col] = nettoyer_texte(base_recherche[col])

    # Nettoyage des valeurs parasites
    base_recherche["Moule"] = base_recherche["Moule"].replace(["nan", "None", ""], pd.NA)
    base_recherche["Libellé article"] = base_recherche["Libellé article"].replace(["nan", "None", ""], pd.NA)

    # Conversion de la date
    base_recherche["Date"] = pd.to_datetime(base_recherche["Date"], errors="coerce")
    base_recherche["Date"] = base_recherche["Date"].dt.strftime("%Y-%m-%d")
    base_recherche["Date"] = base_recherche["Date"].replace(["NaT", "nan", "None"], pd.NA)

    # Remplissage affichage
    base_recherche["Moule"] = base_recherche["Moule"].fillna("Non renseigné")
    base_recherche["Libellé article"] = base_recherche["Libellé article"].fillna("Non renseigné")
    base_recherche["Date"] = base_recherche["Date"].fillna("Non renseignée")

    # Supprimer les doublons
    base_recherche = base_recherche.drop_duplicates().reset_index(drop=True)

    return base_recherche

# =========================
# Chargement principal
# =========================
try:
    base_recherche = charger_donnees()
except FileNotFoundError:
    st.error("Les fichiers Excel ne sont pas trouvés dans le dossier du projet.")
    st.info("Place les fichiers suivants dans le même dossier que app.py :")
    st.code("moule par OF.XLSX\nListe OF 01.04.24 au 03.04.26.xlsx")
    st.stop()
except Exception as e:
    st.error("Une erreur est survenue lors du chargement des données.")
    st.exception(e)
    st.stop()

# =========================
# Sidebar
# =========================
st.sidebar.header("Filtres globaux")

articles = sorted([
    x for x in base_recherche["Code article"].dropna().unique().tolist()
    if x and str(x).lower() != "nan"
])

machines = sorted([
    x for x in base_recherche["Machine"].dropna().unique().tolist()
    if x and str(x).lower() != "nan"
])

moules = sorted([
    x for x in base_recherche["Moule"].dropna().unique().tolist()
    if x and str(x).lower() != "nan"
])

filtre_article = st.sidebar.selectbox("Code article", ["Tous"] + articles)
filtre_machine = st.sidebar.selectbox("Machine", ["Toutes"] + machines)
filtre_moule = st.sidebar.selectbox("Moule", ["Tous"] + moules)

base_filtre = base_recherche.copy()

if filtre_article != "Tous":
    base_filtre = base_filtre[base_filtre["Code article"] == filtre_article]

if filtre_machine != "Toutes":
    base_filtre = base_filtre[base_filtre["Machine"] == filtre_machine]

if filtre_moule != "Tous":
    base_filtre = base_filtre[base_filtre["Moule"] == filtre_moule]

base_filtre = base_filtre.reset_index(drop=True)

# =========================
# Indicateurs
# =========================
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Enregistrements", len(base_filtre))
c2.metric("OF", base_filtre["OF"].nunique())
c3.metric("Articles", base_filtre["Code article"].nunique())
c4.metric("Moules", base_filtre["Moule"].nunique())
c5.metric("Machines", base_filtre["Machine"].nunique())

st.divider()

# =========================
# Onglets
# =========================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Recherche par article",
    "Recherche par OF",
    "Recherche par moule",
    "Recherche par machine",
    "Base complète"
])

with tab1:
    st.subheader("Recherche par code article")
    article_input = st.text_input("Entrer un code article", key="article_input")

    if article_input:
        resultat = base_recherche[
            base_recherche["Code article"].str.contains(article_input.strip(), case=False, na=False)
        ][["Code article", "Libellé article", "OF", "Date", "Moule", "Machine"]].drop_duplicates()

        resultat = resultat.sort_values(by=["Date", "OF"]).reset_index(drop=True)
        st.write(f"Résultats : {len(resultat)} ligne(s)")
        st.dataframe(resultat, use_container_width=True)

with tab2:
    st.subheader("Recherche par OF")
    of_input = st.text_input("Entrer un OF", key="of_input")

    if of_input:
        resultat = base_recherche[
            base_recherche["OF"] == of_input.strip()
        ].drop_duplicates().sort_values(by=["Date"]).reset_index(drop=True)

        st.write(f"Résultats : {len(resultat)} ligne(s)")
        st.dataframe(resultat, use_container_width=True)

with tab3:
    st.subheader("Recherche par moule")
    moule_input = st.text_input("Entrer un moule", key="moule_input")

    if moule_input:
        resultat = base_recherche[
            base_recherche["Moule"].str.contains(moule_input.strip(), case=False, na=False)
        ][["Moule", "Machine", "OF", "Code article", "Libellé article", "Date"]].drop_duplicates()

        resultat = resultat.sort_values(by=["Machine", "Date"]).reset_index(drop=True)
        st.write(f"Résultats : {len(resultat)} ligne(s)")
        st.dataframe(resultat, use_container_width=True)

with tab4:
    st.subheader("Recherche par machine")
    machine_input = st.text_input("Entrer une machine", key="machine_input")

    if machine_input:
        resultat = base_recherche[
            base_recherche["Machine"] == machine_input.strip()
        ][["Machine", "Moule", "OF", "Code article", "Libellé article", "Date"]].drop_duplicates()

        resultat = resultat.sort_values(by=["Moule", "Date"]).reset_index(drop=True)
        st.write(f"Résultats : {len(resultat)} ligne(s)")
        st.dataframe(resultat, use_container_width=True)

with tab5:
    st.subheader("Base complète")
    base_affichage = base_filtre.sort_values(by=["Date", "OF"]).reset_index(drop=True)
    st.dataframe(base_affichage, use_container_width=True, height=500)

    csv = base_affichage.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="Télécharger la base filtrée en CSV",
        data=csv,
        file_name="base_recherche_filtre.csv",
        mime="text/csv"
    )

st.divider()
st.caption("Application Streamlit de traçabilité entre OF, articles, moules et machines.")
