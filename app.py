import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Traçabilité OF / Articles / Moules / Machines",
    page_icon="🏭",
    layout="wide"
)

st.title("Traçabilité OF / Articles / Moules / Machines")
st.write("Application de recherche par article, OF, moule et machine.")

@st.cache_data
def charger_donnees():
    df1 = pd.read_excel("moule par OF.XLSX")
    df2 = pd.read_excel("Liste OF 01.04.24 au 03.04.26.xlsx")

    # Préparer les clés de jointure
    df1["Ordre_str"] = df1["Ordre"].astype(str).str.strip()
    df2["N_OF_GPAO_str"] = df2["N_OF_GPAO"].astype(str).str.strip()

    # Même logique que dans le notebook
    base = df2.merge(
        df1,
        left_on="N_OF_GPAO_str",
        right_on="Ordre_str",
        how="inner"
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

    base_recherche["OF"] = base_recherche["OF"].astype(str).str.strip()
    base_recherche["Code article"] = base_recherche["Code article"].astype(str).str.strip()
    base_recherche["Libellé article"] = base_recherche["Libellé article"].astype(str).str.strip()
    base_recherche["Machine"] = base_recherche["Machine"].astype(str).str.strip()
    base_recherche["Moule"] = base_recherche["Moule"].astype(str).str.strip()

    # Conversion date
    base_recherche["Date"] = pd.to_datetime(base_recherche["Date"], errors="coerce")
    base_recherche["Date"] = base_recherche["Date"].dt.strftime("%Y-%m-%d")

    # Nettoyage affichage
    base_recherche["Moule"] = base_recherche["Moule"].replace(["nan", "None"], pd.NA).fillna("Non renseigné")
    base_recherche["Date"] = base_recherche["Date"].replace(["NaT", "nan", "None"], pd.NA).fillna("Non renseignée")

    base_recherche = base_recherche.drop_duplicates()

    return base_recherche

base_recherche = charger_donnees()

# Sidebar
st.sidebar.header("Filtres")

liste_articles = sorted(base_recherche["Code article"].dropna().unique().tolist())
liste_machines = sorted(base_recherche["Machine"].dropna().unique().tolist())
liste_moules = sorted(base_recherche["Moule"].dropna().unique().tolist())

filtre_article = st.sidebar.selectbox("Code article", ["Tous"] + liste_articles)
filtre_machine = st.sidebar.selectbox("Machine", ["Toutes"] + liste_machines)
filtre_moule = st.sidebar.selectbox("Moule", ["Tous"] + liste_moules)

base_filtre = base_recherche.copy()

if filtre_article != "Tous":
    base_filtre = base_filtre[base_filtre["Code article"] == filtre_article]

if filtre_machine != "Toutes":
    base_filtre = base_filtre[base_filtre["Machine"] == filtre_machine]

if filtre_moule != "Tous":
    base_filtre = base_filtre[base_filtre["Moule"] == filtre_moule]

# Indicateurs
c1, c2, c3, c4 = st.columns(4)
c1.metric("Lignes", len(base_filtre))
c2.metric("OF", base_filtre["OF"].nunique())
c3.metric("Articles", base_filtre["Code article"].nunique())
c4.metric("Moules", base_filtre["Moule"].nunique())

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Recherche par article",
    "Recherche par OF",
    "Recherche par moule",
    "Recherche par machine",
    "Base complète"
])

with tab1:
    st.subheader("Recherche par code article")
    article_input = st.text_input("Entrer un code article")
    if article_input:
        resultat = base_recherche[
            base_recherche["Code article"].str.contains(article_input.strip(), case=False, na=False)
        ][["Code article", "Libellé article", "OF", "Date", "Moule", "Machine"]].drop_duplicates()
        resultat = resultat.sort_values(by=["Date", "OF"])
        st.dataframe(resultat, use_container_width=True)

with tab2:
    st.subheader("Recherche par OF")
    of_input = st.text_input("Entrer un OF")
    if of_input:
        resultat = base_recherche[
            base_recherche["OF"] == of_input.strip()
        ].drop_duplicates().sort_values(by=["Date"])
        st.dataframe(resultat, use_container_width=True)

with tab3:
    st.subheader("Recherche par moule")
    moule_input = st.text_input("Entrer un moule")
    if moule_input:
        resultat = base_recherche[
            base_recherche["Moule"].str.contains(moule_input.strip(), case=False, na=False)
        ][["Moule", "Machine", "OF", "Code article", "Libellé article", "Date"]].drop_duplicates()
        resultat = resultat.sort_values(by=["Machine", "Date"])
        st.dataframe(resultat, use_container_width=True)

with tab4:
    st.subheader("Recherche par machine")
    machine_input = st.text_input("Entrer une machine")
    if machine_input:
        resultat = base_recherche[
            base_recherche["Machine"] == machine_input.strip()
        ][["Machine", "Moule", "OF", "Code article", "Libellé article", "Date"]].drop_duplicates()
        resultat = resultat.sort_values(by=["Moule", "Date"])
        st.dataframe(resultat, use_container_width=True)

with tab5:
    st.subheader("Base complète")
    st.dataframe(base_filtre.sort_values(by=["Date", "OF"]), use_container_width=True)

    csv = base_filtre.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Télécharger la base filtrée en CSV",
        data=csv,
        file_name="base_recherche_filtre.csv",
        mime="text/csv"
    )
