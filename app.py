import streamlit as st
import pandas as pd
import altair as alt

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
    font-size: 1.6rem;
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
    df1 = pd.read_excel("moule par OF.XLSX")
    df2 = pd.read_excel("Liste OF 01.04.24 au 03.04.26.xlsx")

    df1["Ordre_str"] = nettoyer_texte(df1["Ordre"])
    df2["N_OF_GPAO_str"] = nettoyer_texte(df2["N_OF_GPAO"])

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

    for col in ["OF", "Code article", "Libellé article", "Machine", "Moule"]:
        base_recherche[col] = nettoyer_texte(base_recherche[col])

    base_recherche["Moule"] = base_recherche["Moule"].replace(["nan", "None", ""], pd.NA)
    base_recherche["Libellé article"] = base_recherche["Libellé article"].replace(["nan", "None", ""], pd.NA)

    base_recherche["Date"] = pd.to_datetime(base_recherche["Date"], errors="coerce")
    base_recherche["Date"] = base_recherche["Date"].dt.strftime("%Y-%m-%d")
    base_recherche["Date"] = base_recherche["Date"].replace(["NaT", "nan", "None"], pd.NA)

    base_recherche["Moule"] = base_recherche["Moule"].fillna("Non renseigné")
    base_recherche["Libellé article"] = base_recherche["Libellé article"].fillna("Non renseigné")
    base_recherche["Date"] = base_recherche["Date"].fillna("Non renseignée")

    base_recherche = base_recherche.drop_duplicates().reset_index(drop=True)

    return base_recherche

base_recherche = charger_donnees()

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
# Indicateurs globaux
# =========================
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Enregistrements", len(base_filtre))
c2.metric("OF", base_filtre["OF"].nunique())
c3.metric("Articles", base_filtre["Code article"].nunique())
c4.metric("Moules", base_filtre["Moule"].nunique())
c5.metric("Machines", base_filtre["Machine"].nunique())

st.divider()

# =========================
# Indicateurs métier
# =========================
moule_top = (
    base_filtre.groupby("Moule")
    .size()
    .reset_index(name="Nombre")
    .sort_values(by="Nombre", ascending=False)
    .reset_index(drop=True)
)

article_top = (
    base_filtre.groupby("Code article")
    .size()
    .reset_index(name="Nombre")
    .sort_values(by="Nombre", ascending=False)
    .reset_index(drop=True)
)

of_top = (
    base_filtre.groupby("OF")
    .size()
    .reset_index(name="Nombre")
    .sort_values(by="Nombre", ascending=False)
    .reset_index(drop=True)
)

moule_top_val = moule_top.iloc[0]["Moule"] if len(moule_top) > 0 else "-"
moule_top_n = int(moule_top.iloc[0]["Nombre"]) if len(moule_top) > 0 else 0

article_top_val = article_top.iloc[0]["Code article"] if len(article_top) > 0 else "-"
article_top_n = int(article_top.iloc[0]["Nombre"]) if len(article_top) > 0 else 0

of_top_val = of_top.iloc[0]["OF"] if len(of_top) > 0 else "-"
of_top_n = int(of_top.iloc[0]["Nombre"]) if len(of_top) > 0 else 0

st.subheader("Indicateurs métier")
k1, k2, k3 = st.columns(3)
k1.metric("Moule le plus monté", moule_top_val, delta=f"{moule_top_n} fois")
k2.metric("Article le plus utilisé", article_top_val, delta=f"{article_top_n} fois")
k3.metric("OF le plus fréquent", of_top_val, delta=f"{of_top_n} fois")

st.divider()

# =========================
# Recherches d'abord
# =========================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Recherche par article",
    "Recherche par OF",
    "Recherche par moule",
    "Recherche par machine",
    "Analyse par machine",
    "Diagrammes globaux",
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
    st.subheader("Analyse : combien de fois sur chaque machine")

    type_recherche = st.selectbox(
        "Choisir le type de recherche",
        ["Article", "OF", "Moule"],
        key="type_analyse"
    )

    valeur_recherche = st.text_input("Entrer la valeur à analyser", key="valeur_analyse")

    if valeur_recherche:
        if type_recherche == "Article":
            resultat = base_recherche[
                base_recherche["Code article"].str.contains(valeur_recherche.strip(), case=False, na=False)
            ].copy()
            titre = f"Répartition par machine pour l'article : {valeur_recherche}"

        elif type_recherche == "OF":
            resultat = base_recherche[
                base_recherche["OF"] == valeur_recherche.strip()
            ].copy()
            titre = f"Répartition par machine pour l'OF : {valeur_recherche}"

        else:
            resultat = base_recherche[
                base_recherche["Moule"].str.contains(valeur_recherche.strip(), case=False, na=False)
            ].copy()
            titre = f"Répartition par machine pour le moule : {valeur_recherche}"

        if len(resultat) == 0:
            st.warning("Aucun résultat trouvé.")
        else:
            st.markdown(f"### {titre}")

            resume_machine = (
                resultat.groupby("Machine")
                .size()
                .reset_index(name="Nombre de fois")
                .sort_values(by="Nombre de fois", ascending=False)
                .reset_index(drop=True)
            )

            st.dataframe(resume_machine, use_container_width=True)

with tab6:
    st.subheader("Diagrammes globaux")

    # =========================
    # Listes complètes
    # =========================
    liste_complete_moules = pd.DataFrame({
        "Moule": sorted(base_recherche["Moule"].dropna().unique())
    })

    liste_complete_articles = pd.DataFrame({
        "Code article": sorted(base_recherche["Code article"].dropna().unique())
    })

    liste_complete_of = pd.DataFrame({
        "OF": sorted(base_recherche["OF"].dropna().unique())
    })

    # =========================
    # Comptages sur la base filtrée
    # =========================
    compte_moules = (
        base_filtre.groupby("Moule")
        .size()
        .reset_index(name="Nombre total de montages")
    )

    compte_articles = (
        base_filtre.groupby("Code article")
        .size()
        .reset_index(name="Nombre total d'utilisations")
    )

    compte_of = (
        base_filtre.groupby("OF")
        .size()
        .reset_index(name="Nombre total d'occurrences")
    )

    # =========================
    # Jointures pour inclure les 0
    # =========================
    all_moules = liste_complete_moules.merge(
        compte_moules,
        on="Moule",
        how="left"
    )

    all_articles = liste_complete_articles.merge(
        compte_articles,
        on="Code article",
        how="left"
    )

    all_of = liste_complete_of.merge(
        compte_of,
        on="OF",
        how="left"
    )

    # Remplacer les NaN par 0
    all_moules["Nombre total de montages"] = all_moules["Nombre total de montages"].fillna(0)
    all_articles["Nombre total d'utilisations"] = all_articles["Nombre total d'utilisations"].fillna(0)
    all_of["Nombre total d'occurrences"] = all_of["Nombre total d'occurrences"].fillna(0)

    # Trier
    all_moules = all_moules.sort_values(by="Nombre total de montages", ascending=False).reset_index(drop=True)
    all_articles = all_articles.sort_values(by="Nombre total d'utilisations", ascending=False).reset_index(drop=True)
    all_of = all_of.sort_values(by="Nombre total d'occurrences", ascending=False).reset_index(drop=True)

    # =========================
    # Top moules
    # =========================
    st.markdown("### Top moules")

    chart_moules = alt.Chart(all_moules).mark_bar().encode(
        x=alt.X("Moule:N", sort="-y", title="Numéro de moule"),
        y=alt.Y("Nombre total de montages:Q", title="Nombre de montages"),
        tooltip=["Moule", "Nombre total de montages"]
    ).properties(height=400)

    st.altair_chart(chart_moules, use_container_width=True)

    # =========================
    # Top articles
    # =========================
    st.markdown("### Top articles")

    chart_articles = alt.Chart(all_articles).mark_bar().encode(
        x=alt.X(field="Code article", type="nominal", sort="-y", title="Code article"),
        y=alt.Y(field="Nombre total dutilisations", type="quantitative", title="Nombre dutilisations"),
        tooltip=[
            alt.Tooltip(field="Code article", type="nominal", title="Code article"),
            alt.Tooltip(field="Nombre total dutilisations", type="quantitative", title="Nombre dutilisations")
        ]
    ).properties(height=400)

    st.altair_chart(chart_articles, use_container_width=True)

    # =========================
    # Top OF
    # =========================
    st.markdown("### Top OF")

    chart_of = alt.Chart(all_of).mark_bar().encode(
        x=alt.X("OF:N", sort="-y", title="OF"),
        y=alt.Y("Nombre total d'occurrences:Q", title="Nombre d'occurrences"),
        tooltip=["OF", "Nombre total d'occurrences"]
    ).properties(height=400)

    st.altair_chart(chart_of, use_container_width=True)
st.divider()
st.caption("Application Streamlit de traçabilité entre OF, articles, moules et machines.")
