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
# Diagrammes globaux
# =========================
st.subheader("Diagrammes globaux")

top_moules = (
    base_filtre.groupby("Moule")
    .size()
    .reset_index(name="Nombre de montages")
    .sort_values(by="Nombre de montages", ascending=False)
    .head(15)
    .reset_index(drop=True)
)

top_articles = (
    base_filtre.groupby("Code article")
    .size()
    .reset_index(name="Nombre d'utilisations")
    .sort_values(by="Nombre d'utilisations", ascending=False)
    .head(15)
    .reset_index(drop=True)
)

top_of = (
    base_filtre.groupby("OF")
    .size()
    .reset_index(name="Nombre d'occurrences")
    .sort_values(by="Nombre d'occurrences", ascending=False)
    .head(15)
    .reset_index(drop=True)
)

col_g1, col_g2, col_g3 = st.columns(3)

with col_g1:
    st.markdown("### Top moules")
    chart_moules = alt.Chart(top_moules).mark_bar().encode(
        x=alt.X("Moule:N", sort="-y", title="Moule"),
        y=alt.Y("Nombre de montages:Q", title="Nombre"),
        tooltip=["Moule", "Nombre de montages"]
    ).properties(height=350)
    st.altair_chart(chart_moules, use_container_width=True)

with col_g2:
    st.markdown("### Top articles")
    chart_articles = alt.Chart(top_articles).mark_bar().encode(
        x=alt.X("Code article:N", sort="-y", title="Article"),
        y=alt.Y("Nombre d'utilisations:Q", title="Nombre"),
        tooltip=["Code article", "Nombre d'utilisations"]
    ).properties(height=350)
    st.altair_chart(chart_articles, use_container_width=True)

with col_g3:
    st.markdown("### Top OF")
    chart_of = alt.Chart(top_of).mark_bar().encode(
        x=alt.X("OF:N", sort="-y", title="OF"),
        y=alt.Y("Nombre d'occurrences:Q", title="Nombre"),
        tooltip=["OF", "Nombre d'occurrences"]
    ).properties(height=350)
    st.altair_chart(chart_of, use_container_width=True)

st.divider()

# =========================
# Onglets
# =========================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Recherche par article",
    "Recherche par OF",
    "Recherche par moule",
    "Recherche par machine",
    "Analyse par machine",
    "Top moules interactif",
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
    st.subheader("Top moules interactif")
    st.write("Clique sur une barre pour voir le détail par machine.")

    top_moules_interactif = (
        base_filtre.groupby("Moule")
        .size()
        .reset_index(name="Nombre de montages")
        .sort_values(by="Nombre de montages", ascending=False)
        .head(20)
        .reset_index(drop=True)
    )

    selection = alt.selection_point(fields=["Moule"], name="select_moule")

    chart = (
        alt.Chart(top_moules_interactif)
        .mark_bar()
        .encode(
            x=alt.X("Moule:N", sort="-y", title="Moule"),
            y=alt.Y("Nombre de montages:Q", title="Nombre de montages"),
            tooltip=["Moule", "Nombre de montages"]
        )
        .add_params(selection)
        .properties(height=450)
    )

    event = st.altair_chart(
        chart,
        use_container_width=True,
        on_select="rerun",
        selection_mode="select_moule",
        key="chart_moules"
    )

    moule_selectionne = None

    try:
        points = event.selection["select_moule"]
        if "Moule" in points and len(points["Moule"]) > 0:
            moule_selectionne = points["Moule"][0]
    except Exception:
        moule_selectionne = None

    if moule_selectionne:
        st.markdown(f"### Détail pour le moule : {moule_selectionne}")

        detail_machine = (
            base_filtre[base_filtre["Moule"] == moule_selectionne]
            .groupby("Machine")
            .size()
            .reset_index(name="Nombre de fois")
            .sort_values(by="Nombre de fois", ascending=False)
            .reset_index(drop=True)
        )

        st.dataframe(detail_machine, use_container_width=True)

with tab7:
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
