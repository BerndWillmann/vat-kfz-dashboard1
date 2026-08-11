"""
VAT KFZ Werkstatt-Dashboard
============================
Standard-Upload -> automatisches Dashboard fuer Servicemeldungen & Auftraege.

Erwartetes Excel-Format (IMMER gleich, 2 Tabellenblaetter):
  Tabelle1 (Meldungen):
    Ursprungsmeldung, Folgemeldung, Meldungsart, Equipment, Meldung, Auftrag,
    Abschlussdatum, Abschlusszeit, Beschreibung, Prioritaet, Angelegt am,
    Verantw.ArbPl., Angelegt um, Gew.Ende, progn. Ende, Techn. Platz,
    Relevanz, Bezeichnung, Langtext i, Codegruppe, Codier.Grp.Text,
    Codierung, Codier.Code.Txt, Angel. von Name, Melder, Standort,
    Anwenderstat., Systemstatus

  Tabelle2 (Auftraege):
    Auftrag, Auftragsart, Eckstarttermin, Kurztext, Istarbeit Summe,
    Einheit Arbeit, GesKosten Ist, GesKosten Plan, Waehrung, Standortwerk,
    Techn. Platz, Verantw.ArbPl., Bezeichnung

Start lokal:  streamlit run app.py
"""

import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime

# ----------------------------------------------------------------------------
# Konfiguration
# ----------------------------------------------------------------------------
st.set_page_config(page_title="VAT KFZ Werkstatt-Dashboard", layout="wide")

WARTUNGS_ARTEN = ["WK", "WE", "WM", "WF"]
REPARATUR_ART = "MF"
AUSSCHLUSS_ART = "MM"
ERLOES_AUFTRAGSART = "VLAD"  # Leistung an Dritte

REQUIRED_COLS_T1 = [
    "Meldungsart", "Equipment", "Meldung", "Auftrag", "Abschlußdatum",
    "Beschreibung", "Priorität", "Angelegt am", "Gew.Ende", "Techn. Platz",
    "Bezeichnung", "Melder", "Systemstatus",
]
REQUIRED_COLS_T2 = [
    "Auftrag", "Auftragsart", "Eckstarttermin", "Kurztext", "Istarbeit Summe",
    "GesKosten Ist", "GesKosten Plan", "Techn. Platz", "Bezeichnung",
]


def fmt_eur(value: float) -> str:
    """Formatiert Zahl als Euro mit Tausenderpunkt und Komma-Dezimal (AT-Format)."""
    if pd.isna(value):
        value = 0
    s = f"{value:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} €"


def fmt_num(value: float, decimals: int = 0) -> str:
    if pd.isna(value):
        value = 0
    s = f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s


# ----------------------------------------------------------------------------
# Datei-Upload & Validierung
# ----------------------------------------------------------------------------
st.title("🔧 VAT KFZ – Werkstätten-Dashboard")
st.caption("Standard-Excel hochladen → Dashboard aktualisiert sich automatisch.")

uploaded_file = st.file_uploader(
    "Servicemeldungen-Excel hochladen (.xlsx, Tabelle1 + Tabelle2)", type=["xlsx"]
)

if uploaded_file is None:
    st.info("Bitte lade die standardisierte Excel-Datei hoch, um das Dashboard zu sehen.")
    st.stop()

try:
    xl = pd.ExcelFile(uploaded_file)
    df1 = pd.read_excel(xl, sheet_name="Tabelle1")
    df2 = pd.read_excel(xl, sheet_name="Tabelle2")
except Exception as e:
    st.error(f"Datei konnte nicht gelesen werden: {e}")
    st.stop()

missing1 = [c for c in REQUIRED_COLS_T1 if c not in df1.columns]
missing2 = [c for c in REQUIRED_COLS_T2 if c not in df2.columns]

if missing1 or missing2:
    st.error(
        "❌ Die Datei entspricht nicht dem Standard-Template. "
        "Bitte Spaltennamen prüfen und erneut hochladen.\n\n"
        f"Fehlende Spalten Tabelle1: {missing1}\n"
        f"Fehlende Spalten Tabelle2: {missing2}"
    )
    st.stop()

st.success(f"✅ Datei erkannt: {len(df1):,} Meldungen, {len(df2):,} Aufträge".replace(",", "."))

# ----------------------------------------------------------------------------
# Datenaufbereitung
# ----------------------------------------------------------------------------
df1["Angelegt am"] = pd.to_datetime(df1["Angelegt am"], errors="coerce")
df1["Abschlußdatum"] = pd.to_datetime(df1["Abschlußdatum"], errors="coerce")
df1["Gew.Ende"] = pd.to_datetime(df1["Gew.Ende"], errors="coerce")
df2["Eckstarttermin"] = pd.to_datetime(df2["Eckstarttermin"], errors="coerce")

df1["Ist_offen"] = df1["Abschlußdatum"].isna()
df1["Kategorie"] = "Sonstige"
df1.loc[df1["Meldungsart"].isin(WARTUNGS_ARTEN), "Kategorie"] = "Wartung"
df1.loc[df1["Meldungsart"] == REPARATUR_ART, "Kategorie"] = "Reparatur"
df1.loc[df1["Meldungsart"] == AUSSCHLUSS_ART, "Kategorie"] = "Ausschluss"

heute = pd.Timestamp(datetime.today().date())

# aktive Meldungen = ohne die ausgeschlossene Meldungsart MM
df_aktiv = df1[df1["Kategorie"] != "Ausschluss"].copy()

# ----------------------------------------------------------------------------
# Sidebar Filter
# ----------------------------------------------------------------------------
st.sidebar.header("Filter")

min_dat = df_aktiv["Angelegt am"].min()
max_dat = df_aktiv["Angelegt am"].max()
date_range = st.sidebar.date_input(
    "Zeitraum (Angelegt am)",
    value=(min_dat.date() if pd.notna(min_dat) else datetime.today().date(),
           max_dat.date() if pd.notna(max_dat) else datetime.today().date()),
)

techn_platz_options = sorted(df_aktiv["Techn. Platz"].dropna().unique().tolist())
selected_tp = st.sidebar.multiselect("Techn. Platz (Werkstatt)", techn_platz_options)

if isinstance(date_range, tuple) and len(date_range) == 2:
    start_d, end_d = date_range
    mask = (df_aktiv["Angelegt am"] >= pd.Timestamp(start_d)) & (
        df_aktiv["Angelegt am"] <= pd.Timestamp(end_d)
    )
    df_aktiv = df_aktiv[mask]

if selected_tp:
    df_aktiv = df_aktiv[df_aktiv["Techn. Platz"].isin(selected_tp)]

df_wartung = df_aktiv[df_aktiv["Kategorie"] == "Wartung"]
df_reparatur = df_aktiv[df_aktiv["Kategorie"] == "Reparatur"]

# ----------------------------------------------------------------------------
# KPI-Kacheln
# ----------------------------------------------------------------------------
st.subheader("📊 Kennzahlen im Überblick")

wartung_verzug = df_wartung[
    df_wartung["Ist_offen"] & (df_wartung["Gew.Ende"] < heute)
]
reparatur_ueber_10 = df_reparatur[
    df_reparatur["Ist_offen"]
    & ((heute - df_reparatur["Angelegt am"]).dt.days > 10)
]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Anzahl Wartungen", fmt_num(len(df_wartung)))
c2.metric("Anzahl Reparaturen", fmt_num(len(df_reparatur)))
c3.metric("⚠️ Wartungen im Verzug", fmt_num(len(wartung_verzug)))
c4.metric("⚠️ Reparaturen > 10 Tage offen", fmt_num(len(reparatur_ueber_10)))

c5, c6, c7 = st.columns(3)
c5.metric("Offene Wartungen (gesamt)", fmt_num(df_wartung["Ist_offen"].sum()))
c6.metric("Offene Reparaturen (gesamt)", fmt_num(df_reparatur["Ist_offen"].sum()))
c7.metric(
    "Ø Bearbeitungsdauer Reparaturen (Tage)",
    fmt_num(
        (
            df_reparatur.loc[~df_reparatur["Ist_offen"], "Abschlußdatum"]
            - df_reparatur.loc[~df_reparatur["Ist_offen"], "Angelegt am"]
        ).dt.days.mean(),
        1,
    ),
)

st.divider()

# ----------------------------------------------------------------------------
# Charts: Verteilung & Trend
# ----------------------------------------------------------------------------
col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**Verteilung Meldungsarten (aktiv, ohne MM)**")
    kat_counts = df_aktiv["Kategorie"].value_counts().reset_index()
    kat_counts.columns = ["Kategorie", "Anzahl"]
    fig1 = px.pie(kat_counts, names="Kategorie", values="Anzahl", hole=0.4)
    st.plotly_chart(fig1, use_container_width=True)

with col_b:
    st.markdown("**Wartungen vs. Reparaturen pro Monat**")
    trend = df_aktiv[df_aktiv["Kategorie"].isin(["Wartung", "Reparatur"])].copy()
    trend["Monat"] = trend["Angelegt am"].dt.to_period("M").astype(str)
    trend_g = trend.groupby(["Monat", "Kategorie"]).size().reset_index(name="Anzahl")
    fig2 = px.bar(trend_g, x="Monat", y="Anzahl", color="Kategorie", barmode="group")
    st.plotly_chart(fig2, use_container_width=True)

st.divider()

# ----------------------------------------------------------------------------
# Equipment / Werkstatt-Auslastung
# ----------------------------------------------------------------------------
st.subheader("🏭 Werkstätten- / Equipment-Auslastung (Techn. Platz)")

tp_stats = (
    df_aktiv.groupby("Techn. Platz")
    .agg(
        Meldungen_gesamt=("Meldung", "count"),
        Wartungen=("Kategorie", lambda x: (x == "Wartung").sum()),
        Reparaturen=("Kategorie", lambda x: (x == "Reparatur").sum()),
        Offen=("Ist_offen", "sum"),
    )
    .sort_values("Meldungen_gesamt", ascending=False)
    .reset_index()
    .head(15)
)
fig3 = px.bar(
    tp_stats, x="Techn. Platz", y=["Wartungen", "Reparaturen"],
    barmode="stack", title="Top 15 Techn. Plätze nach Meldungsvolumen",
)
st.plotly_chart(fig3, use_container_width=True)
st.dataframe(tp_stats, use_container_width=True, hide_index=True)

st.divider()

# ----------------------------------------------------------------------------
# Verzugslisten (offene, kritische Fälle)
# ----------------------------------------------------------------------------
st.subheader("⏰ Aktuell kritische Fälle")

tab1, tab2 = st.tabs(["Wartungen im Verzug", "Reparaturen > 10 Tage offen"])

with tab1:
    if wartung_verzug.empty:
        st.success("Keine Wartungen im Verzug 🎉")
    else:
        show_cols = ["Meldung", "Equipment", "Techn. Platz", "Bezeichnung",
                     "Angelegt am", "Gew.Ende", "Melder"]
        st.dataframe(
            wartung_verzug[show_cols].sort_values("Gew.Ende"),
            use_container_width=True, hide_index=True,
        )

with tab2:
    if reparatur_ueber_10.empty:
        st.success("Keine Reparaturen über 10 Tage offen 🎉")
    else:
        show_cols = ["Meldung", "Equipment", "Techn. Platz", "Bezeichnung",
                     "Angelegt am", "Melder"]
        df_show = reparatur_ueber_10[show_cols].copy()
        df_show["Tage offen"] = (heute - reparatur_ueber_10["Angelegt am"]).dt.days
        st.dataframe(
            df_show.sort_values("Tage offen", ascending=False),
            use_container_width=True, hide_index=True,
        )

st.divider()

# ----------------------------------------------------------------------------
# Kosten & Erlöse (Tabelle2 - Aufträge)
# ----------------------------------------------------------------------------
st.subheader("💶 Kosten & Erlöse (Aufträge)")

df2_calc = df2.copy()
df2_calc["Ist_Erloes"] = df2_calc["Auftragsart"] == ERLOES_AUFTRAGSART

erloes_kosten = df2_calc.astype(
    {"GesKosten Ist": "float64", "Istarbeit Summe": "float64"}
).groupby("Ist_Erloes").agg(
    Betrag=("GesKosten Ist", "sum"),
    Stunden=("Istarbeit Summe", "sum"),
    Anzahl=("Auftrag", "count"),
)

kosten_betrag = erloes_kosten.loc[False, "Betrag"] if False in erloes_kosten.index else 0
kosten_stunden = erloes_kosten.loc[False, "Stunden"] if False in erloes_kosten.index else 0
erloes_betrag = erloes_kosten.loc[True, "Betrag"] if True in erloes_kosten.index else 0
erloes_stunden = erloes_kosten.loc[True, "Stunden"] if True in erloes_kosten.index else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Interne Kosten (Ist)", fmt_eur(kosten_betrag))
k2.metric("Interne Stunden", f"{fmt_num(kosten_stunden, 1)} h")
k3.metric("Erlöse (VLAD – Leistung an Dritte)", fmt_eur(erloes_betrag))
k4.metric("Verrechnete Stunden an Dritte", f"{fmt_num(erloes_stunden, 1)} h")

st.markdown("**Kosten & Erlöse pro Monat (Eckstarttermin)**")
df2_calc["Monat"] = df2_calc["Eckstarttermin"].dt.to_period("M").astype(str)
df2_calc["Typ"] = df2_calc["Ist_Erloes"].map({True: "Erlöse (VLAD)", False: "Kosten (intern)"})
monat_g = df2_calc.groupby(["Monat", "Typ"])["GesKosten Ist"].sum().reset_index()
fig4 = px.bar(monat_g, x="Monat", y="GesKosten Ist", color="Typ", barmode="group")
fig4.update_yaxes(title="Betrag (€)")
st.plotly_chart(fig4, use_container_width=True)

st.markdown("**Auftragsarten im Detail**")
auftragsart_stats = (
    df2_calc.groupby("Auftragsart")
    .agg(
        Anzahl=("Auftrag", "count"),
        Stunden=("Istarbeit Summe", "sum"),
        Kosten_Ist=("GesKosten Ist", "sum"),
        Kosten_Plan=("GesKosten Plan", "sum"),
    )
    .sort_values("Kosten_Ist", ascending=False)
    .reset_index()
)
auftragsart_stats["Kosten_Ist"] = auftragsart_stats["Kosten_Ist"].apply(fmt_eur)
auftragsart_stats["Kosten_Plan"] = auftragsart_stats["Kosten_Plan"].apply(fmt_eur)
auftragsart_stats["Stunden"] = auftragsart_stats["Stunden"].apply(lambda v: fmt_num(v, 1))
st.dataframe(auftragsart_stats, use_container_width=True, hide_index=True)

st.divider()

# ----------------------------------------------------------------------------
# Kosten je Equipment / Techn. Platz (Verknüpfung Tabelle1 <-> Tabelle2)
# ----------------------------------------------------------------------------
st.subheader("🔗 Kosten je Equipment / Techn. Platz")

df1_link = df_aktiv.copy()
df1_link["Auftrag_key"] = df1_link["Auftrag"].dropna().astype("Int64").astype(str)
df2_link = df2.copy()
df2_link["Auftrag_key"] = df2_link["Auftrag"].astype(str)

merged = df1_link.merge(
    df2_link[["Auftrag_key", "GesKosten Ist", "Istarbeit Summe", "Auftragsart"]],
    on="Auftrag_key", how="left",
)

equip_kosten = (
    merged.groupby("Techn. Platz")
    .agg(
        Anzahl_Meldungen=("Meldung", "count"),
        Gesamtkosten=("GesKosten Ist", "sum"),
        Gesamtstunden=("Istarbeit Summe", "sum"),
    )
    .sort_values("Gesamtkosten", ascending=False)
    .reset_index()
    .head(15)
)
equip_kosten_display = equip_kosten.copy()
equip_kosten_display["Gesamtkosten"] = equip_kosten_display["Gesamtkosten"].apply(fmt_eur)
equip_kosten_display["Gesamtstunden"] = equip_kosten_display["Gesamtstunden"].apply(lambda v: fmt_num(v, 1))
st.dataframe(equip_kosten_display, use_container_width=True, hide_index=True)

fig5 = px.bar(equip_kosten, x="Techn. Platz", y="Gesamtkosten",
              title="Top 15 Techn. Plätze nach Gesamtkosten (€)")
st.plotly_chart(fig5, use_container_width=True)

st.caption(
    "Hinweis: MM-Meldungen sind aus allen Auswertungen ausgeschlossen. "
    "Als 'offen' gilt eine Meldung, solange kein Abschlußdatum eingetragen ist."
)
