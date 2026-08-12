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

import os

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ----------------------------------------------------------------------------
# Konfiguration
# ----------------------------------------------------------------------------
st.set_page_config(page_title="VAT KFZ Werkstatt-Dashboard", layout="wide")

WARTUNGS_ARTEN = ["WK", "WE", "WM", "WF"]
REPARATUR_ART = "MF"
AUSSCHLUSS_ART = "MM"
ERLOES_AUFTRAGSART = "VLAD"  # Leistung an Dritte

# --- Review-Center Einstellungen (MF-Meldungen) --------------------------
MF_KRITISCH_TAGE = 10          # ab wann eine MF-Meldung als "kritisch" gilt
MF_ALARM_TAGE = 20             # ab wann eine MF-Meldung als "sehr kritisch" (rot) gilt
REVIEW_LOG_FILE = os.path.join(os.path.dirname(__file__), "mf_review_log.csv")
REVIEW_LOG_COLS = [
    "Meldung", "Geprüft", "Geprüft am", "Kommentar Teamleiter",
    "SAP schließen empfohlen",
]

# --- Werkstättenleistung / Optimierung ------------------------------------
ZIEL_DURCHLAUFZEIT_TAGE = 7     # Ziel-Bearbeitungsdauer (Benchmark) je Reparatur
ZIEL_AUSLASTUNG_STUNDEN_MONAT = 160  # Ziel-Ist-Stunden je Techn. Platz/Monat (Referenzwert)

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


def ampel_symbol(tage: float) -> str:
    """Gibt ein Ampel-Emoji je nach Alter (Tage offen) der Meldung zurueck."""
    if pd.isna(tage):
        return "⚪"
    if tage > MF_ALARM_TAGE:
        return "🔴"
    if tage > MF_KRITISCH_TAGE:
        return "🟠"
    return "🟡"


def style_tage_offen(df: pd.DataFrame, col: str = "Tage offen"):
    """Faerbt die 'Tage offen'-Spalte gruen->gelb->rot in einer (read-only) Tabelle."""
    def _color(v):
        if pd.isna(v):
            return ""
        if v > MF_ALARM_TAGE:
            return "background-color: #ff4d4d; color: white;"
        if v > MF_KRITISCH_TAGE:
            return "background-color: #ffb84d;"
        return "background-color: #fff2b3;"
    try:
        return df.style.map(_color, subset=[col])
    except AttributeError:
        # Fallback fuer aeltere Pandas-Versionen ohne Styler.map
        return df.style.applymap(_color, subset=[col])


def berechne_score(row) -> float:
    """Handlungsbedarf-Score = Tage offen + Prioritaets-Bonus (niedrige Prio-Zahl = dringender)."""
    tage = row.get("Tage offen", 0) or 0
    prio = row.get("Priorität", None)
    bonus = 0
    try:
        prio_num = float(prio)
        bonus = max(0.0, (5 - prio_num)) * 5
    except (TypeError, ValueError):
        bonus = 0
    return round(tage + bonus, 1)


def load_review_log() -> pd.DataFrame:
    """Laedt das persistierte Review-Log (Abhak-/Kommentarstatus der Teamleiter)."""
    if os.path.exists(REVIEW_LOG_FILE):
        try:
            log = pd.read_csv(REVIEW_LOG_FILE, dtype={"Meldung": str})
            log["Geprüft am"] = pd.to_datetime(log["Geprüft am"], errors="coerce")
            return log
        except Exception:
            pass
    return pd.DataFrame(columns=REVIEW_LOG_COLS)


def save_review_log(log: pd.DataFrame) -> None:
    """Speichert das Review-Log dauerhaft als CSV neben app.py (Upsert je Meldung)."""
    log = log.copy()
    log["Meldung"] = log["Meldung"].astype(str)
    log = log.drop_duplicates(subset="Meldung", keep="last")
    log.to_csv(REVIEW_LOG_FILE, index=False)


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

tab1, tab2 = st.tabs(["Wartungen im Verzug", "🔴 MF Review-Center (Reparaturen)"])

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
    # -- Basisdaten aufbereiten: alle offenen MF-Meldungen + Alter in Tagen ---
    df_mf_offen = df_reparatur[df_reparatur["Ist_offen"]].copy()
    df_mf_offen["Tage offen"] = (heute - df_mf_offen["Angelegt am"]).dt.days
    mf_kritisch = df_mf_offen[df_mf_offen["Tage offen"] > MF_KRITISCH_TAGE].copy()

    if mf_kritisch.empty:
        st.success(f"Keine MF-Meldungen über {MF_KRITISCH_TAGE} Tage offen 🎉")
    else:
        mf_kritisch["Ampel"] = mf_kritisch["Tage offen"].apply(ampel_symbol)
        mf_kritisch["Score"] = mf_kritisch.apply(berechne_score, axis=1)
        mf_kritisch["Meldung"] = mf_kritisch["Meldung"].astype(str)
        mf_kritisch = mf_kritisch.sort_values("Score", ascending=False)

        # -- 1) Warnbanner ----------------------------------------------------
        n_alarm = (mf_kritisch["Tage offen"] > MF_ALARM_TAGE).sum()
        if n_alarm > 0:
            st.error(
                f"⚠️ **{len(mf_kritisch)} MF-Meldungen** sind länger als {MF_KRITISCH_TAGE} Tage offen – "
                f"davon **{n_alarm} über {MF_ALARM_TAGE} Tage** (🔴 sehr kritisch, Teamleiter bitte sofort prüfen!)"
            )
        else:
            st.warning(
                f"⚠️ **{len(mf_kritisch)} MF-Meldungen** sind länger als {MF_KRITISCH_TAGE} Tage offen "
                "und sollten von den Teamleitern geprüft werden."
            )

        # -- KPI-Kacheln Review-Center -----------------------------------------
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Kritische MF-Meldungen", fmt_num(len(mf_kritisch)))
        r2.metric("Älteste offen (Tage)", fmt_num(mf_kritisch["Tage offen"].max()))
        r3.metric("🔴 Sehr kritisch (>%d Tage)" % MF_ALARM_TAGE, fmt_num(n_alarm))
        r4.metric("Ø Score Handlungsbedarf", fmt_num(mf_kritisch["Score"].mean(), 1))

        st.markdown("---")

        # -- 2) Ampel-Karten: Top 5 kritischste Fälle --------------------------
        st.markdown("#### 🚨 Top 5 – höchster Handlungsbedarf")
        top5 = mf_kritisch.head(5)
        card_cols = st.columns(len(top5)) if len(top5) > 0 else []
        for col, (_, row) in zip(card_cols, top5.iterrows()):
            with col:
                with st.container(border=True):
                    st.markdown(f"**{row['Ampel']} Meldung {row['Meldung']}**")
                    st.caption(f"{row.get('Equipment', '')} · {row.get('Techn. Platz', '')}")
                    st.metric("Tage offen", fmt_num(row["Tage offen"]))
                    st.caption(f"Score: {row['Score']:.0f} · Melder: {row.get('Melder', '–')}")
                    with st.expander("Details"):
                        for feld in ["Bezeichnung", "Beschreibung", "Priorität",
                                     "Langtext i", "Codier.Grp.Text", "Codierung",
                                     "Verantw.ArbPl.", "Systemstatus"]:
                            if feld in row.index and pd.notna(row[feld]):
                                st.write(f"**{feld}:** {row[feld]}")

        st.markdown("---")

        # -- 3) Gruppierung nach Teamleiter (Verantw.ArbPl.) -------------------
        st.markdown("#### 👤 Übersicht je Teamleiter (Verantw.ArbPl.)")
        tl_stats = (
            mf_kritisch.groupby("Verantw.ArbPl.")
            .agg(
                Anzahl=("Meldung", "count"),
                Älteste_Tage=("Tage offen", "max"),
                Ø_Score=("Score", "mean"),
            )
            .sort_values("Anzahl", ascending=False)
            .reset_index()
        )
        tl_stats["Ø_Score"] = tl_stats["Ø_Score"].round(1)
        st.dataframe(tl_stats, use_container_width=True, hide_index=True)

        tl_options = ["Alle"] + sorted(mf_kritisch["Verantw.ArbPl."].dropna().unique().tolist())
        selected_tl = st.selectbox("🔎 Nach Teamleiter filtern (Verantw.ArbPl.)", tl_options)
        mf_view = mf_kritisch if selected_tl == "Alle" else mf_kritisch[mf_kritisch["Verantw.ArbPl."] == selected_tl]

        st.markdown("---")

        # -- 4) Review-Log laden & mit aktuellen Meldungen mergen --------------
        st.markdown("#### ✅ Review-Queue – prüfen, kommentieren, zum Schließen vormerken")
        review_log = load_review_log()
        merge_cols = ["Meldung", "Ampel", "Score", "Tage offen", "Equipment",
                      "Techn. Platz", "Bezeichnung", "Verantw.ArbPl.", "Melder", "Auftrag"]
        merge_cols = [c for c in merge_cols if c in mf_view.columns]
        base = mf_view[merge_cols].copy()
        merged_view = base.merge(review_log, on="Meldung", how="left")
        merged_view["Geprüft"] = merged_view["Geprüft"].map(lambda v: bool(v) if pd.notna(v) else False)
        merged_view["SAP schließen empfohlen"] = merged_view["SAP schließen empfohlen"].map(
            lambda v: bool(v) if pd.notna(v) else False
        )
        merged_view["Kommentar Teamleiter"] = merged_view["Kommentar Teamleiter"].fillna("").astype(str)

        # Tage seit letzter Prüfung (statt "letzte SAP-Änderung", da SAP-Export
        # kein Aenderungsdatum liefert - stattdessen: letzte Teamleiter-Pruefung)
        merged_view["Tage seit Prüfung"] = merged_view["Geprüft am"].apply(
            lambda d: (heute - pd.Timestamp(d)).days if pd.notna(d) else None
        )

        edited = st.data_editor(
            merged_view.sort_values("Score", ascending=False),
            use_container_width=True,
            hide_index=True,
            disabled=[c for c in merge_cols if c != "Meldung"] + ["Tage seit Prüfung", "Geprüft am"],
            column_config={
                "Geprüft": st.column_config.CheckboxColumn("Geprüft ✔️"),
                "SAP schließen empfohlen": st.column_config.CheckboxColumn("SAP schließen empf."),
                "Kommentar Teamleiter": st.column_config.TextColumn("Kommentar Teamleiter", width="large"),
                "Ampel": st.column_config.TextColumn("Ampel", width="small"),
            },
            key="mf_review_editor",
        )

        colb1, colb2, colb3 = st.columns(3)
        with colb1:
            if st.button("💾 Status speichern", type="primary"):
                to_save = edited[["Meldung", "Geprüft", "Geprüft am", "Kommentar Teamleiter",
                                   "SAP schließen empfohlen"]].copy()
                # "Geprüft am" automatisch setzen, wenn neu als geprüft markiert
                prev_log = review_log.set_index("Meldung") if not review_log.empty else pd.DataFrame()
                new_dates = []
                for _, r in to_save.iterrows():
                    if r["Geprüft"]:
                        prev_date = prev_log.loc[r["Meldung"], "Geprüft am"] if (
                            not prev_log.empty and r["Meldung"] in prev_log.index
                        ) else pd.NaT
                        new_dates.append(prev_date if pd.notna(prev_date) else heute)
                    else:
                        new_dates.append(pd.NaT)
                to_save["Geprüft am"] = new_dates
                combined = pd.concat([review_log, to_save], ignore_index=True)
                save_review_log(combined)
                st.success("Status gespeichert ✅ (persistiert in mf_review_log.csv)")
                st.rerun()
        with colb2:
            st.download_button(
                "⬇️ Review-Log sichern (CSV)",
                data=load_review_log().to_csv(index=False).encode("utf-8-sig"),
                file_name="mf_review_log.csv",
                mime="text/csv",
                help="Backup, falls das Dashboard neu deployed wird und den Status verliert.",
            )
        with colb3:
            restore_file = st.file_uploader("⬆️ Review-Log wiederherstellen", type=["csv"], key="restore_log")
            if restore_file is not None:
                restored = pd.read_csv(restore_file, dtype={"Meldung": str})
                save_review_log(restored)
                st.success("Review-Log wiederhergestellt ✅")
                st.rerun()

        st.caption(
            "Hinweis: Ein 'Geprüft am'-Datum wird automatisch gesetzt, sobald eine Meldung als "
            "geprüft markiert wird. Da SAP kein Änderungsdatum liefert, zeigt 'Tage seit Prüfung' "
            "die Zeit seit der letzten Teamleiter-Kontrolle (nicht seit der letzten SAP-Buchung)."
        )

        # -- 5) SAP-Schließliste exportieren ------------------------------------
        st.markdown("#### 📤 SAP-Schließliste exportieren")
        schliess_kandidaten = edited[edited["SAP schließen empfohlen"]]
        if schliess_kandidaten.empty:
            st.info("Noch keine Meldungen zum Schließen markiert (Häkchen 'SAP schließen empf.' setzen).")
        else:
            export_cols = [c for c in ["Auftrag", "Meldung", "Equipment", "Techn. Platz",
                                        "Tage offen", "Kommentar Teamleiter"] if c in schliess_kandidaten.columns]
            export_df = schliess_kandidaten[export_cols]
            st.dataframe(export_df, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ SAP-Schließliste herunterladen (CSV)",
                data=export_df.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"SAP_Schliessliste_{heute.date()}.csv",
                mime="text/csv",
            )

        st.markdown("---")

        # -- 6) Backlog-Trend: wann sind die offenen kritischen Faelle entstanden --
        st.markdown("#### 📈 Backlog-Verlauf – wann sind die aktuell offenen MF-Meldungen entstanden?")
        trend_mf = mf_kritisch.copy()
        trend_mf["Woche entstanden"] = trend_mf["Angelegt am"].dt.to_period("W").apply(lambda p: p.start_time)
        trend_g = trend_mf.groupby("Woche entstanden").size().reset_index(name="Anzahl noch offen")
        fig_trend = px.bar(
            trend_g, x="Woche entstanden", y="Anzahl noch offen",
            title="Aktuell noch offene kritische MF-Meldungen, gruppiert nach Entstehungswoche",
        )
        fig_trend.update_traces(marker_color="crimson")
        st.plotly_chart(fig_trend, use_container_width=True)
        st.caption(
            "Je weiter die Balken links liegen, desto älter ist der Rückstau – "
            "das zeigt, ob es sich um ein neues oder ein chronisches Problem handelt."
        )

        # -- 7) Aging-Histogramm ueber ALLE offenen MF-Meldungen ------------------
        st.markdown("#### 📊 Altersverteilung aller offenen MF-Meldungen")
        bins = [-1, 5, 10, 20, 30, 10_000]
        labels = ["0–5 Tage", "6–10 Tage", "11–20 Tage", "21–30 Tage", "> 30 Tage"]
        df_mf_offen["Alterskategorie"] = pd.cut(df_mf_offen["Tage offen"], bins=bins, labels=labels)
        age_g = df_mf_offen["Alterskategorie"].value_counts().reindex(labels).reset_index()
        age_g.columns = ["Alterskategorie", "Anzahl"]
        fig_age = px.bar(
            age_g, x="Alterskategorie", y="Anzahl", color="Alterskategorie",
            color_discrete_sequence=["#2ecc71", "#f1c40f", "#f39c12", "#e67e22", "#e74c3c"],
            title="Wie viele offene MF-Meldungen liegen in welcher Alterskategorie?",
        )
        fig_age.update_layout(showlegend=False)
        st.plotly_chart(fig_age, use_container_width=True)

        # -- 8) Farbige Gesamtuebersicht (read-only) als zusaetzliche Sicht -------
        with st.expander("📋 Farbige Gesamtansicht aller kritischen MF-Meldungen (read-only)"):
            show_cols2 = [c for c in ["Ampel", "Meldung", "Equipment", "Techn. Platz",
                                       "Bezeichnung", "Verantw.ArbPl.", "Angelegt am",
                                       "Tage offen", "Score", "Melder"] if c in mf_kritisch.columns]
            st.dataframe(
                style_tage_offen(mf_kritisch[show_cols2].sort_values("Tage offen", ascending=False)),
                use_container_width=True, hide_index=True,
            )

        # -- 9) Detail-Drilldown fuer eine einzelne Meldung ------------------------
        with st.expander("🔍 Details zu einer einzelnen Meldung anzeigen"):
            meldung_sel = st.selectbox("Meldung wählen", mf_kritisch["Meldung"].tolist())
            row_sel = mf_kritisch[mf_kritisch["Meldung"] == meldung_sel].iloc[0]
            for feld in row_sel.index:
                if pd.notna(row_sel[feld]) and feld not in ("Ampel", "Score"):
                    st.write(f"**{feld}:** {row_sel[feld]}")

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

st.divider()

# ----------------------------------------------------------------------------
# Werkstättenleistung – Optimierungspotenziale
# ----------------------------------------------------------------------------
st.subheader("🚀 Werkstättenleistung – Optimierungspotenziale")
st.caption(
    "Kennzahlen zur Identifikation von Kapazitätsengpässen, Wiederholungsfällen und "
    "Hebeln zur Steigerung der Werkstättenleistung – unabhängig von den Kostenbetrachtungen oben."
)

# -- A) Durchlaufzeit je Techn. Platz (nur abgeschlossene MF-Reparaturen) -----
mf_abgeschlossen = df_reparatur[~df_reparatur["Ist_offen"]].copy()
mf_abgeschlossen["Durchlaufzeit"] = (
    mf_abgeschlossen["Abschlußdatum"] - mf_abgeschlossen["Angelegt am"]
).dt.days

if mf_abgeschlossen.empty:
    st.info("Keine abgeschlossenen MF-Reparaturen im gewählten Zeitraum – Durchlaufzeit-Analyse nicht möglich.")
    dl_tp = pd.DataFrame(columns=["Techn. Platz", "Durchlaufzeit"])
else:
    dl_tp = (
        mf_abgeschlossen.groupby("Techn. Platz")["Durchlaufzeit"]
        .mean().round(1).reset_index()
        .sort_values("Durchlaufzeit", ascending=False)
    )
    fig_dl = px.bar(
        dl_tp.head(15), x="Techn. Platz", y="Durchlaufzeit",
        title=f"Ø Durchlaufzeit MF-Reparaturen je Techn. Platz (Tage) – Ziel: {ZIEL_DURCHLAUFZEIT_TAGE} Tage",
        color="Durchlaufzeit", color_continuous_scale=["#2ecc71", "#f1c40f", "#e74c3c"],
    )
    fig_dl.add_hline(
        y=ZIEL_DURCHLAUFZEIT_TAGE, line_dash="dash", line_color="red",
        annotation_text=f"Ziel: {ZIEL_DURCHLAUFZEIT_TAGE} Tage", annotation_position="top left",
    )
    st.plotly_chart(fig_dl, use_container_width=True)
    st.caption(
        "Werkstätten deutlich über der Ziellinie zeigen strukturelle Engpässe "
        "(Personal, Ersatzteile, Priorisierung) – guter Startpunkt für Prozessverbesserung."
    )

st.markdown("---")

# -- B) Wiederholungsquote (Indikator für Erstlösungsqualität) ----------------
col_wdh, col_verh = st.columns(2)

with col_wdh:
    st.markdown("**🔁 Wiederholungsquote je Techn. Platz**")
    if "Ursprungsmeldung" in df_aktiv.columns and "Folgemeldung" in df_aktiv.columns:
        df_rep = df_aktiv.copy()
        df_rep["Ist_Wiederholung"] = (
            df_rep["Ursprungsmeldung"].notna()
            & df_rep["Folgemeldung"].notna()
            & (df_rep["Ursprungsmeldung"] != df_rep["Folgemeldung"])
        )
        wdh_tp = (
            df_rep.groupby("Techn. Platz")
            .agg(Anzahl=("Meldung", "count"), Wiederholungen=("Ist_Wiederholung", "sum"))
            .reset_index()
        )
        wdh_tp = wdh_tp[wdh_tp["Anzahl"] >= 3]  # kleine Fallzahlen ausblenden
        wdh_tp["Quote_%"] = (wdh_tp["Wiederholungen"] / wdh_tp["Anzahl"] * 100).round(1)
        wdh_tp = wdh_tp.sort_values("Quote_%", ascending=False).head(10)
        if wdh_tp.empty:
            st.info("Nicht genügend Daten für eine belastbare Wiederholungsquote.")
        else:
            fig_wdh = px.bar(
                wdh_tp, x="Techn. Platz", y="Quote_%",
                title="Anteil Folge-/Wiederholungsmeldungen (%)",
                color="Quote_%", color_continuous_scale=["#2ecc71", "#e74c3c"],
            )
            st.plotly_chart(fig_wdh, use_container_width=True)
        st.caption(
            "Hohe Werte deuten auf unvollständige Erstreparaturen hin (Equipment kommt "
            "wiederholt zurück) – ein zentraler Hebel, um Werkstättenkapazität freizuspielen."
        )
    else:
        st.info("Spalten 'Ursprungsmeldung'/'Folgemeldung' nicht in der Datei vorhanden.")

with col_verh:
    st.markdown("**⚖️ Wartung (geplant) vs. Reparatur (ungeplant) je Techn. Platz**")
    verh_tp = (
        df_aktiv[df_aktiv["Kategorie"].isin(["Wartung", "Reparatur"])]
        .groupby(["Techn. Platz", "Kategorie"]).size().reset_index(name="Anzahl")
    )
    if verh_tp.empty:
        st.info("Keine Daten für Wartungs-/Reparaturverhältnis verfügbar.")
    else:
        pivot = verh_tp.pivot(index="Techn. Platz", columns="Kategorie", values="Anzahl").fillna(0)
        pivot["Summe"] = pivot.sum(axis=1)
        pivot = pivot.sort_values("Summe", ascending=False).head(10)
        pivot_pct = pivot[["Wartung", "Reparatur"]].div(pivot["Summe"], axis=0) * 100
        pivot_pct = pivot_pct.reset_index()
        fig_verh = px.bar(
            pivot_pct, x="Techn. Platz", y=["Wartung", "Reparatur"],
            barmode="stack", title="Anteil geplant (Wartung) vs. ungeplant (Reparatur), in %",
            color_discrete_map={"Wartung": "#2ecc71", "Reparatur": "#e74c3c"},
        )
        fig_verh.update_yaxes(title="Anteil (%)")
        st.plotly_chart(fig_verh, use_container_width=True)
        st.caption(
            "Ein hoher Reparatur-Anteil (rot) je Techn. Platz ist ein Signal für zu wenig "
            "präventive Wartung – mehr geplante Wartung reduziert ungeplante Ausfälle."
        )

st.markdown("---")

# -- C) Auslastung: Ist-Stunden je Techn. Platz/Monat vs. Referenzwert --------
st.markdown("**⏱️ Kapazitätsauslastung je Techn. Platz (Ø Ist-Stunden/Monat)**")
df2_ausl = df2.copy()
df2_ausl["Monat"] = df2_ausl["Eckstarttermin"].dt.to_period("M").astype(str)
ausl = (
    df2_ausl.groupby(["Techn. Platz", "Monat"])["Istarbeit Summe"].sum().reset_index()
)
ausl_avg = (
    ausl.groupby("Techn. Platz")["Istarbeit Summe"].mean().round(1).reset_index()
    .sort_values("Istarbeit Summe", ascending=False)
)
if ausl_avg.empty:
    st.info("Keine Auslastungsdaten verfügbar.")
else:
    fig_ausl = px.bar(
        ausl_avg.head(15), x="Techn. Platz", y="Istarbeit Summe",
        title=f"Ø Ist-Stunden/Monat je Techn. Platz – Referenzwert: {ZIEL_AUSLASTUNG_STUNDEN_MONAT} h",
    )
    fig_ausl.add_hline(
        y=ZIEL_AUSLASTUNG_STUNDEN_MONAT, line_dash="dash", line_color="orange",
        annotation_text=f"Referenz: {ZIEL_AUSLASTUNG_STUNDEN_MONAT} h/Monat", annotation_position="top left",
    )
    st.plotly_chart(fig_ausl, use_container_width=True)
    st.caption(
        "Werkstätten deutlich **unter** dem Referenzwert haben freie Kapazität für zusätzliche "
        "Aufträge; Werkstätten deutlich **darüber** sind ein Risiko für Verzögerungen und Burnout "
        "– hier ggf. Personal/Aufgaben umverteilen."
    )

st.markdown("---")

# -- D) Automatisierte Handlungsempfehlungen ----------------------------------
st.markdown("#### 💡 Automatisch abgeleitete Optimierungsempfehlungen")
empfehlungen = []

if not dl_tp.empty:
    schlechtester_dl = dl_tp.iloc[0]
    if schlechtester_dl["Durchlaufzeit"] > ZIEL_DURCHLAUFZEIT_TAGE:
        empfehlungen.append(
            f"🔴 **{schlechtester_dl['Techn. Platz']}** hat die längste Ø Durchlaufzeit "
            f"({schlechtester_dl['Durchlaufzeit']:.0f} Tage, Ziel: {ZIEL_DURCHLAUFZEIT_TAGE}). "
            "Prozess/Personal/Ersatzteilverfügbarkeit dort prüfen."
        )

if "Ursprungsmeldung" in df_aktiv.columns and "Folgemeldung" in df_aktiv.columns and not wdh_tp.empty:
    schlechtester_wdh = wdh_tp.iloc[0]
    empfehlungen.append(
        f"🔁 **{schlechtester_wdh['Techn. Platz']}** hat die höchste Wiederholungsquote "
        f"({schlechtester_wdh['Quote_%']:.0f}%). Erstlösungsqualität und Diagnoseprozess dort verbessern."
    )

if not verh_tp.empty and "pivot_pct" in dir():
    if "Reparatur" in pivot_pct.columns:
        schlechtester_verh = pivot_pct.sort_values("Reparatur", ascending=False).iloc[0]
        if schlechtester_verh["Reparatur"] > 60:
            empfehlungen.append(
                f"⚖️ **{schlechtester_verh['Techn. Platz']}** hat einen sehr hohen ungeplanten "
                f"Reparaturanteil ({schlechtester_verh['Reparatur']:.0f}%). Präventive Wartungsintervalle "
                "dort überprüfen/verdichten."
            )

if not ausl_avg.empty:
    unterausgelastet = ausl_avg[ausl_avg["Istarbeit Summe"] < ZIEL_AUSLASTUNG_STUNDEN_MONAT * 0.6]
    ueberausgelastet = ausl_avg[ausl_avg["Istarbeit Summe"] > ZIEL_AUSLASTUNG_STUNDEN_MONAT * 1.2]
    if not ueberausgelastet.empty and not unterausgelastet.empty:
        empfehlungen.append(
            f"⏱️ **{ueberausgelastet.iloc[0]['Techn. Platz']}** ist deutlich überausgelastet, während "
            f"**{unterausgelastet.iloc[-1]['Techn. Platz']}** freie Kapazität hat – Umverteilung von "
            "Aufträgen/Personal prüfen."
        )
    elif not ueberausgelastet.empty:
        empfehlungen.append(
            f"⏱️ **{ueberausgelastet.iloc[0]['Techn. Platz']}** ist deutlich überausgelastet "
            f"({ueberausgelastet.iloc[0]['Istarbeit Summe']:.0f} h/Monat) – zusätzliche Kapazität einplanen."
        )

if empfehlungen:
    for e in empfehlungen:
        st.markdown(f"- {e}")
else:
    st.success("Keine akuten Optimierungshinweise auf Basis der aktuellen Daten erkennbar 🎉")

st.caption(
    "Hinweis: MM-Meldungen sind aus allen Auswertungen ausgeschlossen. "
    "Als 'offen' gilt eine Meldung, solange kein Abschlußdatum eingetragen ist."
)
