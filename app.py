import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re

st.set_page_config(
    page_title="Cruscotto Dotazioni Organiche",
    layout="wide"
)

# ----------------------------
# Helpers
# ----------------------------
def to_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)

def simplify_qualifica(q1: str, reparto: str):
    if pd.isna(q1):
        return None
    s = str(q1).lower()
    rep = str(reparto).lower() if reparto is not None else ""

    if "infermiere" in s:
        return "INFERMERE"
    if "socio sanit" in s:
        return "OSS"
    if "tec. san radiologia" in s:
        return "TSRM"
    if "tec. san laboratorio" in s:
        return "TSLB"
    if "fisioterap" in s:
        return "FISOTERAP"
    if ("operatore tecnico" in s) or ("operatore tecn. special" in s):
        # in AAT è autista
        if "aat" in rep:
            return "AUTISTI"
        return "OSS/OT"

    return str(q1)

def service_from_row(cdr, reparto):
    s = f"{cdr} {reparto}".lower()

    if "endoscopia digestiva" in s:
        return "ENDOSCOPIA DIGESTIVA"
    if "chirurgia generale" in s or "ortopedia" in s:
        return "AREA CHIRURGICA + sala gessi"
    if "pediatria" in s:
        return "PEDIATRIA"
    if "pronto soccorso" in s:
        return "PRONTO SOCCORSO + AAT 118"
    if "aat118" in s or "aat - 118" in s or "aat poop" in s:
        return "AAT - 118"
    if "anestesia" in s and "rianimazione" in s:
        return "SALA OPERATORIA"
    if "area intensiva" in s:
        return "TERAPIA INTENSIVA"
    if "medicina generale" in s:
        return "MEDICINA INTERNA + DH"
    if "cardiologia" in s:
        return "CARDIOLOGIA - UTIC - AMBULAT"
    if "radiologia" in s:
        return "RADIOLOGIA"
    if "laboratorio analisi" in s:
        return "LABORATORIO"
    if "riabilitazione" in s:
        return "RIABILITAZIONE"
    if "ds04-dapss" in s or "daps" in s:
        return "POLIAMBULATORI + TRASPORTI"
    if "d701-psichiatria adolescenti" in s:
        return "PSICHIATRIA ADOLESCENTI - SPDC"
    if re.search(r"\bd702-psichiatria\b", s):
        return "COMUNITA' - CPS -CD"

    return None


@st.cache_data
def load_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=0)
    return df

def build_analisi_dotazioni(df_raw: pd.DataFrame, only_in_force=True):
    df = df_raw.copy()

    # numeric casting
    num_cols = [c for c in df.columns if c not in [
        "STABILIMENTO", "REPARTO", "DESC. RUOLO", "QUALIFICA", "DESC. AREA",
        "DESC. POSIZIONE", "COGNOME", "NOME", "TIPO DI PART-TIME",
        "DATA DAL", "DATA AL", "DATA ASSUNZIONE", "DATA CESSAZIONE",
        "CDR_DESC", "DESC. DIP.", "CODICE FISCALE"
    ]]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="ignore")

    # derived fields
    df["FTE"] = to_num(df["% PART-TIME"]).replace(0, 100) / 100

    df["FERIE_MAT_2025"] = to_num(df["FERIE"]) + to_num(df["FERIE RX"])
    df["FERIE_FRUITE_2025"] = to_num(df["FERIE GODUTE TOTALE"]) + to_num(df["FERIE GODUTE RX"])
    df["FERIE_RES_0101"] = to_num(df["FERIE RES."]) + to_num(df["FERIE RX RES."]) + to_num(df["FERIE AP RES."])

    df["ASSENZE_MAL104_ECC"] = (
        to_num(df["MALATTIA"]) +
        to_num(df["MALATTIA FIGLIO"]) +
        to_num(df["LEGGE 104"]) +
        to_num(df["PERMESSI"]) +
        to_num(df["AGGIOR."]) +
        to_num(df["INF./MAL.SERV"]) +
        to_num(df["CAR.PUBBLICA"]) +
        to_num(df.get("INFORTUNIO COVID", 0)) +
        to_num(df.get("MALATTIA COVID", 0))
    )

    df["ASP_GRAV_PUER_DIST"] = to_num(df["RECUPERO"]) + to_num(df["MISSIONE SOLO SERVIZIO"])

    df["STRAORD_REC"] = to_num(df["ORE DA RECUP. PROG."])
    df["STRAORD_PD"] = to_num(df["STR. PD. PROG."])
    df["STRAORD_PAG"] = to_num(df["STR. PROG."])
    df["FEST_PAG"] = to_num(df["FEST. INFRASETT. A PAGAMENTO"])
    df["FEST_REC"] = to_num(df["FEST. INFRASETT. A RECUPERO"])

    # Mapping servizio + qualifica
    df["QUALIFICA_S"] = [
        simplify_qualifica(q1, rep)
        for q1, rep in zip(df["QUALIFICA.1"], df["REPARTO"])
    ]
    df["SERVIZIO"] = [
        service_from_row(cdr, rep)
        for cdr, rep in zip(df["CDR_DESC"], df["REPARTO"])
    ]

    target_qual = {"INFERMERE", "OSS", "AUTISTI", "TSRM", "TSLB", "FISOTERAP", "OSS/OT"}
    df_scope = df[df["SERVIZIO"].notna() & df["QUALIFICA_S"].isin(target_qual)].copy()

    if only_in_force:
        max_data = df_scope["DATA AL"].max()
        df_scope = df_scope[df_scope["DATA AL"] == max_data].copy()

    # regola: OSS/OT viene mostrato come "OSS" in quasi tutti i servizi
    df_scope["QUALIFICA_OUT"] = df_scope["QUALIFICA_S"]
    df_scope.loc[
        (df_scope["SERVIZIO"] != "POLIAMBULATORI + TRASPORTI") &
        (df_scope["QUALIFICA_OUT"] == "OSS/OT"),
        "QUALIFICA_OUT"
    ] = "OSS"

    agg = df_scope.groupby(["SERVIZIO", "QUALIFICA_OUT"]).agg(
        OPERATORI=("MATRICOLA", "nunique"),
        FTE=("FTE", "sum"),
        ST_REC=("STRAORD_REC", "sum"),
        ST_PD=("STRAORD_PD", "sum"),
        ST_PAG=("STRAORD_PAG", "sum"),
        FEST_PAG=("FEST_PAG", "sum"),
        FEST_REC=("FEST_REC", "sum"),
        FERIE_MAT=("FERIE_MAT_2025", "sum"),
        FERIE_FRUITE=("FERIE_FRUITE_2025", "sum"),
        FERIE_RES=("FERIE_RES_0101", "sum"),
        ASSENZE=("ASSENZE_MAL104_ECC", "sum"),
        ASP=("ASP_GRAV_PUER_DIST", "sum"),
    ).reset_index()

    agg["MEDIA_PROC"] = agg["FERIE_FRUITE"] / agg["OPERATORI"]
    agg["PREST_AGGIUNTIVE"] = 0.0  # placeholder (da definire meglio)

    # Colonne finali come template ANALISI_DOTAZIONI
    out = agg.rename(columns={
        "SERVIZIO": "UUOO/SERVIZIO",
        "QUALIFICA_OUT": "QUALIFICA",
        "FTE": "N° FTE",
        "ST_REC": "st Recupero",
        "ST_PD": "st PD pagato",
        "ST_PAG": "st Pagato",
        "FEST_PAG": "Festivo pagato",
        "FEST_REC": "Festivo recupero",
        "FERIE_MAT": "Ferie maturate 2025",
        "FERIE_FRUITE": "Ferie fruite 2025",
        "MEDIA_PROC": "Media procapite",
        "FERIE_RES": "Residue al 01/01/2026",
        "ASSENZE": "Assenze mal/104/ecc (ore)",
        "ASP": "Asp/grav/puer/dist (ore)",
        "PREST_AGGIUNTIVE": "Prestazioni aggiuntive (ore)"
    })

    return out, df_scope


# ----------------------------
# UI
# ----------------------------
st.title("🩺 Cruscotto Dotazioni Organiche – Casalmaggiore 2025")

with st.sidebar:
    st.header("📤 Caricamento dati")
    file = st.file_uploader("Carica il file PROSPETTO PERSONALE...", type=["xlsx"])

    only_in_force = st.toggle("Solo in forza a fine periodo (DATA AL max)", value=True)
    st.caption("Se disattivi: includi anche cessati/trasferiti nel corso dell’anno.")

if not file:
    st.info("👈 Carica il file Excel per iniziare.")
    st.stop()

df_raw = load_excel(file)
analisi, df_scope = build_analisi_dotazioni(df_raw, only_in_force=only_in_force)

# ---- Filtri (dip/struttura/servizi)
with st.sidebar:
    st.divider()
    st.header("🎛️ Filtri")

    dip_opts = sorted(df_raw["DESC. DIP."].dropna().unique())
    stab_opts = sorted(df_raw["STABILIMENTO"].dropna().unique())
    cdr_opts = sorted(df_raw["CDR_DESC"].dropna().unique())

    dip_sel = st.multiselect("Dipartimenti", dip_opts, default=dip_opts)
    stab_sel = st.multiselect("Stabilimenti", stab_opts, default=stab_opts)
    cdr_sel = st.multiselect("CDR (centri di costo)", cdr_opts, default=cdr_opts)

# applica filtri alla vista generale (df_raw)
df_f = df_raw.copy()
df_f = df_f[df_f["DESC. DIP."].isin(dip_sel)]
df_f = df_f[df_f["STABILIMENTO"].isin(stab_sel)]
df_f = df_f[df_f["CDR_DESC"].isin(cdr_sel)]

# ---- Layout
tab1, tab2 = st.tabs(["📋 ANALISI_DOTAZIONI", "📊 Vista generale"])

with tab1:
    st.subheader("Tabella ricostruita tipo “ANALISI_DOTAZIONI”")
    st.dataframe(analisi, use_container_width=True)

    st.caption("ℹ️ 'Prestazioni aggiuntive' è impostata a 0 perché non ricavabile in modo univoco dal prospetto attuale.")

with tab2:
    st.subheader("KPI e distribuzioni (ospedale filtrato)")
    # costruisci dataset per boxplot
    df_tmp = df_f.copy()
    df_tmp["QUALIFICA_S"] = [
        simplify_qualifica(q1, rep)
        for q1, rep in zip(df_tmp["QUALIFICA.1"], df_tmp["REPARTO"])
    ]
    df_tmp["FTE"] = to_num(df_tmp["% PART-TIME"]).replace(0, 100)/100

    df_tmp["ASSENZE"] = (
        to_num(df_tmp["MALATTIA"]) +
        to_num(df_tmp["MALATTIA FIGLIO"]) +
        to_num(df_tmp["LEGGE 104"]) +
        to_num(df_tmp["PERMESSI"]) +
        to_num(df_tmp["AGGIOR."]) +
        to_num(df_tmp["INF./MAL.SERV"]) +
        to_num(df_tmp["CAR.PUBBLICA"])
    )

    df_tmp["STRAORD"] = (
        to_num(df_tmp["ORE DA RECUP. PROG."]) +
        to_num(df_tmp["STR. PD. PROG."]) +
        to_num(df_tmp["STR. PROG."])
    )

    # KPI
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Operatori (righe)", len(df_tmp))
    col2.metric("Matricole uniche", df_tmp["MATRICOLA"].nunique())
    col3.metric("FTE totali", round(df_tmp["FTE"].sum(), 2))
    col4.metric("Assenze totali (ore)", round(df_tmp["ASSENZE"].sum(), 1))

    st.divider()

    # BoxPlot assenze per qualifica
    fig1 = px.box(
        df_tmp[df_tmp["QUALIFICA_S"].isin(["INFERMERE", "OSS", "TSRM", "TSLB", "FISOTERAP"])],
        x="QUALIFICA_S", y="ASSENZE",
        points="all",
        title="Distribuzione Assenze (ore) per qualifica"
    )
    st.plotly_chart(fig1, use_container_width=True)

    # BoxPlot straordinari per dipartimento
    fig2 = px.box(
        df_tmp[df_tmp["QUALIFICA_S"].isin(["INFERMERE", "OSS"])],
        x="DESC. DIP.", y="STRAORD",
        points=False,
        title="Distribuzione Straordinari (ore) per Dipartimento (solo Infermieri/OSS)"
    )
    st.plotly_chart(fig2, use_container_width=True)
