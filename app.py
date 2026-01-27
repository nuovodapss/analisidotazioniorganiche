import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

# -----------------------------
# Config
# -----------------------------
st.set_page_config(page_title="Cruscotto Dotazioni Organiche", layout="wide")

# -----------------------------
# Utils: colonne e numeri
# -----------------------------
def norm(s: str) -> str:
    s = str(s)
    s = s.replace("\u00a0", " ")  # NBSP
    for h in ["–", "—", "−"]:
        s = s.replace(h, "-")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def find_col(df: pd.DataFrame, wanted, contains=True):
    """
    Restituisce il nome della prima colonna che matcha.
    wanted: lista di stringhe candidate (nomi attesi)
    contains: se True, match anche per 'contiene'
    """
    if isinstance(wanted, str):
        wanted = [wanted]

    cols = list(df.columns)
    cols_norm = [norm(c) for c in cols]
    wanted_norm = [norm(w) for w in wanted]

    # 1) match esatto normalizzato
    for w in wanted_norm:
        if w in cols_norm:
            return cols[cols_norm.index(w)]

    # 2) match "contiene"
    if contains:
        for i, c in enumerate(cols_norm):
            if any(w in c for w in wanted_norm):
                return cols[i]

    return None

def to_num_series(s: pd.Series) -> pd.Series:
    """Converte in numerico; gestisce virgola decimale; NaN -> 0."""
    if s is None:
        return None
    if pd.api.types.is_numeric_dtype(s):
        return s.fillna(0)

    s2 = s.astype(str).str.replace("\u00a0", " ", regex=False).str.strip()
    # gestisce "1,5"
    s2 = s2.str.replace(".", "", regex=False)  # rimuove separatore migliaia se presente in forma 1.234
    s2 = s2.str.replace(",", ".", regex=False)

    out = pd.to_numeric(s2, errors="coerce").fillna(0)
    return out

def z(df: pd.DataFrame) -> pd.Series:
    return pd.Series(0.0, index=df.index)

# -----------------------------
# Business: mappatura qualifica + servizio automatico
# -----------------------------
def simplify_qualifica(q: str, reparto: str) -> str | None:
    if q is None or (isinstance(q, float) and np.isnan(q)):
        return None

    s = str(q).lower()
    rep = str(reparto).lower() if reparto is not None else ""

    if "infermiere" in s:
        return "INFERMERE"
    if "socio sanit" in s or "operatore socio" in s:
        return "OSS"
    if "tec. san radiologia" in s or "tsrm" in s:
        return "TSRM"
    if "tec. san laboratorio" in s or "tslb" in s:
        return "TSLB"
    if "fisioterap" in s:
        return "FISOTERAP"
    if "operatore tecnico" in s or "operatore tecn. special" in s:
        if "aat" in rep:
            return "AUTISTI"
        return "OSS/OT"

    # fallback: lascia la stringa originale
    return str(q).strip()

def build_analisi(df_raw: pd.DataFrame, only_in_force: bool) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = df_raw.copy()

    # colonne chiave (robuste)
    c_pt = find_col(df, ["% PART-TIME", "% PART TIME", "PART-TIME", "PERC PART"], contains=True)
    c_rep = find_col(df, ["REPARTO"], contains=True)
    c_cdr = find_col(df, ["CDR_DESC", "CDR DESC", "CDR"], contains=True)
    c_dip = find_col(df, ["DESC. DIP.", "DESC DIP", "DIP"], contains=True)
    c_stab = find_col(df, ["STABILIMENTO"], contains=True)
    c_matr = find_col(df, ["MATRICOLA"], contains=True)
    c_data_al = find_col(df, ["DATA AL"], contains=True)

    # qualifica: spesso è duplicata (QUALIFICA + QUALIFICA.1)
    c_qual = find_col(df, ["QUALIFICA.1", "QUALIFICA"], contains=True)

    # colonne ore/assenze/ferie (se mancano => 0)
    def col_or_zero(names):
        c = find_col(df, names, contains=True)
        return to_num_series(df[c]) if c else z(df)

    ferie = col_or_zero(["FERIE"])
    ferie_rx = col_or_zero(["FERIE RX"])
    ferie_god_tot = col_or_zero(["FERIE GODUTE TOTALE"])
    ferie_god_rx = col_or_zero(["FERIE GODUTE RX"])
    ferie_res = col_or_zero(["FERIE RES."])
    ferie_rx_res = col_or_zero(["FERIE RX RES."])
    ferie_ap_res = col_or_zero(["FERIE AP RES."])

    malattia = col_or_zero(["MALATTIA"])
    mal_figlio = col_or_zero(["MALATTIA FIGLIO"])
    legge104 = col_or_zero(["LEGGE 104"])
    permessi = col_or_zero(["PERMESSI"])
    aggior = col_or_zero(["AGGIOR."])
    inf_mal_serv = col_or_zero(["INF./MAL.SERV"])
    car_pub = col_or_zero(["CAR.PUBBLICA"])
    infort_covid = col_or_zero(["INFORTUNIO COVID"])
    mal_covid = col_or_zero(["MALATTIA COVID"])

    rec = col_or_zero(["RECUPERO"])
    missione = col_or_zero(["MISSIONE SOLO SERVIZIO"])

    st_rec = col_or_zero(["ORE DA RECUP. PROG."])
    st_pd = col_or_zero(["STR. PD. PROG."])
    st_pag = col_or_zero(["STR. PROG."])

    fest_pag = col_or_zero(["FEST. INFRASETT. A PAGAMENTO"])
    fest_rec = col_or_zero(["FEST. INFRASETT. A RECUPERO"])

    # FTE
    if c_pt:
        pt = to_num_series(df[c_pt])
        pt = pt.where(pt > 0, 100)  # 0 => full time
        df["FTE"] = pt / 100
    else:
        df["FTE"] = 1.0

    # ferie/assenze derivate
    df["FERIE_MAT_2025"] = ferie + ferie_rx
    df["FERIE_FRUITE_2025"] = ferie_god_tot + ferie_god_rx
    df["FERIE_RES_0101"] = ferie_res + ferie_rx_res + ferie_ap_res

    df["ASSENZE_MAL104_ECC"] = (
        malattia + mal_figlio + legge104 + permessi + aggior + inf_mal_serv + car_pub + infort_covid + mal_covid
    )
    df["ASP_GRAV_PUER_DIST"] = rec + missione

    df["STRAORD_REC"] = st_rec
    df["STRAORD_PD"] = st_pd
    df["STRAORD_PAG"] = st_pag
    df["FEST_PAG"] = fest_pag
    df["FEST_REC"] = fest_rec

    # servizio automatico e scalabile: CDR_DESC (fallback REPARTO)
    if c_cdr:
        df["SERVIZIO"] = df[c_cdr].astype(str).str.strip()
    else:
        df["SERVIZIO"] = ""

    if c_rep:
        rep_str = df[c_rep].astype(str).str.strip()
        df.loc[df["SERVIZIO"].isin(["", "nan", "None"]), "SERVIZIO"] = rep_str
        # se SERVIZIO proprio vuoto, metti reparto
        df.loc[df["SERVIZIO"].eq(""), "SERVIZIO"] = rep_str

    # qualifica semplificata
    rep_for_q = df[c_rep] if c_rep else pd.Series([""] * len(df), index=df.index)
    q_for_q = df[c_qual] if c_qual else pd.Series([None] * len(df), index=df.index)
    df["QUALIFICA_S"] = [
        simplify_qualifica(q, r) for q, r in zip(q_for_q.tolist(), rep_for_q.tolist())
    ]

    # filtra scope "core"
    target = {"INFERMERE", "OSS", "AUTISTI", "TSRM", "TSLB", "FISOTERAP", "OSS/OT"}
    df_scope = df[df["SERVIZIO"].notna() & df["QUALIFICA_S"].isin(target)].copy()

    # opzionale: solo in forza a fine periodo (DATA AL max)
    if only_in_force and c_data_al and len(df_scope) > 0:
        max_data = df_scope[c_data_al].max()
        df_scope = df_scope[df_scope[c_data_al] == max_data].copy()

    # OSS/OT -> OSS (se vuoi lasciarli distinti, commenta queste 2 righe)
    df_scope["QUALIFICA_OUT"] = df_scope["QUALIFICA_S"]
    df_scope.loc[df_scope["QUALIFICA_OUT"] == "OSS/OT", "QUALIFICA_OUT"] = "OSS"

    # aggregazione
    matr_col = c_matr if c_matr else None

    agg = df_scope.groupby(["SERVIZIO", "QUALIFICA_OUT"], dropna=False).agg(
        OPERATORI=(matr_col, "nunique") if matr_col else ("QUALIFICA_OUT", "size"),
        **{
            "N° FTE": ("FTE", "sum"),
            "st Recupero": ("STRAORD_REC", "sum"),
            "st PD pagato": ("STRAORD_PD", "sum"),
            "st Pagato": ("STRAORD_PAG", "sum"),
            "Festivo pagato": ("FEST_PAG", "sum"),
            "Festivo recupero": ("FEST_REC", "sum"),
            "Ferie maturate 2025": ("FERIE_MAT_2025", "sum"),
            "Ferie fruite 2025": ("FERIE_FRUITE_2025", "sum"),
            "Residue al 01/01/2026": ("FERIE_RES_0101", "sum"),
            "Assenze mal/104/ecc (ore)": ("ASSENZE_MAL104_ECC", "sum"),
            "Asp/grav/puer/dist (ore)": ("ASP_GRAV_PUER_DIST", "sum"),
        }
    ).reset_index()

    agg.rename(columns={"SERVIZIO": "UUOO/SERVIZIO", "QUALIFICA_OUT": "QUALIFICA"}, inplace=True)

    agg["Media procapite"] = np.where(
        agg["OPERATORI"] > 0, agg["Ferie fruite 2025"] / agg["OPERATORI"], 0.0
    )
    agg["Prestazioni aggiuntive (ore)"] = 0.0

    # riordino colonne (stile tabella target)
    ordered = [
        "UUOO/SERVIZIO", "QUALIFICA", "OPERATORI", "N° FTE",
        "st Recupero", "st PD pagato", "st Pagato",
        "Festivo pagato", "Festivo recupero",
        "Ferie maturate 2025", "Ferie fruite 2025", "Media procapite",
        "Residue al 01/01/2026",
        "Assenze mal/104/ecc (ore)", "Asp/grav/puer/dist (ore)",
        "Prestazioni aggiuntive (ore)",
    ]
    for c in ordered:
        if c not in agg.columns:
            agg[c] = 0.0

    agg = agg[ordered].sort_values(["UUOO/SERVIZIO", "QUALIFICA"]).reset_index(drop=True)
    return agg, df_scope

# -----------------------------
# UI
# -----------------------------
st.title("🩺 Cruscotto Dotazioni Organiche – Streamlit (Excel upload)")

with st.sidebar:
    st.header("📤 Caricamento")
    file = st.file_uploader("Carica PROSPETTO PERSONALE COMPARTO (xlsx)", type=["xlsx"])
    only_in_force = st.toggle("Solo in forza a fine periodo (DATA AL max)", value=True)

    st.divider()
    st.header("🎛️ Filtri (su dataset grezzo)")

# sempre visibile: mai pagina bianca
if not file:
    st.info("Carica un file Excel dalla sidebar per iniziare.")
    st.stop()

try:
    df_raw = pd.read_excel(file, sheet_name=0)
except Exception as e:
    st.error("Errore durante lettura Excel.")
    st.exception(e)
    st.stop()

# filtri robusti
c_dip = find_col(df_raw, ["DESC. DIP.", "DESC DIP", "DIP"], contains=True)
c_stab = find_col(df_raw, ["STABILIMENTO"], contains=True)
c_cdr = find_col(df_raw, ["CDR_DESC", "CDR DESC", "CDR"], contains=True)
c_rep = find_col(df_raw, ["REPARTO"], contains=True)

with st.sidebar:
    dip_opts = sorted(df_raw[c_dip].dropna().astype(str).unique()) if c_dip else []
    stab_opts = sorted(df_raw[c_stab].dropna().astype(str).unique()) if c_stab else []
    cdr_opts = sorted(df_raw[c_cdr].dropna().astype(str).unique()) if c_cdr else []
    rep_opts = sorted(df_raw[c_rep].dropna().astype(str).unique()) if c_rep else []

    dip_sel = st.multiselect("Dipartimento", dip_opts, default=dip_opts[:]) if dip_opts else []
    stab_sel = st.multiselect("Stabilimento", stab_opts, default=stab_opts[:]) if stab_opts else []
    cdr_sel = st.multiselect("CDR_DESC", cdr_opts, default=cdr_opts[:]) if cdr_opts else []
    rep_sel = st.multiselect("Reparto", rep_opts, default=rep_opts[:]) if rep_opts else []

df_f = df_raw.copy()
if c_dip and dip_sel:
    df_f = df_f[df_f[c_dip].astype(str).isin(dip_sel)]
if c_stab and stab_sel:
    df_f = df_f[df_f[c_stab].astype(str).isin(stab_sel)]
if c_cdr and cdr_sel:
    df_f = df_f[df_f[c_cdr].astype(str).isin(cdr_sel)]
if c_rep and rep_sel:
    df_f = df_f[df_f[c_rep].astype(str).isin(rep_sel)]

# debug colonne
with st.expander("🔎 Debug: colonne lette dal file (utile se qualcosa non torna)"):
    st.write(list(df_raw.columns))

tab1, tab2 = st.tabs(["📋 ANALISI_DOTAZIONI (ricostruita)", "📊 Vista generale (boxplot)"])

with tab1:
    st.subheader("Tabella ANALISI_DOTAZIONI (derivata dal prospetto)")
    try:
        analisi, df_scope = build_analisi(df_f, only_in_force=only_in_force)
        st.dataframe(analisi, use_container_width=True, height=520)

        # download
        csv = analisi.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Scarica CSV", data=csv, file_name="ANALISI_DOTAZIONI_ricostruita.csv", mime="text/csv")

        st.caption("Nota: 'Prestazioni aggiuntive' è impostata a 0 (da definire la fonte nel tracciato).")
    except Exception as e:
        st.error("Errore nella costruzione della tabella ANALISI_DOTAZIONI.")
        st.exception(e)

with tab2:
    st.subheader("Distribuzioni (dataset filtrato)")

    try:
        analisi, df_scope = build_analisi(df_f, only_in_force=only_in_force)

        # KPI
        c_matr = find_col(df_f, ["MATRICOLA"], contains=True)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Righe (grezzo filtrato)", f"{len(df_f):,}".replace(",", "."))
        col2.metric("Matricole uniche", f"{df_f[c_matr].nunique():,}".replace(",", ".") if c_matr else "n/d")
        col3.metric("FTE totali (scope)", round(float(df_scope["FTE"].sum()), 2) if "FTE" in df_scope.columns else "n/d")
        col4.metric("Assenze totali (ore) (scope)", round(float(df_scope["ASSENZE_MAL104_ECC"].sum()), 1) if "ASSENZE_MAL104_ECC" in df_scope.columns else "n/d")

        st.divider()

        # Boxplot assenze per qualifica
        if len(df_scope) > 0:
            fig1 = px.box(
                df_scope,
                x="QUALIFICA_S",
                y="ASSENZE_MAL104_ECC",
                points="all",
                title="BoxPlot Assenze (ore) per qualifica (scope)"
            )
            st.plotly_chart(fig1, use_container_width=True)

            # Boxplot straordinari per servizio (top 20 per numerosità)
            df_scope["STRAORD_TOT"] = df_scope["STRAORD_REC"] + df_scope["STRAORD_PD"] + df_scope["STRAORD_PAG"]
            top_serv = df_scope["SERVIZIO"].value_counts().head(20).index
            fig2 = px.box(
                df_scope[df_scope["SERVIZIO"].isin(top_serv)],
                x="SERVIZIO",
                y="STRAORD_TOT",
                points=False,
                title="BoxPlot Straordinari totali (ore) per SERVIZIO (Top 20 per numerosità)"
            )
            fig2.update_layout(xaxis_tickangle=45)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.warning("Nessun dato nello scope (qualifiche target) con i filtri attuali.")
    except Exception as e:
        st.error("Errore nella vista generale.")
        st.exception(e)

