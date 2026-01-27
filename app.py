# app.py
import io
import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Cruscotto Dotazioni Organiche", layout="wide")


# =========================
# UTIL: normalizzazione
# =========================
def norm(s: str) -> str:
    s = str(s).replace("\u00a0", " ").strip()
    for h in ["–", "—", "−"]:
        s = s.replace(h, "-")
    s = re.sub(r"\s+", " ", s)
    return s.lower()


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).replace("\u00a0", " ").strip() for c in df.columns]

    # rimuove colonne "Unnamed" se (quasi) vuote
    drop_cols = []
    for c in df.columns:
        if str(c).lower().startswith("unnamed"):
            col = df[c]
            nan_ratio = col.isna().mean()
            if nan_ratio >= 0.95:
                drop_cols.append(c)
    if drop_cols:
        df = df.drop(columns=drop_cols)

    df = df.dropna(how="all").reset_index(drop=True)
    return df


def find_col(df: pd.DataFrame, wanted, contains=True):
    """
    Trova una colonna in modo robusto.
    wanted: string oppure lista di possibili nomi.
    """
    if isinstance(wanted, str):
        wanted = [wanted]

    cols = list(df.columns)
    cols_norm = [norm(c) for c in cols]
    wanted_norm = [norm(w) for w in wanted]

    # match esatto
    for w in wanted_norm:
        if w in cols_norm:
            return cols[cols_norm.index(w)]

    # match "contiene"
    if contains:
        for i, c in enumerate(cols_norm):
            if any(w in c for w in wanted_norm):
                return cols[i]

    return None


def to_num_series(s: pd.Series) -> pd.Series:
    """
    Converte una Series in numerico gestendo:
    - virgola decimale
    - separatore migliaia
    - spazi/NBSP
    """
    if s is None:
        return None
    if pd.api.types.is_numeric_dtype(s):
        return s.fillna(0)

    x = s.astype(str).str.replace("\u00a0", " ", regex=False).str.strip()

    # Se contiene sia "." che "," assumiamo formato italiano: "."
migliaia, "," decimale
    mask_both = x.str.contains(r"\.", regex=True) &
x.str.contains(",", regex=False)
    x.loc[mask_both] = x.loc[mask_both].str.replace(".", "",
regex=False).str.replace(",", ".", regex=False)

    # Se contiene solo "," => decimale
    mask_comma = ~mask_both & x.str.contains(",", regex=False)
    x.loc[mask_comma] = x.loc[mask_comma].str.replace(",", ".", regex=False)

    # Se contiene spazi come migliaia (es. "1 234")
    x = x.str.replace(" ", "", regex=False)

    out = pd.to_numeric(x, errors="coerce").fillna(0)
    return out


def z(df: pd.DataFrame) -> pd.Series:
    return pd.Series(0.0, index=df.index)


# =========================
# LETTURA SMART EXCEL
# =========================
EXPECTED_COLS = {
    "ENTE", "COD. STAB", "STABILIMENTO", "COD.REP.", "REPARTO",
    "CONTRATTO", "RUOLO", "DESC. RUOLO", "PROFILO", "QUALIFICA",
    "MATRICOLA", "COGNOME", "NOME", "DATA DAL", "DATA AL",
    "TIPO DI PART-TIME", "% PART-TIME", "DATA ASSUNZIONE", "DATA CESSAZIONE",
    "ORE LAVORATE", "ORE TEORICHE", "ORE DA RECUP. PROG.", "STR. PD.
PROG.", "STR. PROG.",
    "FEST. INFRASETT. A PAGAMENTO", "FEST. INFRASETT. A RECUPERO",
    "FERIE", "FERIE RX", "FERIE GODUTE TOTALE", "FERIE GODUTE RX",
    "FERIE RES.", "FERIE RX RES.", "FERIE AP RES.",
    "PERMESSI", "MALATTIA", "MALATTIA FIGLIO", "LEGGE 104",
    "RECUPERO", "MISSIONE SOLO SERVIZIO",
    "CDR_DESC", "DESC. DIP."
}


def detect_best_sheet_and_header(xls: pd.ExcelFile, max_scan_rows=60):
    expected_norm = set(norm(x) for x in EXPECTED_COLS)
    best = {"sheet": None, "header_row": 0, "score": -1}

    for sheet in xls.sheet_names:
        try:
            preview = pd.read_excel(xls, sheet_name=sheet,
header=None, nrows=max_scan_rows)
        except Exception:
            continue

        # Valuta ogni riga: quante "colonne attese" contiene
        for r in range(min(len(preview), max_scan_rows)):
            row = preview.iloc[r].dropna().tolist()
            if not row:
                continue
            row_norm = set(norm(v) for v in row)
            score = len(row_norm & expected_norm)
            # Piccolo bonus se la riga sembra "larga" (molte celle non vuote)
            score = score * 10 + len(row)
            if score > best["score"]:
                best = {"sheet": sheet, "header_row": r, "score": score}

    return best


@st.cache_data
def load_excel_smart(file_bytes: bytes, sheet_override: str | None,
header_override: int | None):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    # Se override attivi, usa quelli
    if sheet_override is not None and header_override is not None:
        df = pd.read_excel(xls, sheet_name=sheet_override,
header=int(header_override))
        df = clean_columns(df)
        meta = {"sheet": sheet_override, "header_row":
int(header_override), "auto": False, "score": None}
        return df, meta, xls.sheet_names

    # Altrimenti: auto-detect
    best = detect_best_sheet_and_header(xls)
    df = pd.read_excel(xls, sheet_name=best["sheet"],
header=int(best["header_row"]))
    df = clean_columns(df)
    meta = {"sheet": best["sheet"], "header_row":
int(best["header_row"]), "auto": True, "score": best["score"]}
    return df, meta, xls.sheet_names


# =========================
# LOGICA: QUALIFICA / ANALISI
# =========================
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

    return str(q).strip()


def build_analisi(df_raw: pd.DataFrame, only_in_force: bool, core_only: bool):
    df = df_raw.copy()

    # colonne principali (robuste)
    c_pt = find_col(df, ["% PART-TIME", "% PART TIME", "PART-TIME",
"PERC PART"], contains=True)
    c_rep = find_col(df, ["REPARTO"], contains=True)
    c_cdr = find_col(df, ["CDR_DESC", "CDR DESC", "CDR"], contains=True)
    c_matr = find_col(df, ["MATRICOLA"], contains=True)
    c_data_al = find_col(df, ["DATA AL"], contains=True)

    # qualifica: prova prima QUALIFICA.1 poi QUALIFICA
    c_qual = find_col(df, ["QUALIFICA.1", "QUALIFICA"], contains=True)

    def col_or_zero(names):
        c = find_col(df, names, contains=True)
        return to_num_series(df[c]) if c else z(df)

    # ferie
    ferie = col_or_zero(["FERIE"])
    ferie_rx = col_or_zero(["FERIE RX"])
    ferie_god_tot = col_or_zero(["FERIE GODUTE TOTALE"])
    ferie_god_rx = col_or_zero(["FERIE GODUTE RX"])
    ferie_res = col_or_zero(["FERIE RES."])
    ferie_rx_res = col_or_zero(["FERIE RX RES."])
    ferie_ap_res = col_or_zero(["FERIE AP RES."])

    # assenze / permessi
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

    # straordinari / festivi
    st_rec = col_or_zero(["ORE DA RECUP. PROG."])
    st_pd = col_or_zero(["STR. PD. PROG."])
    st_pag = col_or_zero(["STR. PROG."])
    fest_pag = col_or_zero(["FEST. INFRASETT. A PAGAMENTO"])
    fest_rec = col_or_zero(["FEST. INFRASETT. A RECUPERO"])

    # FTE
    if c_pt:
        pt = to_num_series(df[c_pt])
        # Se sembra già in [0,1] (es. 0.8) usalo così; se in [0,100]
dividi per 100
        maxv = float(pt.max()) if len(pt) else 0.0
        if maxv <= 1.0:
            fte = pt
        else:
            pt = pt.where(pt > 0, 100)  # 0 => full time
            fte = pt / 100
        df["FTE"] = fte
    else:
        df["FTE"] = 1.0

    # derive
    df["FERIE_MAT_2025"] = ferie + ferie_rx
    df["FERIE_FRUITE_2025"] = ferie_god_tot + ferie_god_rx
    df["FERIE_RES_0101"] = ferie_res + ferie_rx_res + ferie_ap_res

    df["ASSENZE_MAL104_ECC"] = (
        malattia + mal_figlio + legge104 + permessi + aggior +
inf_mal_serv + car_pub + infort_covid + mal_covid
    )
    df["ASP_GRAV_PUER_DIST"] = rec + missione

    df["STRAORD_REC"] = st_rec
    df["STRAORD_PD"] = st_pd
    df["STRAORD_PAG"] = st_pag
    df["FEST_PAG"] = fest_pag
    df["FEST_REC"] = fest_rec

    # SERVIZIO: CDR_DESC (fallback REPARTO)
    if c_cdr:
        df["SERVIZIO"] = df[c_cdr].astype(str).str.strip()
    else:
        df["SERVIZIO"] = ""

    if c_rep:
        rep_str = df[c_rep].astype(str).str.strip()
        df.loc[df["SERVIZIO"].isin(["", "nan", "None"]), "SERVIZIO"] = rep_str
        df.loc[df["SERVIZIO"].eq(""), "SERVIZIO"] = rep_str

    # QUALIFICA_S
    rep_for_q = df[c_rep] if c_rep else pd.Series([""] * len(df),
index=df.index)
    q_for_q = df[c_qual] if c_qual else pd.Series([None] * len(df),
index=df.index)
    df["QUALIFICA_S"] = [simplify_qualifica(q, r) for q, r in
zip(q_for_q.tolist(), rep_for_q.tolist())]

    # opzionale: solo qualifiche core
    target = {"INFERMERE", "OSS", "AUTISTI", "TSRM", "TSLB",
"FISOTERAP", "OSS/OT"}
    df_scope = df[df["SERVIZIO"].notna()].copy()
    if core_only:
        df_scope = df_scope[df_scope["QUALIFICA_S"].isin(target)].copy()

    # opzionale: solo in forza a fine periodo
    if only_in_force and c_data_al and len(df_scope) > 0:
        try:
            max_data = df_scope[c_data_al].max()
            df_scope = df_scope[df_scope[c_data_al] == max_data].copy()
        except Exception:
            pass

    # OSS/OT -> OSS (se vuoi distinti, commenta)
    df_scope["QUALIFICA_OUT"] = df_scope["QUALIFICA_S"]
    df_scope.loc[df_scope["QUALIFICA_OUT"] == "OSS/OT", "QUALIFICA_OUT"] = "OSS"

    # aggregazione
    matr_col = c_matr if c_matr else None
    agg = df_scope.groupby(["SERVIZIO", "QUALIFICA_OUT"], dropna=False).agg(
        OPERATORI=(matr_col, "nunique") if matr_col else
("QUALIFICA_OUT", "size"),
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

    agg.rename(columns={"SERVIZIO": "UUOO/SERVIZIO", "QUALIFICA_OUT":
"QUALIFICA"}, inplace=True)
    agg["Media procapite"] = np.where(agg["OPERATORI"] > 0, agg["Ferie
fruite 2025"] / agg["OPERATORI"], 0.0)
    agg["Prestazioni aggiuntive (ore)"] = 0.0

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
    agg = agg[ordered].sort_values(["UUOO/SERVIZIO",
"QUALIFICA"]).reset_index(drop=True)

    return agg, df_scope


# =========================
# UI
# =========================
st.title("🩺 Cruscotto Dotazioni Organiche – Upload Excel")

with st.sidebar:
    st.header("📤 Caricamento")
    uploaded = st.file_uploader("Carica PROSPETTO PERSONALE COMPARTO
(xlsx)", type=["xlsx"])

    st.divider()
    st.header("⚙️ Opzioni")
    only_in_force = st.toggle("Solo in forza a fine periodo (DATA AL
max)", value=True)
    core_only = st.toggle("Solo qualifiche core
(Infermiere/OSS/TSRM/TSLB/Fisioterap/Autisti)", value=True)

if not uploaded:
    st.info("Carica un file Excel dalla sidebar per iniziare.")
    st.stop()

file_bytes = uploaded.getvalue()

# Override foglio/header (opzionale ma utilissimo)
with st.sidebar:
    st.divider()
    st.header("🧠 Lettura Excel (smart)")
    use_override = st.checkbox("Override manuale (foglio + riga
header)", value=False)

# Prima lettura (auto o placeholder per avere sheet_names)
try:
    df_raw, meta, sheet_names = load_excel_smart(file_bytes, None, None)
except Exception as e:
    st.error("Errore durante lettura Excel (auto).")
    st.exception(e)
    st.stop()

sheet_override = None
header_override = None

if use_override:
    with st.sidebar:
        sheet_override = st.selectbox("Foglio", sheet_names,
index=sheet_names.index(meta["sheet"]) if meta["sheet"] in sheet_names
else 0)
        header_override = st.number_input("Riga header (0 = prima
riga)", min_value=0, max_value=200, value=int(meta["header_row"]),
step=1)

    try:
        df_raw, meta, sheet_names = load_excel_smart(file_bytes,
sheet_override, int(header_override))
    except Exception as e:
        st.error("Errore durante lettura Excel (override).")
        st.exception(e)
        st.stop()

# Debug lettura
with st.expander("🔎 Debug lettura Excel (foglio + riga header + colonne)"):
    st.write(meta)
    st.write("Colonne lette:")
    st.write(list(df_raw.columns))
    st.dataframe(df_raw.head(15), use_container_width=True)

# Filtri su dataset grezzo
c_dip = find_col(df_raw, ["DESC. DIP.", "DESC DIP", "DIP"], contains=True)
c_stab = find_col(df_raw, ["STABILIMENTO"], contains=True)
c_cdr = find_col(df_raw, ["CDR_DESC", "CDR DESC", "CDR"], contains=True)
c_rep = find_col(df_raw, ["REPARTO"], contains=True)

with st.sidebar:
    st.divider()
    st.header("🎛️ Filtri (dataset grezzo)")
    dip_opts = sorted(df_raw[c_dip].dropna().astype(str).unique()) if
c_dip else []
    stab_opts = sorted(df_raw[c_stab].dropna().astype(str).unique())
if c_stab else []
    cdr_opts = sorted(df_raw[c_cdr].dropna().astype(str).unique()) if
c_cdr else []
    rep_opts = sorted(df_raw[c_rep].dropna().astype(str).unique()) if
c_rep else []

    dip_sel = st.multiselect("Dipartimento", dip_opts,
default=dip_opts) if dip_opts else []
    stab_sel = st.multiselect("Stabilimento", stab_opts,
default=stab_opts) if stab_opts else []
    cdr_sel = st.multiselect("CDR_DESC", cdr_opts, default=cdr_opts)
if cdr_opts else []
    rep_sel = st.multiselect("Reparto", rep_opts, default=rep_opts) if
rep_opts else []

df_f = df_raw.copy()
if c_dip and dip_sel:
    df_f = df_f[df_f[c_dip].astype(str).isin(dip_sel)]
if c_stab and stab_sel:
    df_f = df_f[df_f[c_stab].astype(str).isin(stab_sel)]
if c_cdr and cdr_sel:
    df_f = df_f[df_f[c_cdr].astype(str).isin(cdr_sel)]
if c_rep and rep_sel:
    df_f = df_f[df_f[c_rep].astype(str).isin(rep_sel)]

tab1, tab2 = st.tabs(["📋 ANALISI_DOTAZIONI (ricostruita)", "📊 Vista
generale (boxplot)"])

with tab1:
    st.subheader("Tabella ANALISI_DOTAZIONI (derivata dal prospetto)")
    try:
        analisi, df_scope = build_analisi(df_f,
only_in_force=only_in_force, core_only=core_only)
        st.dataframe(analisi, use_container_width=True, height=520)

        # download CSV
        st.download_button(
            "⬇️ Scarica CSV",
            data=analisi.to_csv(index=False).encode("utf-8"),
            file_name="ANALISI_DOTAZIONI_ricostruita.csv",
            mime="text/csv",
        )

        # download Excel
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            analisi.to_excel(writer, index=False,
sheet_name="ANALISI_DOTAZIONI")
        st.download_button(
            "⬇️ Scarica Excel",
            data=buf.getvalue(),
            file_name="ANALISI_DOTAZIONI_ricostruita.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.caption("Nota: 'Prestazioni aggiuntive' è impostata a 0 (da
definire la fonte nel tracciato).")

        if len(df_scope) == 0:
            st.warning("Scope vuoto: prova a disattivare 'Solo
qualifiche core' oppure verifica che l’header sia corretto.")
    except Exception as e:
        st.error("Errore nella costruzione della tabella.")
        st.exception(e)

with tab2:
    st.subheader("Distribuzioni (dataset filtrato)")
    try:
        analisi, df_scope = build_analisi(df_f,
only_in_force=only_in_force, core_only=core_only)

        c_matr = find_col(df_f, ["MATRICOLA"], contains=True)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Righe (grezzo filtrato)",
f"{len(df_f):,}".replace(",", "."))
        col2.metric("Matricole uniche",
f"{df_f[c_matr].nunique():,}".replace(",", ".") if c_matr else "n/d")
        col3.metric("FTE totali (scope)",
round(float(df_scope["FTE"].sum()), 2) if "FTE" in df_scope.columns
else "n/d")
        col4.metric("Assenze totali (ore) (scope)",
round(float(df_scope["ASSENZE_MAL104_ECC"].sum()), 1) if
"ASSENZE_MAL104_ECC" in df_scope.columns else "n/d")

        st.divider()

        if len(df_scope) == 0:
            st.warning("Nessun dato nello scope con i filtri/opzioni attuali.")
            # mostra possibili qualifiche lette (per capire la mappa)
            c_qual_dbg = find_col(df_f, ["QUALIFICA.1", "QUALIFICA"],
contains=True)
            if c_qual_dbg:
                with st.expander("🔍 Debug: valori QUALIFICA (top 50)"):

st.write(df_f[c_qual_dbg].astype(str).value_counts().head(50))
        else:
            fig1 = px.box(
                df_scope,
                x="QUALIFICA_S",
                y="ASSENZE_MAL104_ECC",
                points="all",
                title="BoxPlot Assenze (ore) per qualifica (scope)"
            )
            st.plotly_chart(fig1, use_container_width=True)

            df_scope = df_scope.copy()
            df_scope["STRAORD_TOT"] = df_scope["STRAORD_REC"] +
df_scope["STRAORD_PD"] + df_scope["STRAORD_PAG"]
            top_serv = df_scope["SERVIZIO"].value_counts().head(20).index

            fig2 = px.box(
                df_scope[df_scope["SERVIZIO"].isin(top_serv)],
                x="SERVIZIO",
                y="STRAORD_TOT",
                points=False,
                title="BoxPlot Straordinari totali (ore) per SERVIZIO
(Top 20 per numerosità)"
            )
            fig2.update_layout(xaxis_tickangle=45)
            st.plotly_chart(fig2, use_container_width=True)

    except Exception as e:
        st.error("Errore nella vista generale.")
        st.exception(e)

