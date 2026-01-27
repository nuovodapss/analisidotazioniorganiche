
import io
import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Cruscotto Dotazioni Organiche", layout="wide")


# =========================
# UTIL
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

    # rimuove colonne Unnamed quasi vuote
    drop_cols = []
    for c in df.columns:
        if str(c).lower().startswith("unnamed"):
            if df[c].isna().mean() >= 0.95:
                drop_cols.append(c)
    if drop_cols:
        df = df.drop(columns=drop_cols)

    df = df.dropna(how="all").reset_index(drop=True)
    return df


def find_col(df: pd.DataFrame, wanted, contains=True):
    if isinstance(wanted, str):
        wanted = [wanted]

    cols = list(df.columns)
    cols_norm = [norm(c) for c in cols]
    wanted_norm = [norm(w) for w in wanted]

    for w in wanted_norm:
        if w in cols_norm:
            return cols[cols_norm.index(w)]

    if contains:
        for i, c in enumerate(cols_norm):
            if any(w in c for w in wanted_norm):
                return cols[i]
    return None


def to_num_series(s: pd.Series) -> pd.Series:
    if s is None:
        return None
    if pd.api.types.is_numeric_dtype(s):
        return s.fillna(0)

    x = s.astype(str).str.replace("\u00a0", " ", regex=False).str.strip()

    # Formato IT: 1.234,5  ->  1234.5
    mask_both = x.str.contains(r"\.", regex=True) & x.str.contains(",", regex=False)
    x.loc[mask_both] = x.loc[mask_both].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)

    # Solo virgola -> decimale
    mask_comma = ~mask_both & x.str.contains(",", regex=False)
    x.loc[mask_comma] = x.loc[mask_comma].str.replace(",", ".", regex=False)

    # Spazi come migliaia
    x = x.str.replace(" ", "", regex=False)

    return pd.to_numeric(x, errors="coerce").fillna(0)


def z(df: pd.DataFrame) -> pd.Series:
    return pd.Series(0.0, index=df.index)


# =========================
# LETTURA SMART EXCEL
# =========================
EXPECTED_COLS = {
    "ENTE", "STABILIMENTO", "REPARTO", "PROFILO", "QUALIFICA", "QUALIFICA.1",
    "MATRICOLA", "DATA AL", "% PART-TIME", "ORE TEORICHE", "ORE LAVORATE",
    "FERIE", "FERIE RX", "FERIE GODUTE TOTALE", "FERIE GODUTE RX",
    "FERIE RES.", "FERIE RX RES.", "FERIE AP RES.",
    "PERMESSI", "MALATTIA", "MALATTIA FIGLIO", "LEGGE 104",
    "GRAVIDANZA", "COMANDO", "ASPETTATIVA", "MISSIONE SOLO SERVIZIO",
    "RECUPERO",
    "ORE DA RECUP. PROG.", "STR. PD. PROG.", "STR. PROG.",
    "FEST. INFRASETT. A PAGAMENTO", "FEST. INFRASETT. A RECUPERO",
    "CDR_DESC", "DESC. DIP.", "DESC. RUOLO"
}


def detect_best_sheet_header(xls: pd.ExcelFile, max_scan_rows=80):
    expected_norm = set(norm(x) for x in EXPECTED_COLS)
    best = {"sheet": xls.sheet_names[0], "header_row": 0, "score": -1}

    for sheet in xls.sheet_names:
        try:
            preview = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=max_scan_rows)
        except Exception:
            continue

        for r in range(min(len(preview), max_scan_rows)):
            row_vals = preview.iloc[r].dropna().tolist()
            if not row_vals:
                continue
            row_norm = set(norm(v) for v in row_vals)
            hits = len(row_norm & expected_norm)
            score = hits * 10 + len(row_vals)
            if score > best["score"]:
                best = {"sheet": sheet, "header_row": r, "score": score}

    return best


@st.cache_data
def load_excel_smart(file_bytes: bytes, sheet_override=None, header_override=None):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    if sheet_override is not None and header_override is not None:
        df = pd.read_excel(xls, sheet_name=sheet_override, header=int(header_override))
        df = clean_columns(df)
        meta = {"auto": False, "sheet": sheet_override, "header_row": int(header_override), "score": None}
        return df, meta, xls.sheet_names

    best = detect_best_sheet_header(xls)
    df = pd.read_excel(xls, sheet_name=best["sheet"], header=int(best["header_row"]))
    df = clean_columns(df)
    meta = {"auto": True, "sheet": best["sheet"], "header_row": int(best["header_row"]), "score": best["score"]}
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


def build_detail_and_analisi(df_raw: pd.DataFrame, only_in_force: bool):
    """
    Restituisce:
    - analisi: tabella aggregata tipo ANALISI_DOTAZIONI
    - df_scope: dettaglio filtrato (serve per KPI e grafici)
    """
    df = df_raw.copy()

    # colonne principali
    c_pt = find_col(df, ["% PART-TIME", "% PART TIME", "PART-TIME", "PERC PART"], contains=True)
    c_rep = find_col(df, ["REPARTO"], contains=True)
    c_cdr = find_col(df, ["CDR_DESC", "CDR DESC", "CDR"], contains=True)
    c_matr = find_col(df, ["MATRICOLA"], contains=True)
    c_data_al = find_col(df, ["DATA AL"], contains=True)
    c_qual = find_col(df, ["QUALIFICA.1", "QUALIFICA"], contains=True)

    c_ore_teo = find_col(df, ["ORE TEORICHE"], contains=True)
    c_ore_lav = find_col(df, ["ORE LAVORATE"], contains=True)

    def col_or_zero(names):
        c = find_col(df, names, contains=True)
        return to_num_series(df[c]) if c else z(df)

    # ferie (ore)
    ferie = col_or_zero(["FERIE"])
    ferie_rx = col_or_zero(["FERIE RX"])
    ferie_god_tot = col_or_zero(["FERIE GODUTE TOTALE"])
    ferie_god_rx = col_or_zero(["FERIE GODUTE RX"])
    ferie_res = col_or_zero(["FERIE RES."])
    ferie_rx_res = col_or_zero(["FERIE RX RES."])
    ferie_ap_res = col_or_zero(["FERIE AP RES."])

    # assenze per causale (ore) - estendibili
    abs_malattia = col_or_zero(["MALATTIA"])
    abs_malfiglio = col_or_zero(["MALATTIA FIGLIO"])
    abs_104 = col_or_zero(["LEGGE 104"])
    abs_permessi = col_or_zero(["PERMESSI"])
    abs_gravidanza = col_or_zero(["GRAVIDANZA"])
    abs_comando = col_or_zero(["COMANDO"])
    abs_aspettativa = col_or_zero(["ASPETTATIVA"])
    abs_aggior = col_or_zero(["AGGIOR."])
    abs_inf_malserv = col_or_zero(["INF./MAL.SERV"])
    abs_car_pub = col_or_zero(["CAR.PUBBLICA"])
    abs_covid = col_or_zero(["INFORTUNIO COVID"]) + col_or_zero(["MALATTIA COVID"])

    # altre causali spesso presenti nel prospetto
    abs_riserva = col_or_zero(["RISERVA ORARIA"])
    abs_missione = col_or_zero(["MISSIONE SOLO SERVIZIO"])
    abs_recupero = col_or_zero(["RECUPERO"])
    abs_altro = col_or_zero(["ALTRO"])
    abs_sciopero = col_or_zero(["SCIOPERO"])
    abs_sindacali = col_or_zero(["PERMESSI SINDACALI"]) + col_or_zero(["DISTACCO SINDACALE"])

    # straordinari / festivi (ore)
    st_rec = col_or_zero(["ORE DA RECUP. PROG."])
    st_pd = col_or_zero(["STR. PD. PROG."])
    st_pag = col_or_zero(["STR. PROG."])
    fest_pag = col_or_zero(["FEST. INFRASETT. A PAGAMENTO"])
    fest_rec = col_or_zero(["FEST. INFRASETT. A RECUPERO"])

    # FTE
    if c_pt:
        pt = to_num_series(df[c_pt])
        maxv = float(pt.max()) if len(pt) else 0.0
        if maxv <= 1.0:
            df["FTE"] = pt
        else:
            pt = pt.where(pt > 0, 100)
            df["FTE"] = pt / 100
    else:
        df["FTE"] = 1.0

    # SERVIZIO: CDR_DESC fallback REPARTO
    if c_cdr:
        df["SERVIZIO"] = df[c_cdr].astype(str).str.strip()
    else:
        df["SERVIZIO"] = ""
    if c_rep:
        rep_str = df[c_rep].astype(str).str.strip()
        df.loc[df["SERVIZIO"].isin(["", "nan", "None"]), "SERVIZIO"] = rep_str
        df.loc[df["SERVIZIO"].eq(""), "SERVIZIO"] = rep_str

    # QUALIFICA semplificata (fallback originale)
    rep_for_q = df[c_rep] if c_rep else pd.Series([""] * len(df), index=df.index)
    q_for_q = df[c_qual] if c_qual else pd.Series([None] * len(df), index=df.index)
    df["QUALIFICA_S"] = [simplify_qualifica(q, r) for q, r in zip(q_for_q.tolist(), rep_for_q.tolist())]
    df["QUALIFICA_OUT"] = df["QUALIFICA_S"].fillna("N/D").astype(str)

    # ferie derivate
    df["FERIE_MAT_2025"] = ferie + ferie_rx
    df["FERIE_FRUITE_2025"] = ferie_god_tot + ferie_god_rx
    df["FERIE_RES_0101"] = ferie_res + ferie_rx_res + ferie_ap_res

    # ore teoriche/lavorate (se presenti)
    df["ORE_TEORICHE"] = to_num_series(df[c_ore_teo]) if c_ore_teo else z(df)
    df["ORE_LAVORATE"] = to_num_series(df[c_ore_lav]) if c_ore_lav else z(df)

    # causali in dettaglio
    df["ABS_MALATTIA"] = abs_malattia
    df["ABS_MALFIGLIO"] = abs_malfiglio
    df["ABS_104"] = abs_104
    df["ABS_PERMESSI"] = abs_permessi
    df["ABS_GRAVIDANZA"] = abs_gravidanza
    df["ABS_COMANDO"] = abs_comando
    df["ABS_ASPETTATIVA"] = abs_aspettativa
    df["ABS_AGGIORNAMENTO"] = abs_aggior
    df["ABS_INF_MALSERV"] = abs_inf_malserv
    df["ABS_CAR_PUB"] = abs_car_pub
    df["ABS_COVID"] = abs_covid
    df["ABS_RISERVA_ORARIA"] = abs_riserva
    df["ABS_MISSIONE"] = abs_missione
    df["ABS_RECUPERO"] = abs_recupero
    df["ABS_SINDACALI"] = abs_sindacali
    df["ABS_SCIOPERO"] = abs_sciopero
    df["ABS_ALTRO"] = abs_altro

    # totale assenze (ore) = somma causali selezionate
    CAUSE_COLS = [
        "ABS_MALATTIA", "ABS_MALFIGLIO", "ABS_104", "ABS_PERMESSI",
        "ABS_GRAVIDANZA", "ABS_COMANDO", "ABS_ASPETTATIVA",
        "ABS_AGGIORNAMENTO", "ABS_INF_MALSERV", "ABS_CAR_PUB",
        "ABS_COVID", "ABS_RISERVA_ORARIA", "ABS_MISSIONE",
        "ABS_RECUPERO", "ABS_SINDACALI", "ABS_SCIOPERO", "ABS_ALTRO"
    ]
    df["ASSENZE_TOT_ORE"] = df[CAUSE_COLS].sum(axis=1)

    # straordinari/festivi
    df["STRAORD_REC"] = st_rec
    df["STRAORD_PD"] = st_pd
    df["STRAORD_PAG"] = st_pag
    df["FEST_PAG"] = fest_pag
    df["FEST_REC"] = fest_rec

    # scope: record con servizio valorizzato
    df_scope = df[df["SERVIZIO"].notna()].copy()

    # solo in forza a fine periodo
    if only_in_force and c_data_al and len(df_scope) > 0:
        try:
            max_data = df_scope[c_data_al].max()
            df_scope = df_scope[df_scope[c_data_al] == max_data].copy()
        except Exception:
            pass

    # aggregazione "ANALISI_DOTAZIONI"
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
            "Assenze totali (ore)": ("ASSENZE_TOT_ORE", "sum"),
        }
    ).reset_index()

    agg.rename(columns={"SERVIZIO": "UUOO/SERVIZIO", "QUALIFICA_OUT": "QUALIFICA"}, inplace=True)
    agg["Media procapite"] = np.where(agg["OPERATORI"] > 0, agg["Ferie fruite 2025"] / agg["OPERATORI"], 0.0)
    agg["Prestazioni aggiuntive (ore)"] = 0.0

    ordered = [
        "UUOO/SERVIZIO", "QUALIFICA", "OPERATORI", "N° FTE",
        "st Recupero", "st PD pagato", "st Pagato",
        "Festivo pagato", "Festivo recupero",
        "Ferie maturate 2025", "Ferie fruite 2025", "Media procapite",
        "Residue al 01/01/2026",
        "Assenze totali (ore)",
        "Prestazioni aggiuntive (ore)"
    ]
    for c in ordered:
        if c not in agg.columns:
            agg[c] = 0.0
    agg = agg[ordered].sort_values(["UUOO/SERVIZIO", "QUALIFICA"]).reset_index(drop=True)

    return agg, df_scope, CAUSE_COLS


def compute_kpi(df_scope: pd.DataFrame, cause_cols: list[str], day_hours: float):
    # headcount
    c_matr = find_col(df_scope, ["MATRICOLA"], contains=True)
    n_operatori = int(df_scope[c_matr].nunique()) if c_matr and c_matr in df_scope.columns else int(len(df_scope))

    fte_tot = float(df_scope["FTE"].sum()) if "FTE" in df_scope.columns else 0.0

    # riduzione per part-time (headcount - fte)
    fte_persi_pt = max(0.0, float(n_operatori) - fte_tot)
    pct_rid_pt = (fte_persi_pt / float(n_operatori) * 100) if n_operatori > 0 else 0.0

    ore_teo_tot = float(df_scope["ORE_TEORICHE"].sum()) if "ORE_TEORICHE" in df_scope.columns else 0.0

    # assenze totali ore
    abs_tot_ore = float(df_scope["ASSENZE_TOT_ORE"].sum()) if "ASSENZE_TOT_ORE" in df_scope.columns else 0.0

    # assenteismo % su ore teoriche
    ass_pct = (abs_tot_ore / ore_teo_tot * 100) if ore_teo_tot > 0 else np.nan

    # ore per 1 FTE nel periodo (calcolata dal dataset stesso)
    ore_per_fte = (ore_teo_tot / fte_tot) if fte_tot > 0 and ore_teo_tot > 0 else np.nan

    # FTE mediamente assenti
    fte_assenti = (abs_tot_ore / ore_per_fte) if ore_per_fte and ore_per_fte > 0 else np.nan

    # FTE disponibili
    fte_disp = (fte_tot - fte_assenti) if (isinstance(fte_assenti, float) and not np.isnan(fte_assenti)) else np.nan

    # ferie
    ferie_mat = float(df_scope["FERIE_MAT_2025"].sum()) if "FERIE_MAT_2025" in df_scope.columns else 0.0
    ferie_fruite = float(df_scope["FERIE_FRUITE_2025"].sum()) if "FERIE_FRUITE_2025" in df_scope.columns else 0.0
    ferie_res_ore = float(df_scope["FERIE_RES_0101"].sum()) if "FERIE_RES_0101" in df_scope.columns else 0.0

    ferie_pct = (ferie_fruite / ferie_mat * 100) if ferie_mat > 0 else np.nan

    # residuo ferie in giorni e medio per operatore
    res_giorni = (ferie_res_ore / day_hours) if day_hours > 0 else np.nan
    res_giorni_media = (res_giorni / n_operatori) if n_operatori > 0 and not np.isnan(res_giorni) else np.nan

    # breakdown riduzione organico per causali assenza (in FTE equivalenti e %)
    breakdown = []
    if ore_per_fte and ore_per_fte > 0 and fte_tot > 0:
        for c in cause_cols:
            if c in df_scope.columns:
                ore = float(df_scope[c].sum())
                fte_lost = ore / ore_per_fte
                pct = (fte_lost / fte_tot) * 100
                breakdown.append({"Causale": c.replace("ABS_", "").replace("_", " ").title(), "Ore": ore, "FTE_persi": fte_lost, "%_su_FTE": pct})

    df_break = pd.DataFrame(breakdown).sort_values("%_su_FTE", ascending=False) if breakdown else pd.DataFrame(columns=["Causale","Ore","FTE_persi","%_su_FTE"])

    return {
        "n_operatori": n_operatori,
        "fte_tot": fte_tot,
        "fte_persi_pt": fte_persi_pt,
        "pct_rid_pt": pct_rid_pt,
        "ore_teo_tot": ore_teo_tot,
        "ass_pct": ass_pct,
        "fte_assenti": fte_assenti,
        "fte_disp": fte_disp,
        "ferie_mat": ferie_mat,
        "ferie_fruite": ferie_fruite,
        "ferie_pct": ferie_pct,
        "res_giorni": res_giorni,
        "res_giorni_media": res_giorni_media,
        "df_break": df_break
    }


# =========================
# UI
# =========================
st.title("🩺 Cruscotto Dotazioni Organiche – Upload Excel")

with st.sidebar:
    st.header("📤 Caricamento")
    uploaded = st.file_uploader("Carica PROSPETTO PERSONALE COMPARTO (xlsx)", type=["xlsx"])

    st.divider()
    st.header("⚙️ Opzioni")
    only_in_force = st.toggle("Solo in forza a fine periodo (DATA AL max)", value=True)
    day_hours = st.number_input("Ore per giorno ferie", min_value=4.0, max_value=12.0, value=7.2, step=0.1)

if not uploaded:
    st.info("Carica un file Excel dalla sidebar per iniziare.")
    st.stop()

file_bytes = uploaded.getvalue()

# lettura auto + override
try:
    df_raw, meta, sheet_names = load_excel_smart(file_bytes)
except Exception as e:
    st.error("Errore durante lettura Excel (auto).")
    st.exception(e)
    st.stop()

with st.sidebar:
    st.divider()
    st.header("🧠 Lettura Excel")
    override = st.checkbox("Override manuale (foglio + riga header)", value=False)

if override:
    with st.sidebar:
        sheet_override = st.selectbox(
            "Foglio",
            sheet_names,
            index=sheet_names.index(meta["sheet"]) if meta["sheet"] in sheet_names else 0
        )
        header_override = st.number_input("Riga header (0=prima riga)", min_value=0, max_value=300, value=int(meta["header_row"]), step=1)
    try:
        df_raw, meta, sheet_names = load_excel_smart(file_bytes, sheet_override, int(header_override))
    except Exception as e:
        st.error("Errore durante lettura Excel (override).")
        st.exception(e)
        st.stop()

with st.expander("🔎 Debug lettura Excel"):
    st.write(meta)
    st.write("Colonne lette:")
    st.write(list(df_raw.columns))
    st.dataframe(df_raw.head(20), use_container_width=True)

# ---- Filtri generali (in sidebar) ----
col_dip = find_col(df_raw, ["DESC. DIP.", "DESC DIP"], contains=True)
col_stab = find_col(df_raw, ["STABILIMENTO"], contains=True)
col_cdr = find_col(df_raw, ["CDR_DESC", "CDR DESC", "CDR"], contains=True)
col_rep = find_col(df_raw, ["REPARTO"], contains=True)
col_prof = find_col(df_raw, ["PROFILO"], contains=True)
col_qual = find_col(df_raw, ["QUALIFICA.1", "QUALIFICA"], contains=True)
col_ruolo = find_col(df_raw, ["DESC. RUOLO", "RUOLO"], contains=True)

with st.sidebar:
    st.divider()
    st.header("🎛️ Filtri (dataset)")

    def opts(col):
        return sorted(df_raw[col].dropna().astype(str).unique()) if col else []

    dip_opts = opts(col_dip)
    stab_opts = opts(col_stab)
    cdr_opts = opts(col_cdr)
    rep_opts = opts(col_rep)
    prof_opts = opts(col_prof)
    qual_opts = opts(col_qual)
    ruolo_opts = opts(col_ruolo)

    dip_sel = st.multiselect("Dipartimento", dip_opts, default=dip_opts) if dip_opts else []
    stab_sel = st.multiselect("Stabilimento", stab_opts, default=stab_opts) if stab_opts else []
    cdr_sel = st.multiselect("CDR_DESC", cdr_opts, default=cdr_opts) if cdr_opts else []
    rep_sel = st.multiselect("Reparto", rep_opts, default=rep_opts) if rep_opts else []

    st.markdown("**Filtri professionali**")
    prof_sel = st.multiselect("Profilo", prof_opts, default=prof_opts) if prof_opts else []
    qual_sel = st.multiselect("Qualifica", qual_opts, default=qual_opts) if qual_opts else []
    ruolo_sel = st.multiselect("Ruolo", ruolo_opts, default=ruolo_opts) if ruolo_opts else []

# applica filtri
df_f = df_raw.copy()
if col_dip and dip_sel:
    df_f = df_f[df_f[col_dip].astype(str).isin(dip_sel)]
if col_stab and stab_sel:
    df_f = df_f[df_f[col_stab].astype(str).isin(stab_sel)]
if col_cdr and cdr_sel:
    df_f = df_f[df_f[col_cdr].astype(str).isin(cdr_sel)]
if col_rep and rep_sel:
    df_f = df_f[df_f[col_rep].astype(str).isin(rep_sel)]
if col_prof and prof_sel:
    df_f = df_f[df_f[col_prof].astype(str).isin(prof_sel)]
if col_qual and qual_sel:
    df_f = df_f[df_f[col_qual].astype(str).isin(qual_sel)]
if col_ruolo and ruolo_sel:
    df_f = df_f[df_f[col_ruolo].astype(str).isin(ruolo_sel)]

# costruisci analisi + dettaglio
try:
    analisi, df_scope, CAUSE_COLS = build_detail_and_analisi(df_f, only_in_force=only_in_force)
except Exception as e:
    st.error("Errore nella costruzione dati (analisi/dettaglio).")
    st.exception(e)
    st.stop()

# KPI
k = compute_kpi(df_scope, CAUSE_COLS, day_hours=day_hours)

st.subheader("📌 KPI (aggiornati dai filtri)")
kpi_box = st.container(border=True)

with kpi_box:
    r1 = st.columns(4)
    r1[0].metric("N Operatori", f"{k['n_operatori']}")
    r1[1].metric("FTE totali", f"{k['fte_tot']:.2f}")
    r1[2].metric("Riduzione organico (Part-time)", f"{k['fte_persi_pt']:.2f}", f"{k['pct_rid_pt']:.1f}%")
    r1[3].metric("Ore teoriche totali", f"{k['ore_teo_tot']:.0f}")

    r2 = st.columns(4)
    ass_pct_txt = f"{k['ass_pct']:.2f}%" if isinstance(k["ass_pct"], float) and not np.isnan(k["ass_pct"]) else "n/d"
    fte_ass_txt = f"{k['fte_assenti']:.2f}" if isinstance(k["fte_assenti"], float) and not np.isnan(k["fte_assenti"]) else "n/d"
    fte_disp_txt = f"{k['fte_disp']:.2f}" if isinstance(k["fte_disp"], float) and not np.isnan(k["fte_disp"]) else "n/d"

    r2[0].metric("Assenteismo % (su ore teoriche)", ass_pct_txt)
    r2[1].metric("FTE mediamente assenti", fte_ass_txt)
    r2[2].metric("FTE disponibili", fte_disp_txt)
    r2[3].metric("Ferie fruite / maturate", f"{k['ferie_fruite']:.0f} / {k['ferie_mat']:.0f}")

    r3 = st.columns(4)
    ferie_pct_txt = f"{k['ferie_pct']:.1f}%" if isinstance(k["ferie_pct"], float) and not np.isnan(k["ferie_pct"]) else "n/d"
    res_g_txt = f"{k['res_giorni']:.1f}" if isinstance(k["res_giorni"], float) and not np.isnan(k["res_giorni"]) else "n/d"
    res_g_med_txt = f"{k['res_giorni_media']:.2f}" if isinstance(k["res_giorni_media"], float) and not np.isnan(k["res_giorni_media"]) else "n/d"

    r3[0].metric("% ferie fruite su maturate", ferie_pct_txt)
    r3[1].metric("Residuo ferie (giorni)", res_g_txt)
    r3[2].metric("Residuo ferie medio (giorni/op)", res_g_med_txt)
    r3[3].metric("Assenze totali (ore)", f"{float(df_scope['ASSENZE_TOT_ORE'].sum()):.0f}" if "ASSENZE_TOT_ORE" in df_scope.columns else "n/d")

    # Breakdown causali: grafico + tabella
    if not k["df_break"].empty:
        st.markdown("**% riduzione organico per causali assenza (su FTE)**")
        fig = px.bar(
            k["df_break"].head(12),
            x="%_su_FTE",
            y="Causale",
            orientation="h",
            title="Top causali per impatto su organico (FTE persi / FTE totali)"
        )
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("Dettaglio causali (tutte)"):
            st.dataframe(k["df_break"], use_container_width=True)
    else:
        st.info("Breakdown causali non disponibile (mancano ORE TEORICHE oppure FTE=0).")

st.divider()

# Tabs: tabella + vista generale
tab1, tab2 = st.tabs(["📋 ANALISI_DOTAZIONI (ricostruita)", "📊 Vista generale (boxplot)"])

with tab1:
    st.subheader("Tabella ANALISI_DOTAZIONI (derivata dal prospetto)")
    st.dataframe(analisi, use_container_width=True, height=520)

    st.download_button(
        "⬇️ Scarica CSV",
        data=analisi.to_csv(index=False).encode("utf-8"),
        file_name="ANALISI_DOTAZIONI_ricostruita.csv",
        mime="text/csv",
    )

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        analisi.to_excel(writer, index=False, sheet_name="ANALISI_DOTAZIONI")
    st.download_button(
        "⬇️ Scarica Excel",
        data=buf.getvalue(),
        file_name="ANALISI_DOTAZIONI_ricostruita.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab2:
    st.subheader("Distribuzioni (dataset filtrato)")

    if len(df_scope) == 0:
        st.warning("Nessun dato disponibile con i filtri attuali.")
    else:
        # Boxplot assenze per profilo (se presente) altrimenti per qualifica semplificata
        if col_prof and col_prof in df_f.columns:
            # ricostruisci una colonna profilo coerente nel df_scope
            df_scope_plot = df_scope.copy()
            df_scope_plot["PROFILO"] = df_f.loc[df_scope_plot.index, col_prof].astype(str).values
            fig1 = px.box(
                df_scope_plot,
                x="PROFILO",
                y="ASSENZE_TOT_ORE",
                points=False,
                title="BoxPlot Assenze totali (ore) per PROFILO"
            )
            fig1.update_layout(xaxis_tickangle=45)
            st.plotly_chart(fig1, use_container_width=True)
        else:
            fig1 = px.box(
                df_scope,
                x="QUALIFICA_OUT",
                y="ASSENZE_TOT_ORE",
                points="all",
                title="BoxPlot Assenze totali (ore) per QUALIFICA"
            )
            st.plotly_chart(fig1, use_container_width=True)

        # Boxplot straordinari totali per servizio (top 20)
        df_scope2 = df_scope.copy()
        df_scope2["STRAORD_TOT"] = df_scope2["STRAORD_REC"] + df_scope2["STRAORD_PD"] + df_scope2["STRAORD_PAG"]
        top_serv = df_scope2["SERVIZIO"].value_counts().head(20).index

        fig2 = px.box(
            df_scope2[df_scope2["SERVIZIO"].isin(top_serv)],
            x="SERVIZIO",
            y="STRAORD_TOT",
            points=False,
            title="BoxPlot Straordinari totali (ore) per SERVIZIO (Top 20)"
        )
        fig2.update_layout(xaxis_tickangle=45)
        st.plotly_chart(fig2, use_container_width=True)


