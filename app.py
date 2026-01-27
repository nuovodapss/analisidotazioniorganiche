# app.py
import io
import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

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

    x = x.str.replace(" ", "", regex=False)
    return pd.to_numeric(x, errors="coerce").fillna(0)


def z(df: pd.DataFrame) -> pd.Series:
    return pd.Series(0.0, index=df.index)


def style_red_black(fig):
    fig.update_traces(marker_color="red", marker_line_color="black", marker_line_width=1)
    return fig


# =========================
# LETTURA SMART EXCEL
# =========================
EXPECTED_COLS = {
    "ENTE", "STABILIMENTO", "REPARTO", "PROFILO", "QUALIFICA", "QUALIFICA.1",
    "MATRICOLA", "DATA AL", "% PART-TIME",
    "FERIE", "FERIE RX", "FERIE GODUTE TOTALE", "FERIE GODUTE RX",
    "FERIE RES.", "FERIE RX RES.", "FERIE AP RES.",
    "PERMESSI", "MALATTIA", "MALATTIA FIGLIO", "LEGGE 104",
    "GRAVIDANZA", "COMANDO", "ASPETTATIVA", "MISSIONE SOLO SERVIZIO",
    "RECUPERO",
    "ORE DA RECUP. PROG.", "STR. PD. PROG.", "STR. PROG.",
    "FEST. INFRASETT. A PAGAMENTO", "FEST. INFRASETT. A RECUPERO",
    "CDR_DESC", "DESC. DIP.", "DESC. RUOLO",
    "PRESTAZIONI AGGIUNTIVE"
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
# LOGICA: DETTAGLIO + ANALISI
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
    df = df_raw.copy()

    # colonne principali
    c_pt = find_col(df, ["% PART-TIME", "% PART TIME", "PART-TIME", "PERC PART"], contains=True)
    c_rep = find_col(df, ["REPARTO"], contains=True)
    c_cdr = find_col(df, ["CDR_DESC", "CDR DESC", "CDR"], contains=True)
    c_matr = find_col(df, ["MATRICOLA"], contains=True)
    c_data_al = find_col(df, ["DATA AL"], contains=True)
    c_qual = find_col(df, ["QUALIFICA.1", "QUALIFICA"], contains=True)

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

    # assenze per causale (ore)
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

    # prestazioni aggiuntive (ore) — best-effort
    c_prest = find_col(df, ["PRESTAZIONI AGGIUNTIVE", "PREST. AGGIUNTIVE", "PRESTAZ"], contains=True)
    prest_agg = to_num_series(df[c_prest]) if c_prest else z(df)

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

    # prestazioni aggiuntive
    df["PREST_AGG_ORE"] = prest_agg

    # scope
    df_scope = df[df["SERVIZIO"].notna()].copy()

    # solo in forza a fine periodo
    if only_in_force and c_data_al and len(df_scope) > 0:
        try:
            max_data = df_scope[c_data_al].max()
            df_scope = df_scope[df_scope[c_data_al] == max_data].copy()
        except Exception:
            pass

    # analisi dotazioni (per servizio+qualifica)
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
            "Prestazioni aggiuntive (ore)": ("PREST_AGG_ORE", "sum"),
        }
    ).reset_index()

    agg.rename(columns={"SERVIZIO": "UUOO/SERVIZIO", "QUALIFICA_OUT": "QUALIFICA"}, inplace=True)
    agg["Media procapite"] = np.where(agg["OPERATORI"] > 0, agg["Ferie fruite 2025"] / agg["OPERATORI"], 0.0)

    ordered = [
        "UUOO/SERVIZIO", "QUALIFICA", "OPERATORI", "N° FTE",
        "st Recupero", "st PD pagato", "st Pagato",
        "Festivo pagato", "Festivo recupero",
        "Ferie maturate 2025", "Ferie fruite 2025", "Media procapite",
        "Residue al 01/01/2026",
        "Assenze totali (ore)",
        "Prestazioni aggiuntive (ore)"
    ]
    agg = agg[ordered].sort_values(["UUOO/SERVIZIO", "QUALIFICA"]).reset_index(drop=True)

    return agg, df_scope, CAUSE_COLS


def compute_kpi(df_scope: pd.DataFrame, cause_cols: list[str], day_hours: float, ore_annue_fte: float):
    # headcount
    c_matr = find_col(df_scope, ["MATRICOLA"], contains=True)
    n_operatori = int(df_scope[c_matr].nunique()) if c_matr and c_matr in df_scope.columns else int(len(df_scope))

    fte_tot = float(df_scope["FTE"].sum()) if "FTE" in df_scope.columns else 0.0

    # 1470 ore teoriche per 1 FTE (stabilito)
    ore_teo_tot = fte_tot * ore_annue_fte if fte_tot > 0 else 0.0

    # riduzione per part-time (headcount - fte)
    fte_persi_pt = max(0.0, float(n_operatori) - fte_tot)
    pct_rid_pt = (fte_persi_pt / float(n_operatori) * 100) if n_operatori > 0 else 0.0

    abs_tot_ore = float(df_scope["ASSENZE_TOT_ORE"].sum()) if "ASSENZE_TOT_ORE" in df_scope.columns else 0.0

    ass_pct = (abs_tot_ore / ore_teo_tot * 100) if ore_teo_tot > 0 else np.nan

    # FTE mediamente assenti = ore assenza / 1470
    fte_assenti = (abs_tot_ore / ore_annue_fte) if ore_annue_fte > 0 else np.nan

    # FTE disponibili = fte - fte assenti
    fte_disp = (fte_tot - fte_assenti) if isinstance(fte_assenti, float) and not np.isnan(fte_assenti) else np.nan

    # ferie
    ferie_mat = float(df_scope["FERIE_MAT_2025"].sum()) if "FERIE_MAT_2025" in df_scope.columns else 0.0
    ferie_fruite = float(df_scope["FERIE_FRUITE_2025"].sum()) if "FERIE_FRUITE_2025" in df_scope.columns else 0.0
    ferie_res_ore = float(df_scope["FERIE_RES_0101"].sum()) if "FERIE_RES_0101" in df_scope.columns else 0.0
    ferie_pct = (ferie_fruite / ferie_mat * 100) if ferie_mat > 0 else np.nan

    res_giorni = (ferie_res_ore / day_hours) if day_hours > 0 else np.nan
    res_giorni_media = (res_giorni / n_operatori) if n_operatori > 0 and not np.isnan(res_giorni) else np.nan

    # breakdown causali: ore -> FTE persi (ore/1470) e % su FTE tot
    breakdown = []
    if ore_annue_fte > 0 and fte_tot > 0:
        for c in cause_cols:
            if c in df_scope.columns:
                ore = float(df_scope[c].sum())
                fte_lost = ore / ore_annue_fte
                pct = (fte_lost / fte_tot) * 100
                breakdown.append({
                    "Causale": c.replace("ABS_", "").replace("_", " ").title(),
                    "Ore": ore,
                    "FTE_persi": fte_lost,
                    "%_su_FTE": pct
                })
    df_break = pd.DataFrame(breakdown).sort_values("%_su_FTE", ascending=False) if breakdown else pd.DataFrame(
        columns=["Causale", "Ore", "FTE_persi", "%_su_FTE"]
    )

    return {
        "n_operatori": n_operatori,
        "fte_tot": fte_tot,
        "ore_teo_tot": ore_teo_tot,
        "fte_persi_pt": fte_persi_pt,
        "pct_rid_pt": pct_rid_pt,
        "ass_pct": ass_pct,
        "fte_assenti": fte_assenti,
        "fte_disp": fte_disp,
        "ferie_mat": ferie_mat,
        "ferie_fruite": ferie_fruite,
        "ferie_pct": ferie_pct,
        "res_giorni": res_giorni,
        "res_giorni_media": res_giorni_media,
        "abs_tot_ore": abs_tot_ore,
        "df_break": df_break
    }


def totals_row_from_scope(df_scope: pd.DataFrame):
    c_matr = find_col(df_scope, ["MATRICOLA"], contains=True)
    n_operatori = int(df_scope[c_matr].nunique()) if c_matr and c_matr in df_scope.columns else int(len(df_scope))

    row = {
        "UUOO/SERVIZIO": "TOTALE",
        "QUALIFICA": "",
        "OPERATORI": n_operatori,
        "N° FTE": float(df_scope["FTE"].sum()) if "FTE" in df_scope.columns else 0.0,
        "st Recupero": float(df_scope["STRAORD_REC"].sum()) if "STRAORD_REC" in df_scope.columns else 0.0,
        "st PD pagato": float(df_scope["STRAORD_PD"].sum()) if "STRAORD_PD" in df_scope.columns else 0.0,
        "st Pagato": float(df_scope["STRAORD_PAG"].sum()) if "STRAORD_PAG" in df_scope.columns else 0.0,
        "Festivo pagato": float(df_scope["FEST_PAG"].sum()) if "FEST_PAG" in df_scope.columns else 0.0,
        "Festivo recupero": float(df_scope["FEST_REC"].sum()) if "FEST_REC" in df_scope.columns else 0.0,
        "Ferie maturate 2025": float(df_scope["FERIE_MAT_2025"].sum()) if "FERIE_MAT_2025" in df_scope.columns else 0.0,
        "Ferie fruite 2025": float(df_scope["FERIE_FRUITE_2025"].sum()) if "FERIE_FRUITE_2025" in df_scope.columns else 0.0,
        "Media procapite": np.nan,
        "Residue al 01/01/2026": float(df_scope["FERIE_RES_0101"].sum()) if "FERIE_RES_0101" in df_scope.columns else 0.0,
        "Assenze totali (ore)": float(df_scope["ASSENZE_TOT_ORE"].sum()) if "ASSENZE_TOT_ORE" in df_scope.columns else 0.0,
        "Prestazioni aggiuntive (ore)": float(df_scope["PREST_AGG_ORE"].sum()) if "PREST_AGG_ORE" in df_scope.columns else 0.0,
    }
    return pd.DataFrame([row])


# =========================
# UI
# =========================
st.title("🩺 Cruscotto Dotazioni Organiche – Upload Excel")

with st.sidebar:
    st.header("📤 Caricamento")
    uploaded = st.file_uploader("Carica PROSPETTO PERSONALE COMPARTO (xlsx)", type=["xlsx"])

    st.divider()
    st.header("⚙️ Parametri calcolo")
    ore_annue_fte = st.number_input("Ore teoriche annue per 1 FTE", min_value=800.0, max_value=2200.0, value=1470.0, step=10.0)
    day_hours = st.number_input("Ore per giorno ferie", min_value=4.0, max_value=12.0, value=7.2, step=0.1)

    st.divider()
    st.header("⚙️ Opzioni")
    only_in_force = st.toggle("Solo in forza a fine periodo (DATA AL max)", value=True)

if not uploaded:
    st.info("Carica un file Excel dalla sidebar per iniziare.")
    st.stop()

file_bytes = uploaded.getvalue()

# lettura excel smart + override
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
    st.dataframe(df_raw.head(15), use_container_width=True)

# ---- Filtri in sidebar ----
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

# costruzione analisi + scope
analisi, df_scope, CAUSE_COLS = build_detail_and_analisi(df_f, only_in_force=only_in_force)

# KPI
k = compute_kpi(df_scope, CAUSE_COLS, day_hours=day_hours, ore_annue_fte=ore_annue_fte)

st.subheader("📌 KPI (aggiornati dai filtri)")
kpi_box = st.container()
with kpi_box:
    r1 = st.columns(4)
    r1[0].metric("N Operatori", f"{k['n_operatori']}")
    r1[1].metric("FTE totali", f"{k['fte_tot']:.2f}")
    r1[2].metric("Riduzione organico (Part-time)", f"{k['fte_persi_pt']:.2f}", f"{k['pct_rid_pt']:.1f}%")
    r1[3].metric("Ore teoriche totali", f"{k['ore_teo_tot']:.0f}")

    r2 = st.columns(4)
    ass_pct_txt = f"{k['ass_pct']:.2f}%" if isinstance(k["ass_pct"], float) and not np.isnan(k["ass_pct"]) else "n/d"
    r2[0].metric("Assenteismo % (su 1470h/FTE)", ass_pct_txt)
    r2[1].metric("FTE mediamente assenti", f"{k['fte_assenti']:.2f}" if not np.isnan(k["fte_assenti"]) else "n/d")
    r2[2].metric("FTE disponibili", f"{k['fte_disp']:.2f}" if isinstance(k["fte_disp"], float) and not np.isnan(k["fte_disp"]) else "n/d")
    r2[3].metric("Assenze totali (ore)", f"{k['abs_tot_ore']:.0f}")

    r3 = st.columns(4)
    r3[0].metric("Ferie fruite / maturate", f"{k['ferie_fruite']:.0f} / {k['ferie_mat']:.0f}")
    r3[1].metric("% ferie fruite su maturate", f"{k['ferie_pct']:.1f}%" if not np.isnan(k["ferie_pct"]) else "n/d")
    r3[2].metric("Residuo ferie (giorni)", f"{k['res_giorni']:.1f}" if not np.isnan(k["res_giorni"]) else "n/d")
    r3[3].metric("Residuo medio (giorni/op)", f"{k['res_giorni_media']:.2f}" if not np.isnan(k["res_giorni_media"]) else "n/d")

    # Breakdown causali
    if not k["df_break"].empty:
        st.markdown("**% riduzione organico per causali assenza (su FTE)**")
        fig = px.bar(
            k["df_break"].head(12),
            x="%_su_FTE",
            y="Causale",
            orientation="h",
            title="Top causali per impatto su organico (FTE persi / FTE totali)"
        )
        style_red_black(fig)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Breakdown causali non disponibile (FTE=0 o dati assenze non presenti).")

st.divider()

# Tabs
tab1, tab2 = st.tabs(["📋 ANALISI_DOTAZIONI (ricostruita)", "📊 Vista generale (interattiva)"])

with tab1:
    st.subheader("Tabella ANALISI_DOTAZIONI (derivata dal prospetto)")

    # Totali sotto tabella (calcolati da df_scope per evitare doppi conteggi)
    df_total = totals_row_from_scope(df_scope)
    analisi_show = pd.concat([analisi, df_total], ignore_index=True)

    st.dataframe(analisi_show, use_container_width=True, height=520)

    st.download_button(
        "⬇️ Scarica CSV",
        data=analisi_show.to_csv(index=False).encode("utf-8"),
        file_name="ANALISI_DOTAZIONI_ricostruita_con_totali.csv",
        mime="text/csv",
    )

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        analisi_show.to_excel(writer, index=False, sheet_name="ANALISI_DOTAZIONI")
    st.download_button(
        "⬇️ Scarica Excel",
        data=buf.getvalue(),
        file_name="ANALISI_DOTAZIONI_ricostruita_con_totali.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab2:
    st.subheader("Grafici interattivi (aggiornati dai filtri)")

    if len(df_scope) == 0:
        st.warning("Nessun dato disponibile con i filtri attuali.")
        st.stop()

    # Dataset per servizio (per histogram)
    c_matr = find_col(df_scope, ["MATRICOLA"], contains=True)

    df_serv = df_scope.groupby("SERVIZIO").agg(
        OPERATORI=(c_matr, "nunique") if c_matr and c_matr in df_scope.columns else ("QUALIFICA_OUT", "size"),
        FTE=("FTE", "sum"),
        ASSENZE_ORE=("ASSENZE_TOT_ORE", "sum"),
        FERIE_RES_ORE=("FERIE_RES_0101", "sum"),
        PREST_AGG_ORE=("PREST_AGG_ORE", "sum"),
        ST_REC=("STRAORD_REC", "sum"),
        ST_PD=("STRAORD_PD", "sum"),
        ST_PAG=("STRAORD_PAG", "sum"),
    ).reset_index()

    df_serv["ORE_TEORICHE"] = df_serv["FTE"] * ore_annue_fte
    df_serv["ASSENTEISMO_%"] = np.where(df_serv["ORE_TEORICHE"] > 0, df_serv["ASSENZE_ORE"] / df_serv["ORE_TEORICHE"] * 100, np.nan)
    df_serv["FTE_PERSI_ASSENZE"] = df_serv["ASSENZE_ORE"] / ore_annue_fte
    df_serv["STRAORD_TOT_ORE"] = df_serv["ST_REC"] + df_serv["ST_PD"] + df_serv["ST_PAG"]
    df_serv["STRAORD_ORE_X_FTE"] = np.where(df_serv["FTE"] > 0, df_serv["STRAORD_TOT_ORE"] / df_serv["FTE"], np.nan)
    df_serv["PREST_ORE_X_FTE"] = np.where(df_serv["FTE"] > 0, df_serv["PREST_AGG_ORE"] / df_serv["FTE"], np.nan)
    df_serv["FERIE_RES_GIORNI_X_TESTA"] = np.where(
        df_serv["OPERATORI"] > 0,
        (df_serv["FERIE_RES_ORE"] / day_hours) / df_serv["OPERATORI"],
        np.nan
    )

    c1, c2 = st.columns(2)

    # 1) Istogramma assenteismo % (equivalente % FTE persi sul teorico)
    fig1 = px.histogram(
        df_serv.dropna(subset=["ASSENTEISMO_%"]),
        x="ASSENTEISMO_%",
        nbins=20,
        title="Istogramma Assenteismo % (su 1470h/FTE) per Servizio"
    )
    style_red_black(fig1)
    c1.plotly_chart(fig1, use_container_width=True)

    # 2) Istogramma FTE persi per assenze
    fig2 = px.histogram(
        df_serv.dropna(subset=["FTE_PERSI_ASSENZE"]),
        x="FTE_PERSI_ASSENZE",
        nbins=20,
        title="Istogramma FTE persi per assenze (ore/1470) per Servizio"
    )
    style_red_black(fig2)
    c2.plotly_chart(fig2, use_container_width=True)

    c3, c4 = st.columns(2)

    # 3) Straordinario ore per FTE
    fig3 = px.histogram(
        df_serv.dropna(subset=["STRAORD_ORE_X_FTE"]),
        x="STRAORD_ORE_X_FTE",
        nbins=20,
        title="Istogramma Straordinario (ore per FTE) per Servizio"
    )
    style_red_black(fig3)
    c3.plotly_chart(fig3, use_container_width=True)

    # 4) Prestazioni aggiuntive ore per FTE
    fig4 = px.histogram(
        df_serv.dropna(subset=["PREST_ORE_X_FTE"]),
        x="PREST_ORE_X_FTE",
        nbins=20,
        title="Istogramma Prestazioni aggiuntive (ore per FTE) per Servizio"
    )
    style_red_black(fig4)
    c4.plotly_chart(fig4, use_container_width=True)

    c5, c6 = st.columns(2)

    # 5) Ferie residue giorni per testa
    fig5 = px.histogram(
        df_serv.dropna(subset=["FERIE_RES_GIORNI_X_TESTA"]),
        x="FERIE_RES_GIORNI_X_TESTA",
        nbins=20,
        title="Istogramma Ferie residue al 01/01/2026 (giorni per testa) per Servizio"
    )
    style_red_black(fig5)
    c5.plotly_chart(fig5, use_container_width=True)

    # 6) Straordinario totale: composizione per tipologia (ore) - totale ospedale filtrato
    st_rec_tot = float(df_scope["STRAORD_REC"].sum())
    st_pd_tot = float(df_scope["STRAORD_PD"].sum())
    st_pag_tot = float(df_scope["STRAORD_PAG"].sum())

    fig6 = go.Figure()
    fig6.add_trace(go.Bar(
        name="st Recupero",
        x=["Straordinario totale"],
        y=[st_rec_tot],
        marker=dict(color="red", opacity=0.35, line=dict(color="black", width=1))
    ))
    fig6.add_trace(go.Bar(
        name="st PD pagato",
        x=["Straordinario totale"],
        y=[st_pd_tot],
        marker=dict(color="red", opacity=0.65, line=dict(color="black", width=1))
    ))
    fig6.add_trace(go.Bar(
        name="st Pagato",
        x=["Straordinario totale"],
        y=[st_pag_tot],
        marker=dict(color="red", opacity=1.0, line=dict(color="black", width=1))
    ))
    fig6.update_layout(
        barmode="stack",
        title="Straordinario totale – composizione per tipologia (ore)",
        yaxis_title="Ore"
    )
    c6.plotly_chart(fig6, use_container_width=True)

