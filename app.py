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

    # Formato IT: 1.234,5 -> 1234.5
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
    "MATRICOLA", "COGNOME", "NOME", "DATA AL", "% PART-TIME",
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

    # prestazioni aggiuntive (ore) — best effort
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

    # QUALIFICA semplificata
    rep_for_q = df[c_rep] if c_rep else pd.Series([""] * len(df), index=df.index)
    q_for_q = df[c_qual] if c_qual else pd.Series([None] * len(df), index=df.index)
    df["QUALIFICA_S"] = [simplify_qualifica(q, r) for q, r in zip(q_for_q.tolist(), rep_for_q.tolist())]
    df["QUALIFICA_OUT"] = df["QUALIFICA_S"].fillna("N/D").astype(str)

    # ferie derivate
    df["FERIE_MAT_2025"] = ferie + ferie_rx
    df["FERIE_FRUITE_2025"] = ferie_god_tot + ferie_god_rx
    df["FERIE_RES_0101"] = ferie_res + ferie_rx_res + ferie_ap_res

    # causali dettaglio
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

    # analisi (per servizio+qualifica)
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


def compute_kpi(df_scope: pd.DataFrame, day_hours: float, ore_annue_fte: float, cause_cols: list[str] | None = None):
    c_matr = find_col(df_scope, ["MATRICOLA"], contains=True)
    n_operatori = int(df_scope[c_matr].nunique()) if c_matr and c_matr in df_scope.columns else int(len(df_scope))

    fte_tot = float(df_scope["FTE"].sum()) if "FTE" in df_scope.columns else 0.0
    ore_teo_tot = fte_tot * ore_annue_fte if fte_tot > 0 else 0.0

    abs_tot_ore = float(df_scope["ASSENZE_TOT_ORE"].sum()) if "ASSENZE_TOT_ORE" in df_scope.columns else 0.0
    ass_pct = (abs_tot_ore / ore_teo_tot * 100) if ore_teo_tot > 0 else np.nan

    fte_assenti = (abs_tot_ore / ore_annue_fte) if ore_annue_fte > 0 else np.nan
    fte_disp = (fte_tot - fte_assenti) if isinstance(fte_assenti, float) and not np.isnan(fte_assenti) else np.nan

    ferie_mat = float(df_scope["FERIE_MAT_2025"].sum()) if "FERIE_MAT_2025" in df_scope.columns else 0.0
    ferie_fruite = float(df_scope["FERIE_FRUITE_2025"].sum()) if "FERIE_FRUITE_2025" in df_scope.columns else 0.0
    ferie_res_ore = float(df_scope["FERIE_RES_0101"].sum()) if "FERIE_RES_0101" in df_scope.columns else 0.0

    ferie_pct = (ferie_fruite / ferie_mat * 100) if ferie_mat > 0 else np.nan
    res_giorni = (ferie_res_ore / day_hours) if day_hours > 0 else np.nan
    res_giorni_media = (res_giorni / n_operatori) if n_operatori > 0 and not np.isnan(res_giorni) else np.nan

    st_tot = float(df_scope["STRAORD_REC"].sum() + df_scope["STRAORD_PD"].sum() + df_scope["STRAORD_PAG"].sum()) if \
        all(c in df_scope.columns for c in ["STRAORD_REC", "STRAORD_PD", "STRAORD_PAG"]) else 0.0
    st_x_fte = (st_tot / fte_tot) if fte_tot > 0 else np.nan

    prest_tot = float(df_scope["PREST_AGG_ORE"].sum()) if "PREST_AGG_ORE" in df_scope.columns else 0.0
    prest_x_fte = (prest_tot / fte_tot) if fte_tot > 0 else np.nan

    df_break = None
    if cause_cols:
        breakdown = []
        if ore_annue_fte > 0 and fte_tot > 0:
            for c in cause_cols:
                if c in df_scope.columns:
                    ore = float(df_scope[c].sum())
                    fte_lost = ore / ore_annue_fte
                    pct = (fte_lost / fte_tot) * 100
                    breakdown.append({"Causale": c.replace("ABS_", "").replace("_", " ").title(),
                                      "Ore": ore, "FTE_persi": fte_lost, "%_su_FTE": pct})
        df_break = pd.DataFrame(breakdown).sort_values("%_su_FTE", ascending=False) if breakdown else pd.DataFrame(
            columns=["Causale", "Ore", "FTE_persi", "%_su_FTE"]
        )

    return {
        "n_operatori": n_operatori,
        "fte_tot": fte_tot,
        "ore_teo_tot": ore_teo_tot,
        "abs_tot_ore": abs_tot_ore,
        "ass_pct": ass_pct,
        "fte_assenti": fte_assenti,
        "fte_disp": fte_disp,
        "ferie_mat": ferie_mat,
        "ferie_fruite": ferie_fruite,
        "ferie_pct": ferie_pct,
        "res_giorni": res_giorni,
        "res_giorni_media": res_giorni_media,
        "st_tot": st_tot,
        "st_x_fte": st_x_fte,
        "prest_tot": prest_tot,
        "prest_x_fte": prest_x_fte,
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


def build_people_table(df_sub: pd.DataFrame, ore_annue_fte: float, day_hours: float, cause_cols: list[str]):
    c_matr = find_col(df_sub, ["MATRICOLA"], contains=True)
    c_cogn = find_col(df_sub, ["COGNOME"], contains=True)
    c_nome = find_col(df_sub, ["NOME"], contains=True)
    c_prof = find_col(df_sub, ["PROFILO"], contains=True)
    c_qual = find_col(df_sub, ["QUALIFICA.1", "QUALIFICA"], contains=True)

    # chiave persona
    if c_matr:
        key_cols = [c_matr]
    else:
        # fallback (meno robusto)
        key_cols = [c_cogn, c_nome] if c_cogn and c_nome else []

    if not key_cols:
        # se proprio non esiste nulla, ritorna vuoto
        return pd.DataFrame()

    # aggrega per persona
    agg_dict = {
        "FTE": ("FTE", "sum"),
        "ASSENZE_ORE": ("ASSENZE_TOT_ORE", "sum"),
        "FERIE_RES_ORE": ("FERIE_RES_0101", "sum"),
        "FERIE_MAT_ORE": ("FERIE_MAT_2025", "sum"),
        "FERIE_FRUITE_ORE": ("FERIE_FRUITE_2025", "sum"),
        "ST_REC": ("STRAORD_REC", "sum"),
        "ST_PD": ("STRAORD_PD", "sum"),
        "ST_PAG": ("STRAORD_PAG", "sum"),
        "PREST_AGG_ORE": ("PREST_AGG_ORE", "sum"),
    }

    # aggiungo causali
    for c in cause_cols:
        if c in df_sub.columns:
            agg_dict[c] = (c, "sum")

    gb = df_sub.groupby(key_cols, dropna=False).agg(**{k: v for k, v in agg_dict.items()}).reset_index()

    # aggiungo colonne anagrafiche “first”
    def add_first(colname, outname):
        if colname and colname in df_sub.columns:
            firsts = df_sub.groupby(key_cols, dropna=False)[colname].first().reset_index().rename(columns={colname: outname})
            return gb.merge(firsts, on=key_cols, how="left")
        return gb

    gb = add_first(c_cogn, "COGNOME")
    gb = add_first(c_nome, "NOME")
    gb = add_first(c_prof, "PROFILO")
    gb = add_first(c_qual, "QUALIFICA_RAW")

    # metriche
    gb["ORE_TEORICHE"] = gb["FTE"] * ore_annue_fte
    gb["ASSENTEISMO_%"] = np.where(gb["ORE_TEORICHE"] > 0, gb["ASSENZE_ORE"] / gb["ORE_TEORICHE"] * 100, np.nan)
    gb["FTE_ASSENTI"] = gb["ASSENZE_ORE"] / ore_annue_fte if ore_annue_fte > 0 else np.nan
    gb["FTE_DISPONIBILI"] = gb["FTE"] - gb["FTE_ASSENTI"]

    gb["STRAORD_TOT_ORE"] = gb["ST_REC"] + gb["ST_PD"] + gb["ST_PAG"]
    gb["STRAORD_ORE_X_FTE"] = np.where(gb["FTE"] > 0, gb["STRAORD_TOT_ORE"] / gb["FTE"], np.nan)
    gb["PREST_ORE_X_FTE"] = np.where(gb["FTE"] > 0, gb["PREST_AGG_ORE"] / gb["FTE"], np.nan)

    gb["FERIE_RES_GIORNI"] = np.where(day_hours > 0, gb["FERIE_RES_ORE"] / day_hours, np.nan)
    gb["FERIE_RES_GIORNI_X_FTE"] = np.where(gb["FTE"] > 0, gb["FERIE_RES_GIORNI"] / gb["FTE"], np.nan)

    # colonna display nome
    if "COGNOME" in gb.columns and "NOME" in gb.columns:
        gb["PERSONA"] = gb["COGNOME"].astype(str).str.strip() + " " + gb["NOME"].astype(str).str.strip()
    elif c_matr and c_matr in gb.columns:
        gb["PERSONA"] = gb[c_matr].astype(str)
    else:
        gb["PERSONA"] = "N/D"

    # ordina colonne principali
    cols_front = []
    if c_matr and c_matr in gb.columns:
        cols_front.append(c_matr)
    cols_front += [c for c in ["PERSONA", "COGNOME", "NOME", "PROFILO", "QUALIFICA_RAW"] if c in gb.columns]

    cols_metrics = [
        "FTE", "ORE_TEORICHE", "ASSENZE_ORE", "ASSENTEISMO_%",
        "FTE_ASSENTI", "FTE_DISPONIBILI",
        "STRAORD_TOT_ORE", "STRAORD_ORE_X_FTE",
        "PREST_AGG_ORE", "PREST_ORE_X_FTE",
        "FERIE_RES_ORE", "FERIE_RES_GIORNI", "FERIE_RES_GIORNI_X_FTE",
        "FERIE_MAT_ORE", "FERIE_FRUITE_ORE",
        "ST_REC", "ST_PD", "ST_PAG",
    ]
    cols_metrics = [c for c in cols_metrics if c in gb.columns]

    other_cols = [c for c in gb.columns if c not in cols_front + cols_metrics]
    gb = gb[cols_front + cols_metrics + other_cols]

    return gb


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
df_raw, meta, sheet_names = load_excel_smart(file_bytes)

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
    df_raw, meta, sheet_names = load_excel_smart(file_bytes, sheet_override, int(header_override))

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

# KPI globali
k_global = compute_kpi(df_scope, day_hours=day_hours, ore_annue_fte=ore_annue_fte, cause_cols=CAUSE_COLS)

st.subheader("📌 KPI (aggiornati dai filtri)")
with st.container(border=True):
    r1 = st.columns(4)
    r1[0].metric("N Operatori", f"{k_global['n_operatori']}")
    r1[1].metric("FTE totali", f"{k_global['fte_tot']:.2f}")
    r1[2].metric("Assenteismo % (su 1470h/FTE)", f"{k_global['ass_pct']:.2f}%" if not np.isnan(k_global["ass_pct"]) else "n/d")
    r1[3].metric("FTE mediamente assenti", f"{k_global['fte_assenti']:.2f}" if not np.isnan(k_global["fte_assenti"]) else "n/d")

    r2 = st.columns(4)
    r2[0].metric("FTE disponibili", f"{k_global['fte_disp']:.2f}" if not np.isnan(k_global["fte_disp"]) else "n/d")
    r2[1].metric("Straordinario (ore/FTE)", f"{k_global['st_x_fte']:.2f}" if not np.isnan(k_global["st_x_fte"]) else "n/d")
    r2[2].metric("Prestazioni agg (ore/FTE)", f"{k_global['prest_x_fte']:.2f}" if not np.isnan(k_global["prest_x_fte"]) else "n/d")
    r2[3].metric("Residuo ferie medio (gg/op)", f"{k_global['res_giorni_media']:.2f}" if not np.isnan(k_global["res_giorni_media"]) else "n/d")

st.divider()

tab1, tab2, tab3 = st.tabs([
    "📋 ANALISI_DOTAZIONI (ricostruita)",
    "📊 Vista Generale (interattiva)",
    "👤 Dettaglio Reparto & Persone"
])

# =========================
# TAB 1
# =========================
with tab1:
    st.subheader("Tabella ANALISI_DOTAZIONI (derivata dal prospetto)")
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

# =========================
# TAB 2
# =========================
with tab2:
    st.subheader("Vista generale (interattiva)")

    if len(df_scope) == 0:
        st.warning("Nessun dato disponibile con i filtri attuali.")
        st.stop()

    col_rep_scope = find_col(df_scope, ["REPARTO"], contains=True)
    dim_label = "REPARTO" if col_rep_scope else "SERVIZIO"
    dim_col = col_rep_scope if col_rep_scope else "SERVIZIO"
    st.caption(f"Aggregazione per: **{dim_label}**")

    top_n = st.slider(f"Top {dim_label} nei grafici", min_value=5, max_value=60, value=20, step=5)

    c_matr = find_col(df_scope, ["MATRICOLA"], contains=True)

    df_dim = df_scope.groupby(dim_col).agg(
        OPERATORI=(c_matr, "nunique") if c_matr and c_matr in df_scope.columns else ("QUALIFICA_OUT", "size"),
        FTE=("FTE", "sum"),
        ASSENZE_ORE=("ASSENZE_TOT_ORE", "sum"),
        FERIE_RES_ORE=("FERIE_RES_0101", "sum"),
        PREST_AGG_ORE=("PREST_AGG_ORE", "sum"),
        ST_REC=("STRAORD_REC", "sum"),
        ST_PD=("STRAORD_PD", "sum"),
        ST_PAG=("STRAORD_PAG", "sum"),
    ).reset_index().rename(columns={dim_col: dim_label})

    df_dim["ORE_TEORICHE"] = df_dim["FTE"] * ore_annue_fte
    df_dim["ASSENTEISMO_%"] = np.where(df_dim["ORE_TEORICHE"] > 0, df_dim["ASSENZE_ORE"] / df_dim["ORE_TEORICHE"] * 100, np.nan)
    df_dim["FTE_PERSI_ASSENZE"] = df_dim["ASSENZE_ORE"] / ore_annue_fte
    df_dim["STRAORD_TOT_ORE"] = df_dim["ST_REC"] + df_dim["ST_PD"] + df_dim["ST_PAG"]
    df_dim["STRAORD_ORE_X_FTE"] = np.where(df_dim["FTE"] > 0, df_dim["STRAORD_TOT_ORE"] / df_dim["FTE"], np.nan)
    df_dim["PREST_ORE_X_FTE"] = np.where(df_dim["FTE"] > 0, df_dim["PREST_AGG_ORE"] / df_dim["FTE"], np.nan)
    df_dim["FERIE_RES_GIORNI_X_TESTA"] = np.where(
        df_dim["OPERATORI"] > 0, (df_dim["FERIE_RES_ORE"] / day_hours) / df_dim["OPERATORI"], np.nan
    )

    # Scatter
    st.subheader("Scatter (interattivo)")
    scatter_opts = {
        "Assenteismo %": "ASSENTEISMO_%",
        "Straordinario (ore/FTE)": "STRAORD_ORE_X_FTE",
        "Prestazioni aggiuntive (ore/FTE)": "PREST_ORE_X_FTE",
        "Ferie residue (gg/testa)": "FERIE_RES_GIORNI_X_TESTA",
        "FTE persi per assenze": "FTE_PERSI_ASSENZE",
    }
    c_sc1, c_sc2, c_sc3 = st.columns(3)
    x_metric = c_sc1.selectbox("Asse X", list(scatter_opts.keys()), index=0, key="sc_x")
    y_metric = c_sc2.selectbox("Asse Y", list(scatter_opts.keys()), index=1, key="sc_y")
    size_metric = c_sc3.selectbox("Dimensione (bolla)", ["FTE", "OPERATORI"], index=0, key="sc_s")

    df_sc = df_dim.dropna(subset=[scatter_opts[x_metric], scatter_opts[y_metric]]).copy()
    if len(df_sc) == 0:
        st.info("Nessun punto disponibile (mancano valori per le metriche selezionate).")
    else:
        fig_sc = px.scatter(
            df_sc,
            x=scatter_opts[x_metric],
            y=scatter_opts[y_metric],
            size=size_metric,
            hover_name=dim_label,
            title=f"{y_metric} vs {x_metric} (bolla={size_metric})"
        )
        fig_sc.update_traces(marker=dict(color="red", line=dict(color="black", width=1)))
        st.plotly_chart(fig_sc, use_container_width=True)

    st.divider()

    # Bar chart Top N
    c1, c2 = st.columns(2)
    df_top = df_dim.sort_values("ASSENTEISMO_%", ascending=False).head(top_n)
    fig1 = px.bar(df_top, x=dim_label, y="ASSENTEISMO_%", title=f"Assenteismo % – Top {top_n} {dim_label}")
    style_red_black(fig1); fig1.update_layout(xaxis_tickangle=45, yaxis_title="%")
    c1.plotly_chart(fig1, use_container_width=True)

    df_top2 = df_dim.sort_values("FTE_PERSI_ASSENZE", ascending=False).head(top_n)
    fig2 = px.bar(df_top2, x=dim_label, y="FTE_PERSI_ASSENZE", title=f"FTE persi per assenze – Top {top_n} {dim_label}")
    style_red_black(fig2); fig2.update_layout(xaxis_tickangle=45, yaxis_title="FTE")
    c2.plotly_chart(fig2, use_container_width=True)

    st.divider()

    # Top causali
    st.subheader("Top causali assenza (impatto su organico)")
    if k_global["df_break"] is not None and not k_global["df_break"].empty:
        df_caus = k_global["df_break"].head(15).copy()
        fig_caus = px.bar(
            df_caus.sort_values("%_su_FTE", ascending=True),
            x="%_su_FTE", y="Causale", orientation="h",
            title="Top causali – % riduzione organico (su FTE)"
        )
        style_red_black(fig_caus)
        st.plotly_chart(fig_caus, use_container_width=True)
    else:
        st.info("Causali non disponibili.")

# =========================
# TAB 3: Dettaglio reparto + persone
# =========================
with tab3:
    st.subheader("Dettaglio Reparto & Persone")

    if len(df_scope) == 0:
        st.warning("Nessun dato disponibile con i filtri attuali.")
        st.stop()

    # Selezione reparto (o fallback servizio)
    col_rep_scope = find_col(df_scope, ["REPARTO"], contains=True)
    dim_label = "REPARTO" if col_rep_scope else "SERVIZIO"
    dim_col = col_rep_scope if col_rep_scope else "SERVIZIO"

    dim_values = sorted(df_scope[dim_col].dropna().astype(str).unique())
    chosen_dim = st.selectbox(f"Seleziona {dim_label}", dim_values, index=0)

    df_rep = df_scope[df_scope[dim_col].astype(str) == str(chosen_dim)].copy()
    k_rep = compute_kpi(df_rep, day_hours=day_hours, ore_annue_fte=ore_annue_fte, cause_cols=CAUSE_COLS)

    st.markdown(f"### KPI – {dim_label}: **{chosen_dim}**")
    with st.container(border=True):
        a = st.columns(4)
        a[0].metric("Operatori", f"{k_rep['n_operatori']}")
        a[1].metric("FTE", f"{k_rep['fte_tot']:.2f}")
        a[2].metric("Assenteismo %", f"{k_rep['ass_pct']:.2f}%" if not np.isnan(k_rep["ass_pct"]) else "n/d")
        a[3].metric("FTE assenti", f"{k_rep['fte_assenti']:.2f}" if not np.isnan(k_rep["fte_assenti"]) else "n/d")

        b = st.columns(4)
        b[0].metric("FTE disponibili", f"{k_rep['fte_disp']:.2f}" if not np.isnan(k_rep["fte_disp"]) else "n/d")
        b[1].metric("Straordinario ore/FTE", f"{k_rep['st_x_fte']:.2f}" if not np.isnan(k_rep["st_x_fte"]) else "n/d")
        b[2].metric("Prestazioni agg ore/FTE", f"{k_rep['prest_x_fte']:.2f}" if not np.isnan(k_rep["prest_x_fte"]) else "n/d")
        b[3].metric("Residuo ferie medio (gg/op)", f"{k_rep['res_giorni_media']:.2f}" if not np.isnan(k_rep["res_giorni_media"]) else "n/d")

    st.divider()

    # Tabella persone
    st.markdown("### Persone (indicatori individuali)")
    df_people = build_people_table(df_rep, ore_annue_fte=ore_annue_fte, day_hours=day_hours, cause_cols=CAUSE_COLS)

    if df_people.empty:
        st.info("Non riesco a costruire la tabella persone (manca MATRICOLA o anagrafica).")
        st.stop()

    # controlli
    cX, cY = st.columns(2)
    sort_by = cX.selectbox(
        "Ordina per",
        ["ASSENTEISMO_%", "ASSENZE_ORE", "STRAORD_TOT_ORE", "FERIE_RES_GIORNI", "FTE_DISPONIBILI", "FTE"],
        index=0
    )
    top_people = cY.slider("Mostra top N persone", 10, 300, 50, step=10)

    df_show = df_people.sort_values(sort_by, ascending=False).head(top_people)
    st.dataframe(df_show, use_container_width=True, height=420)

    # download
    st.download_button(
        "⬇️ Scarica persone (CSV)",
        data=df_people.to_csv(index=False).encode("utf-8"),
        file_name=f"persone_{dim_label.lower()}_{str(chosen_dim)[:40]}.csv".replace(" ", "_"),
        mime="text/csv",
    )
    bufp = io.BytesIO()
    with pd.ExcelWriter(bufp, engine="openpyxl") as writer:
        df_people.to_excel(writer, index=False, sheet_name="PERSONE")
    st.download_button(
        "⬇️ Scarica persone (Excel)",
        data=bufp.getvalue(),
        file_name=f"persone_{dim_label.lower()}_{str(chosen_dim)[:40]}.xlsx".replace(" ", "_"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    # Grafici per persona (Top N)
    st.markdown("### Grafici per persona (Top N)")
    g1, g2 = st.columns(2)

    fig_abs = px.bar(
        df_show.sort_values("ASSENZE_ORE", ascending=False).head(25),
        x="PERSONA",
        y="ASSENZE_ORE",
        title="Top 25 persone – Assenze (ore)"
    )
    style_red_black(fig_abs)
    fig_abs.update_layout(xaxis_tickangle=45)
    g1.plotly_chart(fig_abs, use_container_width=True)

    fig_asspct = px.bar(
        df_show.sort_values("ASSENTEISMO_%", ascending=False).head(25),
        x="PERSONA",
        y="ASSENTEISMO_%",
        title="Top 25 persone – Assenteismo % (ore assenza / ore teoriche)"
    )
    style_red_black(fig_asspct)
    fig_asspct.update_layout(xaxis_tickangle=45, yaxis_title="%")
    g2.plotly_chart(fig_asspct, use_container_width=True)

    # Top causali nel reparto
    st.markdown("### Top causali nel reparto")
    if k_rep["df_break"] is not None and not k_rep["df_break"].empty:
        df_caus_rep = k_rep["df_break"].head(12)
        fig_caus_rep = px.bar(
            df_caus_rep.sort_values("%_su_FTE", ascending=True),
            x="%_su_FTE",
            y="Causale",
            orientation="h",
            title="Top causali – % riduzione organico (su FTE) nel reparto"
        )
        style_red_black(fig_caus_rep)
        st.plotly_chart(fig_caus_rep, use_container_width=True)
    else:
        st.info("Breakdown causali non disponibile per questo reparto.")

