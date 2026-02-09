# app.py
import io
import re
import datetime as dt
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


def build_detail_and_analisi(
    df_raw: pd.DataFrame,
    only_in_force: bool,
    cess_cutoff: dt.date | None,
    ore_annue_fte: float,
):
    """
    Costruisce:
    - df_scope: dataset pulito + colonne derivate
    - agg: tabella ANALISI_DOTAZIONI (stile PDF), aggregata per UUOO/SERVIZIO e QUALIFICA

    Regole chiave richieste:
    - FERIE: SEMPRE in giorni (nessuna conversione)
    - Assenze: conteggiate come somma di 2 blocchi:
        * mal/104/ecc
        * asp/grav/puer/dist
    - Straordinario (ore/FTE): include anche FESTIVI (pagato + recupero)
        """
    df = df_raw.copy()

    c_pt = find_col(df, ["% PART-TIME", "% PART TIME", "PART-TIME", "PERC PART"], contains=True)
    c_rep = find_col(df, ["REPARTO"], contains=True)
    c_cdr = find_col(df, ["CDR_DESC", "CDR DESC", "CDR"], contains=True)
    c_matr = find_col(df, ["MATRICOLA"], contains=True)
    c_data_al = find_col(df, ["DATA AL"], contains=True)
    c_qual = find_col(df, ["QUALIFICA.1", "QUALIFICA"], contains=True)

    # rimuove eventuali righe di totale (matricola vuota) che falsano le somme
    if c_matr and c_matr in df.columns:
        df = df[df[c_matr].notna()].copy()

    def col_or_zero(names):
        c = find_col(df, names, contains=True)
        return to_num_series(df[c]) if c else z(df)

    # -------------------------
    # FERIE (SEMPRE in GIORNI)
    # -------------------------
    ferie = col_or_zero(["FERIE"])
    ferie_rx = col_or_zero(["FERIE RX"])
    ferie_god_tot = col_or_zero(["FERIE GODUTE TOTALE"])  # nel file è già TOTALE (incl. RX + AP)
    ferie_res = col_or_zero(["FERIE RES."])
    ferie_rx_res = col_or_zero(["FERIE RX RES."])
    ferie_ap_res = col_or_zero(["FERIE AP RES."])

    # -------------------------
    # ASSENZE (ORE) - solo blocchi richiesti
    # -------------------------
    abs_malattia = col_or_zero(["MALATTIA"])
    abs_malfiglio = col_or_zero(["MALATTIA FIGLIO"])
    abs_104 = col_or_zero(["LEGGE 104"])
    abs_permessi = col_or_zero(["PERMESSI"])
    abs_aggior = col_or_zero(["AGGIOR."])
    abs_inf_malserv = col_or_zero(["INF./MAL.SERV"])
    abs_car_pub = col_or_zero(["CAR.PUBBLICA"])
    abs_covid = col_or_zero(["INFORTUNIO COVID"]) + col_or_zero(["MALATTIA COVID"])

    abs_aspettativa = col_or_zero(["ASPETTATIVA"])
    abs_gravidanza = col_or_zero(["GRAVIDANZA"]) + col_or_zero(["GRAV.OBBL."])
    abs_comando = col_or_zero(["COMANDO"])
    abs_sindacali = col_or_zero(["PERMESSI SINDACALI"]) + col_or_zero(["DISTACCO SINDACALE"])

    # -------------------------
    # STRAORDINARI / FESTIVI (ORE)
    # -------------------------
    st_rec = col_or_zero(["ORE DA RECUP. PROG."])
    st_pd = col_or_zero(["STR. PD. PROG."])
    st_pag = col_or_zero(["STR. PROG."])
    fest_pag = col_or_zero(["FEST. INFRASETT. A PAGAMENTO"])
    fest_rec = col_or_zero(["FEST. INFRASETT. A RECUPERO"])
    # -------------------------
    # FTE (da % PART-TIME)
    # -------------------------
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

    # FTE 2026: azzera chi ha DATA CESSAZIONE <= cess_cutoff (se impostato)
    c_cess = find_col(df, ["DATA CESSAZIONE", "CESSAZIONE"], contains=True)
    if c_cess and c_cess in df.columns and cess_cutoff is not None:
        cess_dt = pd.to_datetime(df[c_cess], errors="coerce").dt.date
        df["FTE_2026"] = df["FTE"].where(~(cess_dt.notna() & (cess_dt <= cess_cutoff)), 0.0)
    else:
        df["FTE_2026"] = df["FTE"]

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

    # -------------------------
    # FERIE derivate (giorni)
    # -------------------------
    df["FERIE_MAT_2025"] = ferie + ferie_rx
    df["FERIE_FRUITE_2025"] = ferie_god_tot
    df["FERIE_RES_0101"] = ferie_res + ferie_rx_res + ferie_ap_res

    # -------------------------
    # ASSENZE: blocchi richiesti (ore)
    # -------------------------
    df["MAL_104_ECC_ORE"] = (
        abs_malattia + abs_malfiglio + abs_104 +
        abs_permessi + abs_aggior + abs_inf_malserv +
        abs_car_pub + abs_covid
    )
    df["ASP_GRAV_PUER_DIST_ORE"] = (
        abs_aspettativa + abs_gravidanza + abs_comando + abs_sindacali
    )
    df["ASSENZE_TOT_ORE"] = df["MAL_104_ECC_ORE"] + df["ASP_GRAV_PUER_DIST_ORE"]

    # -------------------------
    # STRAORDINARI/FESTIVI
    # -------------------------
    df["STRAORD_REC"] = st_rec
    df["STRAORD_PD"] = st_pd
    df["STRAORD_PAG"] = st_pag
    df["FEST_PAG"] = fest_pag
    df["FEST_REC"] = fest_rec

    # scope
    df_scope = df[df["SERVIZIO"].notna()].copy()

    # solo in forza a fine periodo (DATA AL max)
    if only_in_force and c_data_al and len(df_scope) > 0:
        try:
            max_data = df_scope[c_data_al].max()
            df_scope = df_scope[df_scope[c_data_al] == max_data].copy()
        except Exception:
            pass

    # -------------------------
    # ANALISI (per servizio + qualifica)
    # -------------------------
    matr_col = c_matr if c_matr else None

    agg_cols = {
        "FTE 2025": ("FTE", "sum"),
        "FTE 2026": ("FTE_2026", "sum"),
        "st Recupero": ("STRAORD_REC", "sum"),
        "st PD pagato": ("STRAORD_PD", "sum"),
        "st Pagato": ("STRAORD_PAG", "sum"),
        "Festivo pagato": ("FEST_PAG", "sum"),
        "Festivo recupero": ("FEST_REC", "sum"),
        "Ferie maturate 2025": ("FERIE_MAT_2025", "sum"),
        "Ferie fruite 2025": ("FERIE_FRUITE_2025", "sum"),
        "Ferie residue al 01/01/2026": ("FERIE_RES_0101", "sum"),
        "Assenze totali (ore)": ("ASSENZE_TOT_ORE", "sum"),
        "mal/104/ecc (ore)": ("MAL_104_ECC_ORE", "sum"),
        "asp/grav/puer/dist (ore)": ("ASP_GRAV_PUER_DIST_ORE", "sum"),
    }

    agg = df_scope.groupby(["SERVIZIO", "QUALIFICA_OUT"], dropna=False).agg(
        OPERATORI=(matr_col, "nunique") if matr_col else ("QUALIFICA_OUT", "size"),
        **agg_cols
    ).reset_index()

    agg.rename(columns={"SERVIZIO": "UUOO/SERVIZIO", "QUALIFICA_OUT": "QUALIFICA"}, inplace=True)

    # Derivate come nel PDF
    agg["Assenze medie FTE"] = np.where(ore_annue_fte > 0, agg["Assenze totali (ore)"] / ore_annue_fte, np.nan)

    st_tot = agg["st Recupero"] + agg["st PD pagato"] + agg["st Pagato"] + agg["Festivo pagato"] + agg["Festivo recupero"]
    agg["ST media (ore/FTE)"] = np.where(agg["FTE 2025"] > 0, st_tot / agg["FTE 2025"], np.nan)

    agg["Media procapite ferie fruite 2025"] = np.where(agg["OPERATORI"] > 0, agg["Ferie fruite 2025"] / agg["OPERATORI"], np.nan)
    agg["Media procapite ferie residue al 01/01/2026"] = np.where(agg["OPERATORI"] > 0, agg["Ferie residue al 01/01/2026"] / agg["OPERATORI"], np.nan)

    ordered = [
        "UUOO/SERVIZIO", "QUALIFICA", "OPERATORI",
        "FTE 2025", "FTE 2026",
        "Assenze totali (ore)",
        "Assenze medie FTE",
        "st Recupero", "st PD pagato", "st Pagato",
        "Festivo pagato", "Festivo recupero",
        "ST media (ore/FTE)",
        "Ferie maturate 2025", "Ferie fruite 2025",
        "Media procapite ferie fruite 2025",
        "Ferie residue al 01/01/2026",
        "Media procapite ferie residue al 01/01/2026",
        "mal/104/ecc (ore)", "asp/grav/puer/dist (ore)",
    ]

    agg = agg[ordered].sort_values(["UUOO/SERVIZIO", "QUALIFICA"]).reset_index(drop=True)
    return agg, df_scope

def compute_kpi(df_scope: pd.DataFrame, ore_annue_fte: float):
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
    ferie_res_giorni = float(df_scope["FERIE_RES_0101"].sum()) if "FERIE_RES_0101" in df_scope.columns else 0.0

    ferie_pct = (ferie_fruite / ferie_mat * 100) if ferie_mat > 0 else np.nan
    res_giorni_media = (ferie_res_giorni / n_operatori) if n_operatori > 0 else np.nan

    # Straordinario (ore/FTE) = straordinari + festivi
    needed = ["STRAORD_REC", "STRAORD_PD", "STRAORD_PAG", "FEST_PAG", "FEST_REC"]
    if all(c in df_scope.columns for c in needed):
        st_tot = float(
            df_scope["STRAORD_REC"].sum() + df_scope["STRAORD_PD"].sum() + df_scope["STRAORD_PAG"].sum() +
            df_scope["FEST_PAG"].sum() + df_scope["FEST_REC"].sum()
        )
    else:
        st_tot = 0.0
    st_x_fte = (st_tot / fte_tot) if fte_tot > 0 else np.nan

    # Breakdown (solo 2 blocchi coerenti con ASSENZE_TOT_ORE)
    breakdown = []
    if ore_annue_fte > 0 and fte_tot > 0:
        for col, label in [
            ("MAL_104_ECC_ORE", "mal/104/ecc"),
            ("ASP_GRAV_PUER_DIST_ORE", "asp/grav/puer/dist"),
        ]:
            if col in df_scope.columns:
                ore = float(df_scope[col].sum())
                fte_lost = ore / ore_annue_fte
                pct = (fte_lost / fte_tot) * 100
                breakdown.append({"Causale": label, "Ore": ore, "FTE_persi": fte_lost, "%_su_FTE": pct})

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
        "ferie_res_giorni": ferie_res_giorni,
        "res_giorni_media": res_giorni_media,
        "st_tot": st_tot,
        "st_x_fte": st_x_fte,
        "df_break": df_break,
    }

def totals_row_from_scope(df_scope: pd.DataFrame, ore_annue_fte: float):
    """Riga totali coerente con la tabella unica stile PDF."""
    c_matr = find_col(df_scope, ["MATRICOLA"], contains=True)
    n_operatori = int(df_scope[c_matr].nunique()) if c_matr and c_matr in df_scope.columns else int(len(df_scope))

    fte_2025 = float(df_scope["FTE"].sum()) if "FTE" in df_scope.columns else 0.0
    fte_2026 = float(df_scope["FTE_2026"].sum()) if "FTE_2026" in df_scope.columns else fte_2025

    st_rec = float(df_scope["STRAORD_REC"].sum()) if "STRAORD_REC" in df_scope.columns else 0.0
    st_pd = float(df_scope["STRAORD_PD"].sum()) if "STRAORD_PD" in df_scope.columns else 0.0
    st_pag = float(df_scope["STRAORD_PAG"].sum()) if "STRAORD_PAG" in df_scope.columns else 0.0

    fest_pag = float(df_scope["FEST_PAG"].sum()) if "FEST_PAG" in df_scope.columns else 0.0
    fest_rec = float(df_scope["FEST_REC"].sum()) if "FEST_REC" in df_scope.columns else 0.0

    st_media = ((st_rec + st_pd + st_pag + fest_pag + fest_rec) / fte_2025) if fte_2025 > 0 else np.nan

    ferie_mat = float(df_scope["FERIE_MAT_2025"].sum()) if "FERIE_MAT_2025" in df_scope.columns else 0.0
    ferie_fruite = float(df_scope["FERIE_FRUITE_2025"].sum()) if "FERIE_FRUITE_2025" in df_scope.columns else 0.0
    ferie_res = float(df_scope["FERIE_RES_0101"].sum()) if "FERIE_RES_0101" in df_scope.columns else 0.0

    media_fruite = (ferie_fruite / n_operatori) if n_operatori > 0 else np.nan
    media_residue = (ferie_res / n_operatori) if n_operatori > 0 else np.nan

    abs_tot = float(df_scope["ASSENZE_TOT_ORE"].sum()) if "ASSENZE_TOT_ORE" in df_scope.columns else 0.0
    assenze_medie_fte = (abs_tot / ore_annue_fte) if ore_annue_fte > 0 else np.nan

    mal_104_ecc = float(df_scope["MAL_104_ECC_ORE"].sum()) if "MAL_104_ECC_ORE" in df_scope.columns else 0.0
    asp_grav = float(df_scope["ASP_GRAV_PUER_DIST_ORE"].sum()) if "ASP_GRAV_PUER_DIST_ORE" in df_scope.columns else 0.0

    row = {
        "UUOO/SERVIZIO": "TOTALE",
        "QUALIFICA": "",
        "OPERATORI": n_operatori,
        "FTE 2025": fte_2025,
        "FTE 2026": fte_2026,
        "Assenze totali (ore)": abs_tot,
        "Assenze medie FTE": assenze_medie_fte,
        "st Recupero": st_rec,
        "st PD pagato": st_pd,
        "st Pagato": st_pag,
        "Festivo pagato": fest_pag,
        "Festivo recupero": fest_rec,
        "ST media (ore/FTE)": st_media,
        "Ferie maturate 2025": ferie_mat,
        "Ferie fruite 2025": ferie_fruite,
        "Media procapite ferie fruite 2025": media_fruite,
        "Ferie residue al 01/01/2026": ferie_res,
        "Media procapite ferie residue al 01/01/2026": media_residue,
        "mal/104/ecc (ore)": mal_104_ecc,
        "asp/grav/puer/dist (ore)": asp_grav,
    }

    return pd.DataFrame([row])

def build_tabella_dotazioni(df_scope: pd.DataFrame, ore_annue_fte: float, cess_cutoff: dt.date):
    """Tabella sintetica richiesta:
    UUOO/SERVIZIO, QUALIFICA, OPERATORI, N°FTE 2025, N°FTE 2026,
    Assenze medie FTE (= Assenze totali / ore_annue_fte),
    Straordinario (ore/FTE) (= (straordinari + festivi) / FTE).
    """
    cols_out = [
        "UUOO/SERVIZIO", "QUALIFICA", "OPERATORI",
        "N°FTE 2025", "N°FTE 2026", "Assenze medie FTE", "Straordinario (ore/FTE)"
    ]
    if df_scope is None or len(df_scope) == 0:
        return pd.DataFrame(columns=cols_out)

    df = df_scope.copy()

    c_matr = find_col(df, ["MATRICOLA"], contains=True)
    service_col = "SERVIZIO" if "SERVIZIO" in df.columns else find_col(df, ["CDR_DESC", "REPARTO"], contains=True)
    qual_col = "QUALIFICA_OUT" if "QUALIFICA_OUT" in df.columns else find_col(df, ["QUALIFICA.1", "QUALIFICA"], contains=True)

    if not service_col or not qual_col:
        return pd.DataFrame(columns=cols_out)

    # cessazioni
    c_cess = find_col(df, ["DATA CESSAZIONE", "CESSAZIONE"], contains=True)
    if c_cess and c_cess in df.columns:
        cess_dt = pd.to_datetime(df[c_cess], errors="coerce").dt.date
    else:
        cess_dt = pd.Series([pd.NaT] * len(df), index=df.index)

    df["CESSATO_AL_CUTOFF"] = False
    if cess_cutoff:
        df["CESSATO_AL_CUTOFF"] = cess_dt.notna() & (cess_dt <= cess_cutoff)

    df["FTE_2025"] = to_num_series(df["FTE"]) if "FTE" in df.columns else 0.0
    df["FTE_2026"] = df["FTE_2025"].where(~df["CESSATO_AL_CUTOFF"], 0.0)

    # straordinari + festivi (ore)
    for c in ["STRAORD_REC", "STRAORD_PD", "STRAORD_PAG", "FEST_PAG", "FEST_REC"]:
        if c not in df.columns:
            df[c] = 0.0
    df["STRAORD_TOT_ORE"] = (
        to_num_series(df["STRAORD_REC"]) + to_num_series(df["STRAORD_PD"]) + to_num_series(df["STRAORD_PAG"]) +
        to_num_series(df["FEST_PAG"]) + to_num_series(df["FEST_REC"])
    )

    # assenze (ore) - già coerenti con i blocchi richiesti
    if "ASSENZE_TOT_ORE" not in df.columns:
        df["ASSENZE_TOT_ORE"] = 0.0

    gb = df.groupby([service_col, qual_col], dropna=False).agg(
        OPERATORI=(c_matr, "nunique") if c_matr and c_matr in df.columns else (qual_col, "size"),
        **{
            "N°FTE 2025": ("FTE_2025", "sum"),
            "N°FTE 2026": ("FTE_2026", "sum"),
            "_ASSENZE_ORE": ("ASSENZE_TOT_ORE", "sum"),
            "_STRAORD_ORE": ("STRAORD_TOT_ORE", "sum"),
        }
    ).reset_index()

    gb.rename(columns={service_col: "UUOO/SERVIZIO", qual_col: "QUALIFICA"}, inplace=True)
    gb["Assenze medie FTE"] = np.where(ore_annue_fte > 0, gb["_ASSENZE_ORE"] / ore_annue_fte, np.nan)
    gb["Straordinario (ore/FTE)"] = np.where(gb["N°FTE 2025"] > 0, gb["_STRAORD_ORE"] / gb["N°FTE 2025"], np.nan)

    gb = gb[cols_out].sort_values(["UUOO/SERVIZIO", "QUALIFICA"]).reset_index(drop=True)
    return gb

def build_people_table(df_sub: pd.DataFrame, ore_annue_fte: float):
    c_matr = find_col(df_sub, ["MATRICOLA"], contains=True)
    c_cogn = find_col(df_sub, ["COGNOME"], contains=True)
    c_nome = find_col(df_sub, ["NOME"], contains=True)
    c_prof = find_col(df_sub, ["PROFILO"], contains=True)
    c_qual = find_col(df_sub, ["QUALIFICA.1", "QUALIFICA"], contains=True)

    # --- chiave persona ---
    if c_matr and c_matr in df_sub.columns:
        key_cols = [c_matr]
    else:
        key_cols = []
        if c_cogn and c_cogn in df_sub.columns:
            key_cols.append(c_cogn)
        if c_nome and c_nome in df_sub.columns:
            key_cols.append(c_nome)

    if not key_cols:
        return pd.DataFrame()

    # --- aggregazioni per persona ---
    agg_dict = {
        "FTE": ("FTE", "sum"),
        "ASSENZE_ORE": ("ASSENZE_TOT_ORE", "sum"),
        # FERIE: sempre in giorni
        "FERIE_RES_GIORNI": ("FERIE_RES_0101", "sum"),
        "FERIE_MAT_GIORNI": ("FERIE_MAT_2025", "sum"),
        "FERIE_FRUITE_GIORNI": ("FERIE_FRUITE_2025", "sum"),
        # Straordinari + festivi (ore)
        "ST_REC": ("STRAORD_REC", "sum"),
        "ST_PD": ("STRAORD_PD", "sum"),
        "ST_PAG": ("STRAORD_PAG", "sum"),
        "FEST_PAG": ("FEST_PAG", "sum"),
        "FEST_REC": ("FEST_REC", "sum"),
        # Blocchi assenze (ore)
        "mal/104/ecc (ore)": ("MAL_104_ECC_ORE", "sum"),
        "asp/grav/puer/dist (ore)": ("ASP_GRAV_PUER_DIST_ORE", "sum"),
    }

    gb = df_sub.groupby(key_cols, dropna=False).agg(**{k: v for k, v in agg_dict.items()}).reset_index()

    # --- helper: aggiunge colonna "first" solo se serve e senza collisioni ---
    def add_first(colname, outname):
        if outname in gb.columns:
            return gb
        if not colname or colname not in df_sub.columns:
            return gb
        if colname in key_cols:
            if outname != colname and colname in gb.columns and outname not in gb.columns:
                return gb.rename(columns={colname: outname})
            return gb
        firsts = (
            df_sub.groupby(key_cols, dropna=False)[colname]
            .first()
            .reset_index(name=outname)
        )
        return gb.merge(firsts, on=key_cols, how="left")

    gb = add_first(c_cogn, "COGNOME")
    gb = add_first(c_nome, "NOME")
    gb = add_first(c_prof, "PROFILO")
    gb = add_first(c_qual, "QUALIFICA_RAW")

    # --- metriche individuali ---
    gb["ORE_TEORICHE"] = gb["FTE"] * ore_annue_fte
    gb["ASSENTEISMO_%"] = np.where(gb["ORE_TEORICHE"] > 0, gb["ASSENZE_ORE"] / gb["ORE_TEORICHE"] * 100, np.nan)
    gb["FTE_ASSENTI"] = gb["ASSENZE_ORE"] / ore_annue_fte if ore_annue_fte > 0 else np.nan
    gb["FTE_DISPONIBILI"] = gb["FTE"] - gb["FTE_ASSENTI"]

    gb["STRAORD_TOT_ORE"] = gb["ST_REC"] + gb["ST_PD"] + gb["ST_PAG"] + gb["FEST_PAG"] + gb["FEST_REC"]
    gb["STRAORD_ORE_X_FTE"] = np.where(gb["FTE"] > 0, gb["STRAORD_TOT_ORE"] / gb["FTE"], np.nan)

    gb["FERIE_RES_GIORNI_X_FTE"] = np.where(gb["FTE"] > 0, gb["FERIE_RES_GIORNI"] / gb["FTE"], np.nan)

    # --- display PERSONA ---
    if "COGNOME" in gb.columns and "NOME" in gb.columns:
        gb["PERSONA"] = gb["COGNOME"].astype(str).str.strip() + " " + gb["NOME"].astype(str).str.strip()
    elif c_matr and c_matr in gb.columns:
        gb["PERSONA"] = gb[c_matr].astype(str)
    else:
        gb["PERSONA"] = "N/D"

    # ordine colonne
    cols_front = []
    if c_matr and c_matr in gb.columns:
        cols_front.append(c_matr)
    cols_front += [c for c in ["PERSONA", "COGNOME", "NOME", "PROFILO", "QUALIFICA_RAW"] if c in gb.columns]

    cols_metrics = [
        "FTE", "ORE_TEORICHE", "ASSENZE_ORE", "ASSENTEISMO_%",
        "FTE_ASSENTI", "FTE_DISPONIBILI",
        "STRAORD_TOT_ORE", "STRAORD_ORE_X_FTE",
        "ST_REC", "ST_PD", "ST_PAG", "FEST_PAG", "FEST_REC",
        "FERIE_RES_GIORNI", "FERIE_RES_GIORNI_X_FTE",
        "FERIE_MAT_GIORNI", "FERIE_FRUITE_GIORNI",
        "mal/104/ecc (ore)", "asp/grav/puer/dist (ore)",
    ]

    cols_metrics = [c for c in cols_metrics if c in gb.columns]
    other_cols = [c for c in gb.columns if c not in cols_front + cols_metrics]

    return gb[cols_front + cols_metrics + other_cols]

# =========================
# UI
# =========================
st.title("Cruscotto Dotazioni Organiche")

with st.sidebar:
    st.header("📤 Caricamento")
    uploaded = st.file_uploader("Carica dotazioni.xlsx (xlsx)", type=["xlsx"])

    st.divider()
    st.header("⚙️ Parametri calcolo")
    ore_annue_fte = st.number_input("Ore teoriche annue per 1 FTE", min_value=800.0, max_value=2200.0, value=1470.0, step=10.0)

    cess_cutoff = st.date_input("Data riferimento cessazioni (FTE 2026)", value=dt.date.today())

    st.divider()
    st.header("⚙️ Opzioni")
    only_in_force = st.toggle("Solo in forza a fine periodo (DATA AL max)", value=False)
    use_stab_cdc = st.toggle("Usa mappa Stabilimento (CDC→Stabilimento)", value=True)

if not uploaded:
    st.info("Carica un file Excel dalla sidebar per iniziare.")
    st.stop()

file_bytes = uploaded.getvalue()

# lettura excel smart + override
df_raw, meta, sheet_names = load_excel_smart(file_bytes)

with st.sidebar:
    st.divider()
    st.header("Lettura Excel")
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

    # check colonne essenziali (best effort)
    essential = [
        "CDR_DESC", "REPARTO", "MATRICOLA", "% PART-TIME", "DATA AL",
        "DATA CESSAZIONE", "ORE DA RECUP. PROG.", "STR. PD. PROG.", "STR. PROG.",
        "MALATTIA", "PERMESSI", "LEGGE 104", "MALATTIA FIGLIO", "GRAVIDANZA",
    ]
    missing = [c for c in essential if find_col(df_raw, [c], contains=False) is None]
    st.write(f"Numero colonne: {len(df_raw.columns)}")
    if missing:
        st.warning(f"Colonne essenziali non trovate (controlla header): {missing}")
    else:
        st.success("Tutte le colonne essenziali sono presenti.")
    st.dataframe(df_raw.head(15), use_container_width=True)


# ---- Filtri in sidebar ----
col_dip = find_col(df_raw, ["DESC. DIP.", "DESC DIP"], contains=True)
col_stab = find_col(df_raw, ["STABILIMENTO"], contains=True)
col_cdr = find_col(df_raw, ["CDR_DESC", "CDR DESC", "CDR"], contains=True)
col_rep = find_col(df_raw, ["REPARTO"], contains=True)
col_prof = find_col(df_raw, ["PROFILO"], contains=True)
col_qual = find_col(df_raw, ["QUALIFICA.1", "QUALIFICA"], contains=True)
col_ruolo = find_col(df_raw, ["DESC. RUOLO", "RUOLO"], contains=True)

# ---- DAPSS: mappa CDC (COD.REP.) -> Area Funzionale ----
col_cdc = find_col(df_raw, ["COD.REP.", "COD.REP", "COD REP", "CDC", "CODICE REPARTO"], contains=True)

def _cdc_to_int(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, np.integer)):
        return int(x)
    if isinstance(x, float) and not np.isnan(x):
        return int(x)
    s = str(x).strip()
    m = re.match(r"^0*(\d+)", s)
    return int(m.group(1)) if m else np.nan

DAPSS_DEFAULT = {
    # AREA MEDICO - ONCOLOGICA
    105: "AREA MEDICO - ONCOLOGICA", 134: "AREA MEDICO - ONCOLOGICA", 135: "AREA MEDICO - ONCOLOGICA",
    152: "AREA MEDICO - ONCOLOGICA", 161: "AREA MEDICO - ONCOLOGICA", 162: "AREA MEDICO - ONCOLOGICA",
    173: "AREA MEDICO - ONCOLOGICA", 175: "AREA MEDICO - ONCOLOGICA", 176: "AREA MEDICO - ONCOLOGICA",
    189: "AREA MEDICO - ONCOLOGICA", 191: "AREA MEDICO - ONCOLOGICA", 193: "AREA MEDICO - ONCOLOGICA",
    197: "AREA MEDICO - ONCOLOGICA", 201: "AREA MEDICO - ONCOLOGICA", 202: "AREA MEDICO - ONCOLOGICA",
    231: "AREA MEDICO - ONCOLOGICA", 232: "AREA MEDICO - ONCOLOGICA", 233: "AREA MEDICO - ONCOLOGICA",
    234: "AREA MEDICO - ONCOLOGICA", 236: "AREA MEDICO - ONCOLOGICA", 326: "AREA MEDICO - ONCOLOGICA",
    324: "AREA MEDICO - ONCOLOGICA", 2700: "AREA MEDICO - ONCOLOGICA",

    # AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO
    120: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO", 121: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO",
    122: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO", 181: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO",
    182: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO", 221: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO",
    260: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO", 261: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO",
    292: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO", 300: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO",
    312: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO", 361: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO",
    362: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO", 412: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO",
    430: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO", 437: "AREA GESTIONE RISORSE INFERMIERISTICHE E DI SUPPORTO OGLIO PO",

    # SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI
    204: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 422: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    426: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 442: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    450: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 451: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    460: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 461: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    462: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 465: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    470: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 473: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    480: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 482: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    483: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 484: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    488: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 490: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    493: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 500: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    510: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 560: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",
    570: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI", 585: "SERVIZI DIAGNOSTICI - AMBULATORIALI - RIABILITATIVI",

    # AREA CHIRURGICA
    116: "AREA CHIRURGICA", 282: "AREA CHIRURGICA", 315: "AREA CHIRURGICA", 320: "AREA CHIRURGICA",
    321: "AREA CHIRURGICA", 329: "AREA CHIRURGICA", 331: "AREA CHIRURGICA", 340: "AREA CHIRURGICA",
    341: "AREA CHIRURGICA", 351: "AREA CHIRURGICA", 354: "AREA CHIRURGICA", 392: "AREA CHIRURGICA",
    393: "AREA CHIRURGICA", 505: "AREA CHIRURGICA", 582: "AREA CHIRURGICA", 586: "AREA CHIRURGICA",
    633: "AREA CHIRURGICA",

    # AREA MATERNO-INFANTILE
    211: "AREA MATERNO-INFANTILE", 212: "AREA MATERNO-INFANTILE", 214: "AREA MATERNO-INFANTILE",
    215: "AREA MATERNO-INFANTILE", 219: "AREA MATERNO-INFANTILE", 401: "AREA MATERNO-INFANTILE",
    402: "AREA MATERNO-INFANTILE", 403: "AREA MATERNO-INFANTILE",

    # AREA DELL'EMERGENZA - URGENZA
    251: "AREA DELL'EMERGENZA - URGENZA", 255: "AREA DELL'EMERGENZA - URGENZA", 271: "AREA DELL'EMERGENZA - URGENZA",

    # AAT 118
    263: "AAT 118", 274: "AAT 118",

    # AREA DELLA SALUTE MENTALE E DIPENDENZE
    521: "AREA DELLA SALUTE MENTALE E DIPENDENZE", 522: "AREA DELLA SALUTE MENTALE E DIPENDENZE",
    523: "AREA DELLA SALUTE MENTALE E DIPENDENZE", 525: "AREA DELLA SALUTE MENTALE E DIPENDENZE",
    526: "AREA DELLA SALUTE MENTALE E DIPENDENZE", 527: "AREA DELLA SALUTE MENTALE E DIPENDENZE",
    528: "AREA DELLA SALUTE MENTALE E DIPENDENZE", 529: "AREA DELLA SALUTE MENTALE E DIPENDENZE",
    531: "AREA DELLA SALUTE MENTALE E DIPENDENZE", 532: "AREA DELLA SALUTE MENTALE E DIPENDENZE",
    534: "AREA DELLA SALUTE MENTALE E DIPENDENZE", 538: "AREA DELLA SALUTE MENTALE E DIPENDENZE",
    543: "AREA DELLA SALUTE MENTALE E DIPENDENZE", 544: "AREA DELLA SALUTE MENTALE E DIPENDENZE",
    545: "AREA DELLA SALUTE MENTALE E DIPENDENZE", 1004: "AREA DELLA SALUTE MENTALE E DIPENDENZE",
    1005: "AREA DELLA SALUTE MENTALE E DIPENDENZE",

    # SERVIZI DISTRETTUALI E COT
    140: "SERVIZI DISTRETTUALI E COT", 626: "SERVIZI DISTRETTUALI E COT", 894: "SERVIZI DISTRETTUALI E COT",
    1008: "SERVIZI DISTRETTUALI E COT", 1009: "SERVIZI DISTRETTUALI E COT", 1011: "SERVIZI DISTRETTUALI E COT",
    1013: "SERVIZI DISTRETTUALI E COT", 1021: "SERVIZI DISTRETTUALI E COT", 1022: "SERVIZI DISTRETTUALI E COT",
    1034: "SERVIZI DISTRETTUALI E COT", 1035: "SERVIZI DISTRETTUALI E COT", 1200: "SERVIZI DISTRETTUALI E COT",
    1203: "SERVIZI DISTRETTUALI E COT", 1204: "SERVIZI DISTRETTUALI E COT",
    2804: "SERVIZI DISTRETTUALI E COT", 2805: "SERVIZI DISTRETTUALI E COT", 2810: "SERVIZI DISTRETTUALI E COT",
    2814: "SERVIZI DISTRETTUALI E COT",
}

def parse_dapss_override(text: str) -> dict[int, str]:
    mapping = {}
    for line in (text or "").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        # accetta: AREA<TAB>cdc1,c2,c3 oppure AREA;cdc1,c2
        parts = re.split(r"\t|;", line, maxsplit=1)
        if len(parts) < 2:
            continue
        area = parts[0].strip()
        codes_part = parts[1].strip()
        codes = re.split(r"[ ,]+", codes_part)
        for c in codes:
            c = c.strip()
            if not c:
                continue
            m = re.match(r"^0*(\d+)", c)
            if m:
                mapping[int(m.group(1))] = area
    return mapping

# costruisce colonna DAPSS_AREA (se possibile)
if col_cdc:
    df_raw["_CDC_CODE"] = df_raw[col_cdc].apply(_cdc_to_int)
    df_raw["DAPSS_AREA"] = df_raw["_CDC_CODE"].map(DAPSS_DEFAULT).fillna("NON MAPPATO")

    # ---- Stabilimento: override tramite mappa CDC (opzionale) ----
    CDC_STAB_MAP = {
        # CREMONA
        140: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        503: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        528: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        529: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        531: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        532: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        534: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        538: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        543: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        545: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        596: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        626: "AREA SOCIOSANITARIA TERRITORIALE CREMONA",
        # CASALMAGGIORE
        521: "AREA SOCIOSANITARIA CASALMAGGIORE",
        522: "AREA SOCIOSANITARIA CASALMAGGIORE",
        523: "AREA SOCIOSANITARIA CASALMAGGIORE",
        525: "AREA SOCIOSANITARIA CASALMAGGIORE",
        544: "AREA SOCIOSANITARIA CASALMAGGIORE",
    }

    col_stab_base = col_stab
    if col_stab_base:
        df_raw["STABILIMENTO_DA_CDC"] = df_raw[col_stab_base].astype(str)
        df_raw.loc[df_raw["_CDC_CODE"].isin(CDC_STAB_MAP.keys()), "STABILIMENTO_DA_CDC"] = df_raw["_CDC_CODE"].map(CDC_STAB_MAP)



# colonna Stabilimento effettiva (eventuale override CDC→Stabilimento)
col_stab_used = "STABILIMENTO_DA_CDC" if (use_stab_cdc and "STABILIMENTO_DA_CDC" in df_raw.columns) else col_stab


with st.sidebar:
    st.divider()
    st.header("Filtri")

    st.caption("Scegli se filtrare per Dipartimento (DESC. DIP) oppure per Area funzionale DAPSS (mappa da CDC/COD.REP.).")

    def opts(col):
        if not col:
            return []
        # col può essere nome colonna (str) oppure variabile già risolta
        c = col if isinstance(col, str) else col
        if isinstance(c, str) and c in df_raw.columns:
            return sorted(df_raw[c].dropna().astype(str).unique())
        return []

    dip_opts = opts(col_dip)
    stab_opts = opts(col_stab_used)
    cdr_opts = opts(col_cdr)
    rep_opts = opts(col_rep)
    prof_opts = opts(col_prof)
    qual_opts = opts(col_qual)
    ruolo_opts = opts(col_ruolo)
    dapss_opts = opts("DAPSS_AREA") if "DAPSS_AREA" in df_raw.columns else []

    filtro_org = st.sidebar.radio(
        "Filtro organizzativo principale",
        ["Dipartimento (DESC. DIP)", "Area funzionale DAPSS (CDC)"],
        index=0,
        key="filtro_org",
    )

    # Dipartimento vs DAPSS
    if filtro_org.startswith("Dipartimento"):
        dip_sel = st.sidebar.multiselect("Dipartimento", dip_opts, default=dip_opts) if dip_opts else []
        dapss_sel = []
    else:
        dip_sel = dip_opts  # non filtrare per dipartimento
        # override facoltativo della mappa CDC->DAPSS
        with st.sidebar.expander("🗺️ Mappa CDC → Area DAPSS (override facoltativo)"):
            st.caption("Formato: AREA<TAB>cdc1,c2,c3 (una riga per area). Esempio: AREA CHIRURGICA\t116,320,321")
            override_txt = st.text_area("Override mappa", value="", height=110)
            override_map = parse_dapss_override(override_txt)
            if override_map and "_CDC_CODE" in df_raw.columns:
                _map = {**DAPSS_DEFAULT, **override_map}
                df_raw["DAPSS_AREA"] = df_raw["_CDC_CODE"].map(_map).fillna("NON MAPPATO")
                dapss_opts = sorted(df_raw["DAPSS_AREA"].dropna().astype(str).unique())

        default_dapss = [x for x in dapss_opts if x != "NON MAPPATO"]
        dapss_sel = st.sidebar.multiselect("Area DAPSS", dapss_opts, default=default_dapss) if dapss_opts else []

        if "_CDC_CODE" in df_raw.columns and "DAPSS_AREA" in df_raw.columns:
            unmapped = sorted(df_raw.loc[df_raw["DAPSS_AREA"] == "NON MAPPATO", "_CDC_CODE"].dropna().astype(int).unique().tolist())
            if len(unmapped) > 0:
                st.sidebar.info(f"CDC non mappati: {len(unmapped)} (es. {', '.join(map(str, unmapped[:10]))}{'...' if len(unmapped) > 10 else ''})")

    # altri filtri
    stab_sel = st.sidebar.multiselect("Stabilimento", stab_opts, default=stab_opts) if stab_opts else []
    cdr_sel = st.sidebar.multiselect("CDR_DESC", cdr_opts, default=cdr_opts) if cdr_opts else []
    rep_sel = st.sidebar.multiselect("Reparto", rep_opts, default=rep_opts) if rep_opts else []

    st.markdown("**Filtri professionali**")
    prof_sel = st.sidebar.multiselect("Profilo", prof_opts, default=prof_opts) if prof_opts else []
    qual_sel = st.sidebar.multiselect("Qualifica", qual_opts, default=qual_opts) if qual_opts else []
    ruolo_sel = st.sidebar.multiselect("Ruolo", ruolo_opts, default=ruolo_opts) if ruolo_opts else []
# applica filtri
df_f = df_raw.copy()
# filtro organizzativo principale
if filtro_org.startswith("Dipartimento"):
    if col_dip and dip_sel:
        df_f = df_f[df_f[col_dip].astype(str).isin(dip_sel)]
else:
    if "DAPSS_AREA" in df_f.columns and dapss_sel:
        df_f = df_f[df_f["DAPSS_AREA"].astype(str).isin(dapss_sel)]
if col_stab_used and stab_sel:
    df_f = df_f[df_f[col_stab_used].astype(str).isin(stab_sel)]
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
analisi, df_scope = build_detail_and_analisi(
    df_f,
    only_in_force=only_in_force,
    cess_cutoff=cess_cutoff,
    ore_annue_fte=ore_annue_fte,
)
# KPI globali
k_global = compute_kpi(df_scope, ore_annue_fte=ore_annue_fte)

st.subheader("KPI (aggiornati dai filtri)")
with st.container(border=True):
    r1 = st.columns(4)
    r1[0].metric("N Operatori", f"{k_global['n_operatori']}")
    r1[1].metric("FTE totali", f"{k_global['fte_tot']:.2f}")
    r1[2].metric("Assenteismo % (su 1470h/FTE)", f"{k_global['ass_pct']:.2f}%" if not np.isnan(k_global["ass_pct"]) else "n/d")
    r1[3].metric("FTE mediamente assenti", f"{k_global['fte_assenti']:.2f}" if not np.isnan(k_global["fte_assenti"]) else "n/d")

    r2 = st.columns(3)
    r2[0].metric("FTE disponibili", f"{k_global['fte_disp']:.2f}" if not np.isnan(k_global["fte_disp"]) else "n/d")
    r2[1].metric("Straordinario (ore/FTE)", f"{k_global['st_x_fte']:.2f}" if not np.isnan(k_global["st_x_fte"]) else "n/d")
    r2[2].metric("Residuo ferie medio (gg/op)", f"{k_global['res_giorni_media']:.2f}" if not np.isnan(k_global["res_giorni_media"]) else "n/d")

st.divider()

tab1, tab2, tab3 = st.tabs([
    "Analisi Dotazioni",
    "Vista Generale",
    "Dettaglio Reparto"
])

# =========================
# TAB 1
# =========================
with tab1:
    st.subheader("Tabella ANALISI_DOTAZIONI (unica – come report PDF)")
    st.caption("La tabella include FTE 2025/2026 (con cessazioni), assenze, ferie e straordinari/festivi, con riga TOTALE in fondo.")

    df_total = totals_row_from_scope(df_scope, ore_annue_fte=ore_annue_fte)
    analisi_show = pd.concat([analisi, df_total], ignore_index=True)

    # arrotonda solo le colonne numeriche (senza toccare i testi)
    num_cols = analisi_show.select_dtypes(include=[np.number]).columns
    analisi_show[num_cols] = analisi_show[num_cols].round(2)

    st.dataframe(analisi_show, use_container_width=True, height=640)

    st.download_button(
        "Scarica tabella (CSV)",
        data=analisi_show.to_csv(index=False).encode("utf-8"),
        file_name="ANALISI_DOTAZIONI_unica.csv",
        mime="text/csv",
    )

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        analisi_show.to_excel(writer, index=False, sheet_name="ANALISI_DOTAZIONI")
    st.download_button(
        "Scarica tabella (Excel)",
        data=buf.getvalue(),
        file_name="ANALISI_DOTAZIONI_unica.xlsx",
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


    agg_map = {
        "OPERATORI": (c_matr, "nunique") if c_matr and c_matr in df_scope.columns else ("QUALIFICA_OUT", "size"),
        "FTE": ("FTE", "sum"),
        "ASSENZE_ORE": ("ASSENZE_TOT_ORE", "sum"),
        "FERIE_RES_GIORNI": ("FERIE_RES_0101", "sum"),
        "ST_REC": ("STRAORD_REC", "sum"),
        "ST_PD": ("STRAORD_PD", "sum"),
        "ST_PAG": ("STRAORD_PAG", "sum"),
        "FEST_PAG": ("FEST_PAG", "sum"),
        "FEST_REC": ("FEST_REC", "sum"),
    }

    df_dim = (
        df_scope.groupby(dim_col)
        .agg(**agg_map)
        .reset_index()
        .rename(columns={dim_col: dim_label})
    )

    df_dim["ORE_TEORICHE"] = df_dim["FTE"] * ore_annue_fte
    df_dim["ASSENTEISMO_%"] = np.where(df_dim["ORE_TEORICHE"] > 0, df_dim["ASSENZE_ORE"] / df_dim["ORE_TEORICHE"] * 100, np.nan)
    df_dim["FTE_PERSI_ASSENZE"] = df_dim["ASSENZE_ORE"] / ore_annue_fte

    df_dim["STRAORD_TOT_ORE"] = df_dim["ST_REC"] + df_dim["ST_PD"] + df_dim["ST_PAG"] + df_dim["FEST_PAG"] + df_dim["FEST_REC"]
    df_dim["STRAORD_ORE_X_FTE"] = np.where(df_dim["FTE"] > 0, df_dim["STRAORD_TOT_ORE"] / df_dim["FTE"], np.nan)


    df_dim["FERIE_RES_GIORNI_X_TESTA"] = np.where(df_dim["OPERATORI"] > 0, df_dim["FERIE_RES_GIORNI"] / df_dim["OPERATORI"], np.nan)


    # Scatter
    st.subheader("Scatter (interattivo)")
    scatter_opts = {
        "Assenteismo %": "ASSENTEISMO_%",
        "Straordinario (ore/FTE)": "STRAORD_ORE_X_FTE",
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
    chosen_dims = st.multiselect(
        f"Seleziona uno o più {dim_label}",
        dim_values,
        default=[dim_values[0]] if dim_values else []
    )

    if not chosen_dims:
        st.info(f"Seleziona almeno un {dim_label} per vedere il dettaglio.")
        st.stop()

    chosen_label_short = ", ".join(chosen_dims[:3]) + (" …" if len(chosen_dims) > 3 else "")
    df_rep = df_scope[df_scope[dim_col].astype(str).isin([str(x) for x in chosen_dims])].copy()
    k_rep = compute_kpi(df_rep, ore_annue_fte=ore_annue_fte)

    st.markdown(f"### Indicatori di composizione – {dim_label}: **{chosen_label_short}**")

    # -------------------------
    # Deltas "interni" al reparto
    # -------------------------
    teste = int(k_rep.get("n_operatori", 0))
    fte = float(k_rep.get("fte_tot", 0.0))

    delta_teste_fte = (teste - fte) if fte is not None else np.nan
    ratio_teste_fte = (teste / fte) if fte > 0 else np.nan

    # Ore lavorate vs ore teoriche (ore)
    col_ore_lav = find_col(df_rep, ["ORE LAVORATE"], contains=True)
    col_ore_teo_file = find_col(df_rep, ["ORE TEORICHE"], contains=True)

    ore_lav = float(to_num_series(df_rep[col_ore_lav]).sum()) if col_ore_lav else np.nan
    ore_teo_file = float(to_num_series(df_rep[col_ore_teo_file]).sum()) if col_ore_teo_file else np.nan
    ore_teo = ore_teo_file if (col_ore_teo_file and ore_teo_file > 0) else float(fte * ore_annue_fte)

    ore_gap = (ore_lav - ore_teo) if (not np.isnan(ore_lav) and ore_teo > 0) else np.nan
    copertura = (ore_lav / ore_teo * 100) if (not np.isnan(ore_lav) and ore_teo > 0) else np.nan

    # Ferie (giorni) – sempre in giorni
    ferie_res = float(to_num_series(df_rep["FERIE_RES_0101"]).sum()) if "FERIE_RES_0101" in df_rep.columns else np.nan
    ferie_mat = float(to_num_series(df_rep["FERIE_MAT_2025"]).sum()) if "FERIE_MAT_2025" in df_rep.columns else np.nan
    ferie_fruite = float(to_num_series(df_rep["FERIE_FRUITE_2025"]).sum()) if "FERIE_FRUITE_2025" in df_rep.columns else np.nan

    saldo_ferie = (ferie_mat - ferie_fruite) if (not np.isnan(ferie_mat) and not np.isnan(ferie_fruite)) else np.nan
    ferie_da_fruire = max(saldo_ferie, 0.0) if not np.isnan(saldo_ferie) else np.nan
    ferie_res_media = (ferie_res / teste) if (teste > 0 and not np.isnan(ferie_res)) else np.nan

    # Mostra KPI principali richiesti (senza box)
    rA = st.columns(4)
    rA[0].metric(
        "Teste (n)",
        f"{teste}",
        delta=(f"{delta_teste_fte:+.2f} vs FTE" if not np.isnan(delta_teste_fte) else ""),
        delta_color="off",
    )
    rA[1].metric("FTE", f"{fte:.2f}")
    rA[2].metric(
        "Teste/FTE",
        f"{ratio_teste_fte:.2f}" if not np.isnan(ratio_teste_fte) else "n/d",
        delta=(f"{(ratio_teste_fte - 1) * 100:+.1f}% vs 1.00" if not np.isnan(ratio_teste_fte) else ""),
        delta_color="off",
    )
    rA[3].metric(
        "FTE disponibili",
        f"{k_rep['fte_disp']:.2f}" if not np.isnan(k_rep.get("fte_disp", np.nan)) else "n/d",
    )

    rB = st.columns(4)
    rB[0].metric("Ore teoriche (h)", f"{ore_teo:.0f}" if not np.isnan(ore_teo) else "n/d")
    rB[1].metric(
        "Ore lavorate (h)",
        f"{ore_lav:.0f}" if not np.isnan(ore_lav) else "n/d",
        delta=(f"{ore_gap:+.0f} vs teoriche" if not np.isnan(ore_gap) else ""),
        delta_color="off",
    )
    rB[2].metric(
        "Copertura ore (%)",
        f"{copertura:.1f}%" if not np.isnan(copertura) else "n/d",
        delta=(f"{(copertura - 100):+.1f} pp vs 100%" if not np.isnan(copertura) else ""),
        delta_color="off",
    )
    rB[3].metric(
        "Straordinario+Festivi (h/FTE)",
        f"{k_rep['st_x_fte']:.2f}" if not np.isnan(k_rep.get("st_x_fte", np.nan)) else "n/d",
    )

    rC = st.columns(4)
    rC[0].metric(
        "Residuo ferie al 01/01/2026 (gg)",
        f"{ferie_res:.0f}" if not np.isnan(ferie_res) else "n/d",
        delta=(f"{ferie_res:+.0f} vs 0" if not np.isnan(ferie_res) else ""),
        delta_color="inverse",
    )
    rC[1].metric(
        "Residuo ferie medio (gg/op)",
        f"{ferie_res_media:.1f}" if not np.isnan(ferie_res_media) else "n/d",
        delta=(f"{ferie_res_media:+.1f} vs 0" if not np.isnan(ferie_res_media) else ""),
        delta_color="inverse",
    )
    rC[2].metric(
        "Saldo ferie 2025 (mat - fruite) (gg)",
        f"{saldo_ferie:.0f}" if not np.isnan(saldo_ferie) else "n/d",
        delta=(f"{saldo_ferie:+.0f} vs 0" if not np.isnan(saldo_ferie) else ""),
        delta_color="inverse",
    )
    rC[3].metric(
        "Ferie da far fruire (gg)",
        f"{ferie_da_fruire:.0f}" if not np.isnan(ferie_da_fruire) else "n/d",
        delta=(f"{ferie_da_fruire:+.0f} vs 0" if not np.isnan(ferie_da_fruire) else ""),
        delta_color="inverse",
    )

    st.divider()

    # Tabella persone
    st.markdown("### Persone (indicatori individuali)")
    df_people = build_people_table(df_rep, ore_annue_fte=ore_annue_fte)

    if df_people.empty:
        st.info("Non riesco a costruire la tabella persone (manca MATRICOLA o anagrafica).")
        st.stop()

    # riga totale (sempre in fondo)
    def _make_people_total_row(df_in: pd.DataFrame) -> pd.DataFrame:
        if df_in.empty:
            return pd.DataFrame(columns=df_in.columns)

        out = {c: "" for c in df_in.columns}

        # label
        if "PERSONA" in df_in.columns:
            out["PERSONA"] = "TOTALE"

        # somme per colonne numeriche
        num_cols = df_in.select_dtypes(include=[np.number]).columns.tolist()
        sums = df_in[num_cols].sum(numeric_only=True)
        for c in num_cols:
            out[c] = float(sums.get(c, 0.0))

        fte = float(out.get("FTE", 0.0)) if "FTE" in df_in.columns else 0.0
        ore_teo = float(out.get("ORE_TEORICHE", fte * ore_annue_fte)) if "ORE_TEORICHE" in df_in.columns else fte * ore_annue_fte
        ass_ore = float(out.get("ASSENZE_ORE", 0.0)) if "ASSENZE_ORE" in df_in.columns else 0.0

        if "ORE_TEORICHE" in df_in.columns:
            out["ORE_TEORICHE"] = ore_teo

        # indicatori ricalcolati (non somma di percentuali/ratio)
        if "ASSENTEISMO_%" in df_in.columns:
            out["ASSENTEISMO_%"] = (ass_ore / ore_teo * 100) if ore_teo > 0 else np.nan

        if "FTE_ASSENTI" in df_in.columns:
            out["FTE_ASSENTI"] = (ass_ore / ore_annue_fte) if ore_annue_fte > 0 else np.nan

        if "FTE_DISPONIBILI" in df_in.columns and "FTE_ASSENTI" in df_in.columns:
            out["FTE_DISPONIBILI"] = fte - float(out["FTE_ASSENTI"]) if not pd.isna(out["FTE_ASSENTI"]) else np.nan

        if "STRAORD_ORE_X_FTE" in df_in.columns and "STRAORD_TOT_ORE" in df_in.columns:
            st_tot = float(out.get("STRAORD_TOT_ORE", 0.0))
            out["STRAORD_ORE_X_FTE"] = (st_tot / fte) if fte > 0 else np.nan

        if "FERIE_RES_GIORNI_X_FTE" in df_in.columns and "FERIE_RES_GIORNI" in df_in.columns:
            ferie_res = float(out.get("FERIE_RES_GIORNI", 0.0))
            out["FERIE_RES_GIORNI_X_FTE"] = (ferie_res / fte) if fte > 0 else np.nan

        return pd.DataFrame([out], columns=df_in.columns)

    df_tot_row = _make_people_total_row(df_people)
    df_people_all = pd.concat([df_people, df_tot_row], ignore_index=True)


    # colonne leggibili per la tabella persone
    people_col_labels = {
        "MATRICOLA": "Matricola",
        "PERSONA": "Persona",
        "COGNOME": "Cognome",
        "NOME": "Nome",
        "PROFILO": "Profilo",
        "QUALIFICA_RAW": "Qualifica (raw)",
        "FTE": "FTE",
        "ORE_TEORICHE": "Ore teoriche (h)",
        "ASSENZE_ORE": "Assenze (h)",
        "ASSENTEISMO_%": "Assenteismo (%)",
        "FTE_ASSENTI": "FTE assenti",
        "FTE_DISPONIBILI": "FTE disponibili",
        "STRAORD_TOT_ORE": "Straordinario+Festivi (h)",
        "STRAORD_ORE_X_FTE": "Straordinario+Festivi (h/FTE)",
        "ST_REC": "Recupero (h)",
        "ST_PD": "Straordinario PD (h)",
        "ST_PAG": "Straordinario pagato (h)",
        "FEST_PAG": "Festivo pagato (h)",
        "FEST_REC": "Festivo a recupero (h)",
        "FERIE_RES_GIORNI": "Ferie residue (gg)",
        "FERIE_RES_GIORNI_X_FTE": "Ferie residue (gg/FTE)",
        "FERIE_MAT_GIORNI": "Ferie maturate 2025 (gg)",
        "FERIE_FRUITE_GIORNI": "Ferie fruite 2025 (gg)",
        "mal/104/ecc (ore)": "Assenze mal/104/ecc (h)",
        "asp/grav/puer/dist (ore)": "Assenze asp/grav/puer/dist (h)",
    }
    matr_col_people = find_col(df_people, ["MATRICOLA"], contains=True)
    if matr_col_people and matr_col_people != "MATRICOLA":
        people_col_labels[matr_col_people] = "Matricola"

    # controlli
    cX, cY = st.columns(2)
    sort_options = {
        "Assenteismo (%)": "ASSENTEISMO_%",
        "Assenze (h)": "ASSENZE_ORE",
        "Straordinario+Festivi (h)": "STRAORD_TOT_ORE",
        "Ferie residue (gg)": "FERIE_RES_GIORNI",
        "FTE disponibili": "FTE_DISPONIBILI",
        "FTE": "FTE",
    }
    sort_label = cX.selectbox("Ordina per", list(sort_options.keys()), index=0)
    sort_by = sort_options[sort_label]
    top_people = cY.slider("Mostra top N persone", 10, 300, 50, step=10)

    df_show = df_people.sort_values(sort_by, ascending=False).head(top_people)
    df_show_disp = pd.concat([df_show, df_tot_row], ignore_index=True)
    df_show_disp = df_show_disp.rename(columns=people_col_labels)

    st.dataframe(df_show_disp, use_container_width=True, height=420)

    # KPI con delta entro reparto: confronto reparto vs resto reparto (escludendo Top N persone mostrate)
    st.markdown(f"### KPI – Delta entro reparto (reparto vs resto, esclusi Top {top_people})")

    # Costruisco il "resto reparto" escludendo le persone presenti in df_show
    c_matr_rep = find_col(df_rep, ["MATRICOLA"], contains=True)
    df_rest = pd.DataFrame(columns=df_rep.columns)

    if c_matr_rep and c_matr_rep in df_rep.columns and "MATRICOLA" in df_show.columns:
        top_ids = df_show["MATRICOLA"].dropna().unique().tolist()
        if top_ids:
            df_rest = df_rep.loc[~df_rep[c_matr_rep].isin(top_ids)].copy()
    elif "PERSONA" in df_show.columns and "PERSONA" in df_rep.columns:
        top_ids = df_show["PERSONA"].dropna().unique().tolist()
        if top_ids:
            df_rest = df_rep.loc[~df_rep["PERSONA"].isin(top_ids)].copy()

    k_rest = compute_kpi(df_rest, ore_annue_fte=ore_annue_fte) if not df_rest.empty else None

    def _safe(v):
        try:
            v = float(v)
            return v if not np.isnan(v) else np.nan
        except Exception:
            return np.nan

    def _h_per_fte(k):
        f = _safe(k.get("fte_tot"))
        h = _safe(k.get("abs_tot_ore"))
        return (h / f) if f and f > 0 else np.nan

    def _ferie_res_per_fte(k):
        f = _safe(k.get("fte_tot"))
        g = _safe(k.get("ferie_res_giorni"))
        return (g / f) if f and f > 0 else np.nan

    # valori reparto
    rep_ass_pct = _safe(k_rep.get("ass_pct"))
    rep_hpf = _h_per_fte(k_rep)
    rep_fte_disp = _safe(k_rep.get("fte_disp"))
    rep_st_x = _safe(k_rep.get("st_x_fte"))
    rep_res_media = _safe(k_rep.get("res_giorni_media"))
    rep_res_pf = _ferie_res_per_fte(k_rep)

    # valori resto reparto (se disponibile)
    rest_ass_pct = _safe(k_rest.get("ass_pct")) if k_rest else np.nan
    rest_hpf = _h_per_fte(k_rest) if k_rest else np.nan
    rest_fte_disp = _safe(k_rest.get("fte_disp")) if k_rest else np.nan
    rest_st_x = _safe(k_rest.get("st_x_fte")) if k_rest else np.nan
    rest_res_media = _safe(k_rest.get("res_giorni_media")) if k_rest else np.nan
    rest_res_pf = _ferie_res_per_fte(k_rest) if k_rest else np.nan

    with st.container(border=True):
        r1 = st.columns(3)
        r1[0].metric(
            "Assenteismo (%) – reparto",
            f"{rep_ass_pct:.2f}%" if not np.isnan(rep_ass_pct) else "n/d",
            # più alto = peggio -> delta negativo se il reparto è peggiore del resto
            delta=(rest_ass_pct - rep_ass_pct) if (not np.isnan(rest_ass_pct) and not np.isnan(rep_ass_pct)) else None,
        )
        r1[1].metric(
            "Assenze (h/FTE) – reparto",
            f"{rep_hpf:.1f}" if not np.isnan(rep_hpf) else "n/d",
            # più alto = peggio
            delta=(rest_hpf - rep_hpf) if (not np.isnan(rest_hpf) and not np.isnan(rep_hpf)) else None,
        )
        r1[2].metric(
            "FTE disponibili – reparto",
            f"{rep_fte_disp:.2f}" if not np.isnan(rep_fte_disp) else "n/d",
            # più alto = meglio
            delta=(rep_fte_disp - rest_fte_disp) if (not np.isnan(rest_fte_disp) and not np.isnan(rep_fte_disp)) else None,
        )

        r2 = st.columns(3)
        r2[0].metric(
            "Straordinario+Festivi (h/FTE) – reparto",
            f"{rep_st_x:.2f}" if not np.isnan(rep_st_x) else "n/d",
            # più alto = peggio
            delta=(rest_st_x - rep_st_x) if (not np.isnan(rest_st_x) and not np.isnan(rep_st_x)) else None,
        )
        r2[1].metric(
            "Residuo ferie medio (gg/op) – reparto",
            f"{rep_res_media:.2f}" if not np.isnan(rep_res_media) else "n/d",
            # più alto = peggio (più accumulo)
            delta=(rest_res_media - rep_res_media) if (not np.isnan(rest_res_media) and not np.isnan(rep_res_media)) else None,
        )
        r2[2].metric(
            "Residuo ferie (gg/FTE) – reparto",
            f"{rep_res_pf:.2f}" if not np.isnan(rep_res_pf) else "n/d",
            # più alto = peggio
            delta=(rest_res_pf - rep_res_pf) if (not np.isnan(rest_res_pf) and not np.isnan(rep_res_pf)) else None,
        )

    st.caption("Δ rispetto al **resto del reparto** (stesso reparto, escludendo le Top N persone mostrate): **positivo = meglio**, **negativo = peggiora** (freccia rossa ↓).")


    # download (con nomi colonne leggibili)
    df_people_all_disp = df_people_all.rename(columns=people_col_labels)

    st.download_button(
        "Scarica persone (CSV)",
        data=df_people_all_disp.to_csv(index=False).encode("utf-8"),
        file_name=f"persone_{dim_label.lower()}_{len(chosen_dims)}_selezionati.csv",
        mime="text/csv",
    )
    bufp = io.BytesIO()
    with pd.ExcelWriter(bufp, engine="openpyxl") as writer:
        df_people_all_disp.to_excel(writer, index=False, sheet_name="PERSONE")
    st.download_button(
        "Scarica persone (Excel)",
        data=bufp.getvalue(),
        file_name=f"persone_{dim_label.lower()}_{len(chosen_dims)}_selezionati.xlsx",
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
