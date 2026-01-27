import re
import numpy as np
import pandas as pd

def normalize_colname(s: str) -> str:
    s = str(s)
    s = s.replace("\u00a0", " ")  # NBSP
    for h in ["–", "—", "-", "−"]:
        s = s.replace(h, "-")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def to_num(x):
    # accetta Series o scalari; restituisce Series di float (o float) con NaN -> 0
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = pd.to_numeric(x, errors="coerce")
    if isinstance(s, pd.Series):
        return s.fillna(0)
    return 0.0 if pd.isna(s) else float(s)

def zeros_like(df: pd.DataFrame):
    return pd.Series(0.0, index=df.index)

def resolve_col(df: pd.DataFrame, candidates, contains_ok=True):
    if isinstance(candidates, str):
        candidates = [candidates]

    cols = list(df.columns)
    norm_map = {normalize_colname(c).lower(): c for c in cols}

    # match esatto (normalizzato)
    for w in candidates:
        k = normalize_colname(w).lower()
        if k in norm_map:
            return norm_map[k]

    # match "contains" (utile per varianti tipo "% PART TIME" o "% PART-TIME ")
    if contains_ok:
        wanted_norm = [normalize_colname(w).lower() for w in candidates]
        for c in cols:
            c_norm = normalize_colname(c).lower()
            if any(w in c_norm for w in wanted_norm):
                return c

    return None

@pd.api.extensions.register_dataframe_accessor("cleancols")
class _CleanColsAccessor:
    def __init__(self, pandas_obj):
        self._obj = pandas_obj

    def apply(self):
        self._obj.columns = [normalize_colname(c) for c in self._obj.columns]
        return self._obj

@st.cache_data
def load_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=0)
    df = df.cleancols.apply()
    return df

with st.expander("🔎 Debug: colonne lette dal file"):
    st.write(list(df_raw.columns))


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
        if "aat" in rep:
            return "AUTISTI"
        return "OSS/OT"

    return str(q1)

def build_analisi_dotazioni(df_raw: pd.DataFrame, only_in_force=True):
    df = df_raw.copy()

    # risoluzione colonne (robusta)
    col_pt = resolve_col(df, ["% PART-TIME", "% PART TIME", "PERC PART-TIME", "PART-TIME"], contains_ok=True)
    col_qual2 = resolve_col(df, ["QUALIFICA.1", "QUALIFICA"], contains_ok=True)
    col_rep = resolve_col(df, ["REPARTO"], contains_ok=True)
    col_cdr = resolve_col(df, ["CDR_DESC", "CDR DESC", "CDR"], contains_ok=True)

    # FTE
    if col_pt is None:
        df["FTE"] = 1.0
    else:
        pt = to_num(df[col_pt])
        # se 0 o mancante -> consideriamo 100
        pt = pt.where(pt > 0, 100)
        df["FTE"] = pt / 100

    # ferie
    col_ferie = resolve_col(df, ["FERIE"], True)
    col_ferie_rx = resolve_col(df, ["FERIE RX"], True)
    col_ferie_god_tot = resolve_col(df, ["FERIE GODUTE TOTALE"], True)
    col_ferie_god_rx = resolve_col(df, ["FERIE GODUTE RX"], True)
    col_ferie_res = resolve_col(df, ["FERIE RES."], True)
    col_ferie_rx_res = resolve_col(df, ["FERIE RX RES."], True)
    col_ferie_ap_res = resolve_col(df, ["FERIE AP RES."], True)

    df["FERIE_MAT_2025"] = to_num(df[col_ferie]) + to_num(df[col_ferie_rx])
    df["FERIE_FRUITE_2025"] = to_num(df[col_ferie_god_tot]) + to_num(df[col_ferie_god_rx])
    df["FERIE_RES_0101"] = to_num(df[col_ferie_res]) + to_num(df[col_ferie_rx_res]) + to_num(df[col_ferie_ap_res])

    # assenze principali (se qualche colonna manca -> 0)
    def s(colname):
        c = resolve_col(df, [colname], True)
        return to_num(df[c]) if c else zeros_like(df)

    df["ASSENZE_MAL104_ECC"] = (
        s("MALATTIA") +
        s("MALATTIA FIGLIO") +
        s("LEGGE 104") +
        s("PERMESSI") +
        s("AGGIOR.") +
        s("INF./MAL.SERV") +
        s("CAR.PUBBLICA") +
        s("INFORTUNIO COVID") +
        s("MALATTIA COVID")
    )

    df["ASP_GRAV_PUER_DIST"] = s("RECUPERO") + s("MISSIONE SOLO SERVIZIO")

    # straordinari / festivi
    df["STRAORD_REC"] = s("ORE DA RECUP. PROG.")
    df["STRAORD_PD"] = s("STR. PD. PROG.")
    df["STRAORD_PAG"] = s("STR. PROG.")
    df["FEST_PAG"] = s("FEST. INFRASETT. A PAGAMENTO")
    df["FEST_REC"] = s("FEST. INFRASETT. A RECUPERO")

    # QUALIFICA semplificata
    df["QUALIFICA_S"] = [
        simplify_qualifica(q1, rep)
        for q1, rep in zip(df[col_qual2] if col_qual2 else [None]*len(df),
                           df[col_rep] if col_rep else [None]*len(df))
    ]

    # ✅ UUOO/SERVIZIO automatico e scalabile: CDR_DESC (fallback REPARTO)
    if col_cdr and col_rep:
        df["SERVIZIO"] = df[col_cdr].astype(str).str.strip()
        df.loc[df["SERVIZIO"].isin(["", "nan", "None"]), "SERVIZIO"] = df[col_rep].astype(str).str.strip()
    elif col_cdr:
        df["SERVIZIO"] = df[col_cdr].astype(str).str.strip()
    else:
        df["SERVIZIO"] = df[col_rep].astype(str).str.strip()

    # scope qualifica (puoi allargare)
    target_qual = {"INFERMERE", "OSS", "AUTISTI", "TSRM", "TSLB", "FISOTERAP", "OSS/OT"}
    df_scope = df[df["SERVIZIO"].notna() & df["QUALIFICA_S"].isin(target_qual)].copy()

    # opzionale: solo in forza a fine periodo
    col_data_al = resolve_col(df_scope, ["DATA AL"], True)
    if only_in_force and col_data_al:
        max_data = df_scope[col_data_al].max()
        df_scope = df_scope[df_scope[col_data_al] == max_data].copy()

    # regola: OSS/OT -> OSS (se vuoi tenerli distinti, rimuovi queste 2 righe)
    df_scope["QUALIFICA_OUT"] = df_scope["QUALIFICA_S"]
    df_scope.loc[df_scope["QUALIFICA_OUT"] == "OSS/OT", "QUALIFICA_OUT"] = "OSS"

    col_matr = resolve_col(df_scope, ["MATRICOLA"], True)

    agg = df_scope.groupby(["SERVIZIO", "QUALIFICA_OUT"]).agg(
        OPERATORI=(col_matr, "nunique") if col_matr else ("QUALIFICA_OUT", "size"),
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

    agg["Media procapite"] = agg["FERIE_FRUITE"] / agg["OPERATORI"]
    agg["Prestazioni aggiuntive (ore)"] = 0.0

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
        "FERIE_RES": "Residue al 01/01/2026",
        "ASSENZE": "Assenze mal/104/ecc (ore)",
        "ASP": "Asp/grav/puer/dist (ore)",
    })

    return out, df_scope

