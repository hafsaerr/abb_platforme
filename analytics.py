# analytics.py - Market statistics, synthesis and ranking tables

import pandas as pd
import numpy as np
from config import CATEGORY_CODES, CLASSIFICATION_KEYWORDS


def get_market_df(df_full: pd.DataFrame, category: str) -> pd.DataFrame:
    """
    Return all funds from df_full belonging to `category` (OCT / OMLT / DIVERSIFIE).
    Detection is by classification column keyword matching.
    """
    kws = [k.lower() for k in CLASSIFICATION_KEYWORDS.get(category, [])]
    mask = df_full["classification"].str.lower().apply(
        lambda v: any(kw in v for kw in kws)
    )
    return df_full[mask].copy()


def get_abb_df(df_full: pd.DataFrame, category: str, periodicity: str) -> pd.DataFrame:
    """Return ABB funds for the given category + periodicity."""
    codes = CATEGORY_CODES.get(category, {}).get(periodicity.lower(), [])
    return df_full[df_full["isin"].isin(codes)].copy()


def compute_tableau1(df_abb: pd.DataFrame) -> pd.DataFrame:
    """
    Table 1 – Fonds ABB.
    Columns: code_isin, opcvm, societe_gestion, souscripteurs,
             periodicite, actif_net, classification, perf_1j, ytd
    Sorted by perf_1j descending.
    """
    cols = ["isin", "opcvm", "societe_gestion", "souscripteurs",
            "periodicite", "actif_net", "classification", "perf_1j", "ytd"]
    existing = [c for c in cols if c in df_abb.columns]
    t1 = df_abb[existing].copy()
    t1 = t1.sort_values("perf_1j", ascending=False).reset_index(drop=True)
    return t1


def compute_tableau2(df_market: pd.DataFrame, df_abb: pd.DataFrame,
                     category: str) -> pd.DataFrame:
    """
    Table 2 – Synthèse du marché.
    Returns a single-column description DataFrame for display.
    """
    market_1j  = df_market["perf_1j"].dropna()
    market_ytd = df_market["ytd"].dropna()
    abb_ytd    = df_abb["ytd"].dropna()

    n_total = len(df_market)
    mean_1j  = market_1j.mean() if len(market_1j) else np.nan
    median_1j = market_1j.median() if len(market_1j) else np.nan

    if len(market_1j) and df_market["perf_1j"].notna().any():
        idx_best  = df_market["perf_1j"].idxmax()
        idx_worst = df_market["perf_1j"].idxmin()
        plus_perf_name  = df_market.loc[idx_best,  "opcvm"]
        plus_perf_val   = df_market.loc[idx_best,  "perf_1j"]
        moins_perf_name = df_market.loc[idx_worst, "opcvm"]
        moins_perf_val  = df_market.loc[idx_worst, "perf_1j"]
    else:
        plus_perf_name  = moins_perf_name = "N/A"
        plus_perf_val   = moins_perf_val  = np.nan

    ytd_moyen_marche = market_ytd.mean() if len(market_ytd) else np.nan
    ytd_moyen_abb    = abb_ytd.mean()    if len(abb_ytd)    else np.nan
    ecart_moyen      = (ytd_moyen_abb - ytd_moyen_marche) if (
        not np.isnan(ytd_moyen_abb) and not np.isnan(ytd_moyen_marche)
    ) else np.nan

    rows = [
        ("Nombre total",                          f"{n_total}",                          ""),
        ("Moyenne de performance quotidienne",     _fmt_pct(mean_1j),                    ""),
        ("Médiane de performance quotidienne",     _fmt_pct(median_1j),                  ""),
        ("Plus performant",                        _fmt_pct(plus_perf_val),              plus_perf_name),
        ("Moins performant",                       _fmt_pct(moins_perf_val),             moins_perf_name),
        (f"YTD moyen marché {category}",           _fmt_pct(ytd_moyen_marche),           ""),
        ("YTD moyen Position ABB",                 _fmt_pct(ytd_moyen_abb),              ""),
        ("Écart moyen (ABB - marché)",             _fmt_pct(ecart_moyen),                ""),
    ]
    df = pd.DataFrame(rows, columns=["Indicateur", "Valeur", "Détail"])
    # Store raw values for downstream use
    df.attrs["ytd_moyen_abb"]    = ytd_moyen_abb
    df.attrs["ytd_moyen_marche"] = ytd_moyen_marche
    df.attrs["mean_1j"]          = mean_1j
    df.attrs["n_total"]          = n_total
    return df


def compute_tableau3(df_market: pd.DataFrame, df_abb: pd.DataFrame,
                     synth: pd.DataFrame) -> pd.DataFrame:
    """
    Table 3 – Ranking.
    Columns: opcvm, rang_interne, rang_marche, position,
             ecart_plus, ecart_moy, ecart_moins, ytd,
             ytd_marche_moyen, ecart_marche, note
    """
    ytd_moyen_marche = synth.attrs.get("ytd_moyen_marche", np.nan)
    ytd_moyen_abb    = synth.attrs.get("ytd_moyen_abb",    np.nan)

    n_market = len(df_market)
    market_valid = df_market.dropna(subset=["ytd"]).copy()
    market_valid = market_valid.sort_values("ytd", ascending=False).reset_index(drop=True)
    market_valid["_mkt_rank"] = range(1, len(market_valid) + 1)

    abb_valid = df_abb.dropna(subset=["ytd"]).copy()
    abb_valid = abb_valid.sort_values("ytd", ascending=False).reset_index(drop=True)
    abb_valid["_abb_rank"] = range(1, len(abb_valid) + 1)

    best_abb_ytd  = abb_valid["ytd"].max()  if len(abb_valid) else np.nan
    worst_abb_ytd = abb_valid["ytd"].min()  if len(abb_valid) else np.nan
    n_abb         = len(abb_valid)

    # Merge market rank into abb
    abb_merged = abb_valid.merge(
        market_valid[["isin", "_mkt_rank"]],
        on="isin", how="left",
    )

    rows = []
    for _, row in abb_merged.iterrows():
        ytd         = row["ytd"]
        mkt_rank    = int(row["_mkt_rank"]) if pd.notna(row["_mkt_rank"]) else None
        abb_rank    = int(row["_abb_rank"])

        pct_rank = (mkt_rank / n_market) if mkt_rank and n_market else 1.0
        if pct_rank <= 0.25:
            position = "Top 25%"
            note = 10
        elif pct_rank <= 0.50:
            position = "Milieu"
            note = 5
        elif pct_rank <= 0.75:
            position = "Milieu"
            note = 3
        else:
            position = "Bas 25%"
            note = 3

        ecart_plus   = (best_abb_ytd  - ytd) if not np.isnan(best_abb_ytd)  else np.nan
        ecart_moins  = (ytd - worst_abb_ytd) if not np.isnan(worst_abb_ytd) else np.nan
        ecart_moy    = (ytd - ytd_moyen_abb) if not np.isnan(ytd_moyen_abb) else np.nan
        ecart_marche = (ytd - ytd_moyen_marche) if not np.isnan(ytd_moyen_marche) else np.nan

        rows.append({
            "OPCVM":               row["opcvm"],
            "Rang interne":        f"{abb_rank}/{n_abb}",
            "Rang marché":         f"{mkt_rank}/{n_market}" if mkt_rank else f"?/{n_market}",
            "Position":            position,
            "Écart vs Plus perf.": ecart_plus,
            "Écart vs Moyenne":    ecart_moy,
            "Écart vs Moins perf.":ecart_moins,
            "YTD":                 ytd,
            "YTD Marché Moyen":    ytd_moyen_marche,
            "Écart vs Marché":     ecart_marche,
            "Note /10":            note,
            # internal keys for sorting
            "_abb_rank":           abb_rank,
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values("_abb_rank").reset_index(drop=True)
        df = df.drop(columns=["_abb_rank"])
    return df


# ─── FORMATTING HELPERS ────────────────────────────────────────────────────────

def _fmt_pct(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"{v:+.3f}%" if v != 0 else "0.000%"


def fmt_pct(v) -> str:
    return _fmt_pct(v)


def fmt_number(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"{v:,.2f}"
