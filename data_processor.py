# data_processor.py - Parse and normalise Excel files from ASFIM

import pandas as pd
import numpy as np
import re
from config import COLUMN_ALIASES, CLASSIFICATION_KEYWORDS, ALL_ABB_CODES


def process_sfin_file(uploaded_file) -> tuple[pd.DataFrame, list[str]]:
    """
    Read an ASFIM Excel file and return:
      - df_full : normalised DataFrame with ALL funds
      - warnings: list of warning messages
    Columns guaranteed in output: isin, opcvm, societe_gestion,
    souscripteurs, periodicite, actif_net, classification, perf_1j, ytd
    """
    warnings = []

    # ── Try reading with different sheet strategies ──────────────────────────
    df_raw = None
    try:
        xl = pd.ExcelFile(uploaded_file)
        # Prefer sheet 0 unless another sheet looks better
        for sheet in xl.sheet_names:
            try:
                candidate = pd.read_excel(xl, sheet_name=sheet, header=None)
                # Find the header row (first row containing an ISIN-like value)
                header_row = _detect_header_row(candidate)
                if header_row is not None:
                    df_raw = pd.read_excel(xl, sheet_name=sheet, header=header_row)
                    df_raw.columns = df_raw.columns.astype(str).str.strip()
                    break
            except Exception:
                continue
    except Exception as e:
        warnings.append(f"Erreur lecture fichier: {e}")
        return pd.DataFrame(), warnings

    if df_raw is None or df_raw.empty:
        warnings.append("Impossible de lire les données du fichier.")
        return pd.DataFrame(), warnings

    # ── Map columns ──────────────────────────────────────────────────────────
    col_map = _map_columns(df_raw.columns.tolist(), warnings)

    # ── Build normalised DataFrame ───────────────────────────────────────────
    df = pd.DataFrame()

    for target, source in col_map.items():
        if source:
            df[target] = df_raw[source]
        else:
            df[target] = np.nan

    # ── Drop rows without ISIN ───────────────────────────────────────────────
    df = df.dropna(subset=["isin"])
    df["isin"] = df["isin"].astype(str).str.strip().str.upper()
    df = df[df["isin"].str.match(r"^MA\d{10}$")]

    if df.empty:
        warnings.append("Aucun code ISIN valide trouvé (format attendu: MA + 10 chiffres).")
        return df, warnings

    # ── Clean numeric columns ─────────────────────────────────────────────────
    df["perf_1j"] = df["perf_1j"].apply(_parse_percent)
    df["ytd"]     = df["ytd"].apply(_parse_percent)
    df["actif_net"] = df["actif_net"].apply(_parse_number)

    # ── Pour les fichiers hebdomadaires : si perf_1j est vide, utiliser perf_1s ─
    if "perf_1s" in df.columns:
        df["perf_1s"] = df["perf_1s"].apply(_parse_percent)
        df["perf_1j"] = df["perf_1j"].combine_first(df["perf_1s"])

    # ── Clean text columns ────────────────────────────────────────────────────
    for col in ["opcvm", "societe_gestion", "souscripteurs", "periodicite", "classification"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # ── Infer classification from text if blank ───────────────────────────────
    df["classification"] = df.apply(_fix_classification, axis=1)

    # ── Add ABB flag ──────────────────────────────────────────────────────────
    df["is_abb"] = df["isin"].isin(ALL_ABB_CODES)

    return df.reset_index(drop=True), warnings


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _detect_header_row(df_raw: pd.DataFrame) -> int | None:
    """Scan first 20 rows to find the one most likely to be the column header."""
    for i in range(min(20, len(df_raw))):
        row_vals = df_raw.iloc[i].astype(str).str.lower().tolist()
        hits = sum(1 for v in row_vals if any(
            kw in v for kw in ["isin", "opcvm", "1j", "ytd", "classification",
                               "actif", "société", "gestion", "souscript"]
        ))
        if hits >= 3:
            return i
    # Default: try row 0
    return 0


def _map_columns(columns: list[str], warnings: list[str]) -> dict[str, str | None]:
    """Return {target_field: source_column_name or None}."""
    col_lower = {c.lower().strip(): c for c in columns}
    mapping = {}

    for target, aliases in COLUMN_ALIASES.items():
        found = None
        for alias in aliases:
            # Exact match first
            if alias in columns:
                found = alias
                break
            # Case-insensitive
            if alias.lower() in col_lower:
                found = col_lower[alias.lower()]
                break
        if found is None:
            # Fuzzy: check if any alias is a substring of a column name
            for alias in aliases:
                for col in columns:
                    if alias.lower() in col.lower():
                        found = col
                        break
                if found:
                    break
        if found is None:
            warnings.append(f"Colonne « {target} » non trouvée — les alias testés: {aliases[:3]}")
        mapping[target] = found

    return mapping


def _parse_percent(val) -> float | None:
    """Convert '1.23%', '-0.45%', 0.0123, etc. to a float percentage value."""
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        # If stored as decimal (e.g. 0.0123 for 1.23%) convert
        if abs(val) < 1.0 and val != 0:
            return round(val * 100, 4)
        return round(float(val), 4)
    s = str(val).strip().replace(",", ".").replace(" ", "").replace("\xa0", "")
    s = re.sub(r"[^\d.\-\+%]", "", s)
    is_pct = "%" in s
    s = s.replace("%", "")
    try:
        num = float(s)
        if not is_pct and abs(num) < 1.0 and num != 0:
            num = round(num * 100, 4)
        return round(num, 4)
    except ValueError:
        return None


def _parse_number(val) -> float | None:
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _fix_classification(row) -> str:
    cls = row["classification"].upper()
    if cls:
        return cls
    # Try to infer from OPCVM name
    name = row["opcvm"].lower()
    for cat, kws in CLASSIFICATION_KEYWORDS.items():
        if any(kw in name for kw in kws):
            return cat
    return "AUTRE"


def get_category_classification_keywords(category: str) -> list[str]:
    return CLASSIFICATION_KEYWORDS.get(category, [])
