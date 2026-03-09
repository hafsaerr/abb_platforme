# storage.py - Persistent data storage across sessions

import os
import pickle
import json
from pathlib import Path
from datetime import datetime, date
import pandas as pd

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

HISTORY_FILE = DATA_DIR / "history.pkl"
META_FILE    = DATA_DIR / "metadata.json"


# ─── SAVE ─────────────────────────────────────────────────────────────────────

def save_upload(report_date: date, periodicity: str, df_full: pd.DataFrame):
    """
    Persist a full processed DataFrame for a given date + periodicity.
    Overwrites if the same date+periodicity already exists.
    """
    history = _load_history()
    key = _make_key(report_date, periodicity)
    history[key] = {
        "date": report_date.isoformat(),
        "periodicity": periodicity,
        "df": df_full,
        "saved_at": datetime.now().isoformat(),
    }
    _save_history(history)
    _update_meta(report_date, periodicity)


# ─── LOAD ─────────────────────────────────────────────────────────────────────

def load_all_keys() -> list[dict]:
    """Return sorted list of {key, date, periodicity} dicts."""
    history = _load_history()
    entries = []
    for key, val in history.items():
        entries.append({
            "key": key,
            "date": val["date"],
            "periodicity": val["periodicity"],
        })
    return sorted(entries, key=lambda x: x["date"], reverse=True)


def load_by_key(key: str) -> dict | None:
    history = _load_history()
    return history.get(key)


def load_latest(periodicity: str | None = None) -> dict | None:
    """Load the most recent entry, optionally filtered by periodicity."""
    history = _load_history()
    if not history:
        return None
    filtered = [
        v for v in history.values()
        if periodicity is None or v["periodicity"].lower() == periodicity.lower()
    ]
    if not filtered:
        return None
    return sorted(filtered, key=lambda x: x["date"], reverse=True)[0]


def get_historical_series(isin: str, periodicity: str) -> pd.DataFrame:
    """
    Returns a DataFrame with columns [date, perf_1j, ytd] for a given ISIN,
    across all stored uploads of the given periodicity.
    """
    history = _load_history()
    rows = []
    for val in history.values():
        if val["periodicity"].lower() != periodicity.lower():
            continue
        df: pd.DataFrame = val["df"]
        row = df[df["isin"] == isin]
        if row.empty:
            continue
        r = row.iloc[0]
        rows.append({
            "date":    val["date"],
            "perf_1j": r.get("perf_1j", None),
            "ytd":     r.get("ytd", None),
            "opcvm":   r.get("opcvm", isin),
        })
    if not rows:
        return pd.DataFrame(columns=["date", "perf_1j", "ytd", "opcvm"])
    out = pd.DataFrame(rows).sort_values("date")
    out["date"] = pd.to_datetime(out["date"])
    return out


def delete_entry(key: str):
    history = _load_history()
    if key in history:
        del history[key]
        _save_history(history)


def count_entries() -> int:
    return len(_load_history())


# ─── INTERNAL HELPERS ─────────────────────────────────────────────────────────

def _make_key(report_date: date, periodicity: str) -> str:
    return f"{report_date.isoformat()}_{periodicity.lower()}"


def _load_history() -> dict:
    if not HISTORY_FILE.exists():
        return {}
    try:
        with open(HISTORY_FILE, "rb") as f:
            return pickle.load(f)
    except Exception:
        return {}


def _save_history(history: dict):
    with open(HISTORY_FILE, "wb") as f:
        pickle.dump(history, f)


def _update_meta(report_date: date, periodicity: str):
    meta = {}
    if META_FILE.exists():
        try:
            with open(META_FILE, "r") as f:
                meta = json.load(f)
        except Exception:
            pass
    meta["last_upload"] = {
        "date": report_date.isoformat(),
        "periodicity": periodicity,
        "ts": datetime.now().isoformat(),
    }
    with open(META_FILE, "w") as f:
        json.dump(meta, f, indent=2)
