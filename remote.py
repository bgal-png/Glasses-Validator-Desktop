"""Remote reference-data loader.

Fetches the master and name-master workbooks from the existing
Glasses-Import-Validator GitHub repo (the same files the web app uses), caches
them locally, and falls back to the cache when offline. So any update you push
to that repo reaches the desktop app automatically.
"""
from __future__ import annotations
import io
import os
import pathlib
import time
import pandas as pd
import requests

RAW_BASE = "https://raw.githubusercontent.com/bgal-png/Glasses-Import-Validator/main"
MASTER_URL = f"{RAW_BASE}/master_clean.xlsx"
NAME_MASTER_URL = f"{RAW_BASE}/name_master_clean.xlsx"

APP_NAME = "GlassesValidator"


def cache_dir() -> pathlib.Path:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~/.cache")
    d = pathlib.Path(base) / APP_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def _fetch(url: str, cache_name: str, timeout: int = 30, max_age_h: float = 6.0) -> bytes:
    """Return file bytes. Uses a fresh download when possible, else the cache.
    A cache younger than max_age_h is used without re-downloading."""
    cache_path = cache_dir() / cache_name

    # Use a recent cache directly
    if cache_path.exists():
        age_h = (time.time() - cache_path.stat().st_mtime) / 3600.0
        if age_h < max_age_h:
            return cache_path.read_bytes()

    # Try to refresh from the network
    try:
        resp = requests.get(url, timeout=timeout)
        resp.raise_for_status()
        cache_path.write_bytes(resp.content)
        return resp.content
    except Exception:
        if cache_path.exists():
            return cache_path.read_bytes()  # offline fallback
        raise


def force_refresh() -> None:
    """Delete cached copies so the next load re-downloads."""
    for n in ("master_clean.xlsx", "name_master_clean.xlsx"):
        p = cache_dir() / n
        if p.exists():
            p.unlink()


def load_master(force: bool = False) -> pd.DataFrame:
    if force:
        p = cache_dir() / "master_clean.xlsx"
        if p.exists():
            p.unlink()
    data = _fetch(MASTER_URL, "master_clean.xlsx", max_age_h=0 if force else 6.0)
    return pd.read_excel(io.BytesIO(data), dtype=str, engine="openpyxl")


def load_name_master(force: bool = False) -> list[str] | None:
    """Return the list of validated glasses names (name where name_private
    contains 'glasses'), matching the web app's surgical loader."""
    if force:
        p = cache_dir() / "name_master_clean.xlsx"
        if p.exists():
            p.unlink()

    def colf(c):
        if not isinstance(c, str):
            return False
        cl = c.strip().lower()
        return cl == "name" or "name_private" in cl

    try:
        data = _fetch(NAME_MASTER_URL, "name_master_clean.xlsx", max_age_h=0 if force else 6.0)
    except Exception:
        return None
    df = pd.read_excel(io.BytesIO(data), dtype=str, engine="openpyxl", usecols=colf)
    df.columns = df.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    private_col = next((c for c in df.columns if "name_private" in c), None)
    name_col = next((c for c in df.columns if c == "name"), None)
    if not private_col or not name_col:
        return None
    filtered = df[df[private_col].str.contains("glasses", case=False, na=False)]
    return filtered[name_col].dropna().str.strip().unique().tolist()
