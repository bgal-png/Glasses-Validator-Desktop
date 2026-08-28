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
import pandas as pd
import requests

RAW_BASE = "https://raw.githubusercontent.com/bgal-png/Glasses-Import-Validator/main"
MASTER_URL = f"{RAW_BASE}/master_clean.xlsx"
NAME_MASTER_URL = f"{RAW_BASE}/name_master_clean.xlsx"

APP_NAME = "GlassesValidator"

# Per-file status of the last load: "up-to-date" | "updated" | "downloaded"
# | "offline-cache" | "error". Read by the UI for a friendly message.
LAST_STATUS: dict[str, str] = {}


def cache_dir() -> pathlib.Path:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~/.cache")
    d = pathlib.Path(base) / APP_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def _etag_path(cache_path: pathlib.Path) -> pathlib.Path:
    return cache_path.with_name(cache_path.name + ".etag")


def _fetch(url: str, cache_name: str, timeout: int = 30, force: bool = False) -> bytes:
    """Return file bytes, re-downloading ONLY when the repo copy changed.

    Uses an HTTP conditional request (If-None-Match with the stored ETag). If
    GitHub replies 304 Not Modified, the local cached file is used with no
    download. A full download happens only on first run or a real change.
    """
    cache_path = cache_dir() / cache_name
    etag_path = _etag_path(cache_path)

    headers = {}
    if not force and cache_path.exists() and etag_path.exists():
        headers["If-None-Match"] = etag_path.read_text().strip()

    try:
        resp = requests.get(url, timeout=timeout, headers=headers)
        if resp.status_code == 304 and cache_path.exists():
            LAST_STATUS[cache_name] = "up-to-date"
            return cache_path.read_bytes()
        resp.raise_for_status()
        cache_path.write_bytes(resp.content)
        etag = resp.headers.get("ETag")
        if etag:
            etag_path.write_text(etag)
        LAST_STATUS[cache_name] = "updated" if (cache_path.exists() and headers) else "downloaded"
        return resp.content
    except Exception:
        if cache_path.exists():
            LAST_STATUS[cache_name] = "offline-cache"
            return cache_path.read_bytes()  # offline fallback
        LAST_STATUS[cache_name] = "error"
        raise


def force_refresh() -> None:
    """Delete cached copies, ETags and parsed caches so the next load re-downloads."""
    names = ("master_clean.xlsx", "name_master_clean.xlsx",
             "master_clean.pkl", "name_master_names.json")
    for n in names:
        for p in (cache_dir() / n, _etag_path(cache_dir() / n)):
            if p.exists():
                p.unlink()


def load_master(force: bool = False) -> pd.DataFrame:
    data = _fetch(MASTER_URL, "master_clean.xlsx", force=force)
    pkl = cache_dir() / "master_clean.pkl"
    status = LAST_STATUS.get("master_clean.xlsx")
    # Reuse the fast parsed cache when the Excel hasn't changed
    if not force and status in ("up-to-date", "offline-cache") and pkl.exists():
        try:
            return pd.read_pickle(pkl)
        except Exception:
            pass
    df = pd.read_excel(io.BytesIO(data), dtype=str, engine="openpyxl")
    try:
        df.to_pickle(pkl)
    except Exception:
        pass
    return df


def load_name_master(force: bool = False) -> list[str] | None:
    """Return the list of validated glasses names (name where name_private
    contains 'glasses'), matching the web app's surgical loader."""
    def colf(c):
        if not isinstance(c, str):
            return False
        cl = c.strip().lower()
        return cl == "name" or "name_private" in cl

    try:
        data = _fetch(NAME_MASTER_URL, "name_master_clean.xlsx", force=force)
    except Exception:
        return None

    import json
    names_cache = cache_dir() / "name_master_names.json"
    status = LAST_STATUS.get("name_master_clean.xlsx")
    if not force and status in ("up-to-date", "offline-cache") and names_cache.exists():
        try:
            return json.loads(names_cache.read_text(encoding="utf-8"))
        except Exception:
            pass

    df = pd.read_excel(io.BytesIO(data), dtype=str, engine="openpyxl", usecols=colf)
    df.columns = df.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    private_col = next((c for c in df.columns if "name_private" in c), None)
    name_col = next((c for c in df.columns if c == "name"), None)
    if not private_col or not name_col:
        return None
    filtered = df[df[private_col].str.contains("glasses", case=False, na=False)]
    names = filtered[name_col].dropna().str.strip().unique().tolist()
    try:
        names_cache.write_text(json.dumps(names, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass
    return names
