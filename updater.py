"""Self-update via GitHub Releases.

Checks the latest release of the desktop repo; if its tag is newer than the
running version, downloads the packaged .exe asset and swaps it in on exit.
Only performs the swap when running as a frozen PyInstaller build.
"""
from __future__ import annotations
import os
import sys
import subprocess
import tempfile
import requests

from version import __version__

# GitHub repo that hosts the desktop-app releases (create this repo, publish
# releases with the built .exe attached as an asset).
GH_OWNER = "bgal-png"
GH_REPO = "Glasses-Validator-Desktop"
API_LATEST = f"https://api.github.com/repos/{GH_OWNER}/{GH_REPO}/releases/latest"


def _parse(tag: str):
    tag = tag.lstrip("vV").strip()
    parts = []
    for p in tag.split("."):
        num = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(num) if num else 0)
    return tuple(parts) or (0,)


def is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def check_latest(timeout: int = 15) -> dict | None:
    """Return {'tag','newer','exe_url','notes'} or None on failure/no release."""
    try:
        r = requests.get(API_LATEST, timeout=timeout,
                         headers={"Accept": "application/vnd.github+json"})
        if r.status_code != 200:
            return None
        data = r.json()
        tag = data.get("tag_name") or ""
        exe_url = None
        for asset in data.get("assets", []):
            if asset.get("name", "").lower().endswith(".exe"):
                exe_url = asset.get("browser_download_url")
                break
        return {
            "tag": tag,
            "newer": _parse(tag) > _parse(__version__),
            "exe_url": exe_url,
            "notes": data.get("body") or "",
        }
    except Exception:
        return None


def download_and_apply(exe_url: str, timeout: int = 120) -> bool:
    """Download the new exe and schedule a swap+relaunch on process exit.
    Returns True if the swap was scheduled (caller should then quit the app)."""
    if not is_frozen():
        # In dev there is no exe to replace.
        return False
    current = sys.executable  # the running .exe
    new_path = os.path.join(tempfile.gettempdir(), "GlassesValidator_new.exe")
    try:
        r = requests.get(exe_url, timeout=timeout, stream=True)
        r.raise_for_status()
        with open(new_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1 << 16):
                f.write(chunk)
    except Exception:
        return False

    # A batch script waits for this process to exit, replaces the exe, relaunches.
    bat = os.path.join(tempfile.gettempdir(), "GlassesValidator_update.bat")
    with open(bat, "w", encoding="utf-8") as f:
        f.write(
            "@echo off\r\n"
            "timeout /t 2 /nobreak >nul\r\n"
            f'move /y "{new_path}" "{current}" >nul\r\n'
            f'start "" "{current}"\r\n'
            'del "%~f0"\r\n'
        )
    subprocess.Popen(["cmd", "/c", bat], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return True
