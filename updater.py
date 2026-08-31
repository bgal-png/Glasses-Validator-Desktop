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


def _looks_like_exe(path: str, min_bytes: int = 5_000_000) -> bool:
    """Guard against replacing a working app with an error page or a truncated
    download: must be a real PE binary of a plausible size."""
    try:
        if os.path.getsize(path) < min_bytes:
            return False
        with open(path, "rb") as f:
            return f.read(2) == b"MZ"
    except OSError:
        return False


def download_and_apply(exe_url: str, timeout: int = 600) -> bool:
    """Download the new exe and schedule a swap+relaunch once this process exits.
    Returns True if the swap was scheduled (caller should then quit the app).

    Uses PowerShell rather than a .bat: the previous batch script called
    `timeout`, which fails instantly with "Input redirection is not supported"
    when there is no console (CREATE_NO_WINDOW). The batch file then carried on
    and tried to move the exe while it was still locked, so the move failed and
    the update silently did nothing.
    """
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
                if chunk:
                    f.write(chunk)
    except Exception:
        return False

    if not _looks_like_exe(new_path):
        try:
            os.remove(new_path)
        except OSError:
            pass
        return False

    def q(p):  # PowerShell single-quoted literal
        return p.replace("'", "''")

    backup = current + ".old"
    script = f"""$ErrorActionPreference = 'SilentlyContinue'
# Wait for the app to actually exit rather than guessing with a sleep.
try {{ Wait-Process -Id {os.getpid()} -Timeout 180 }} catch {{ }}
$new = '{q(new_path)}'
$cur = '{q(current)}'
$bak = '{q(backup)}'
Remove-Item -LiteralPath $bak -Force -ErrorAction SilentlyContinue
$ok = $false
for ($i = 0; $i -lt 60; $i++) {{
    try {{
        # Keep the old exe aside so a failed swap can be rolled back.
        Move-Item -LiteralPath $cur -Destination $bak -Force -ErrorAction Stop
        Move-Item -LiteralPath $new -Destination $cur -Force -ErrorAction Stop
        $ok = $true
        break
    }} catch {{
        if ((Test-Path -LiteralPath $bak) -and -not (Test-Path -LiteralPath $cur)) {{
            Move-Item -LiteralPath $bak -Destination $cur -Force -ErrorAction SilentlyContinue
        }}
        Start-Sleep -Milliseconds 500
    }}
}}
if ($ok) {{ Remove-Item -LiteralPath $bak -Force -ErrorAction SilentlyContinue }}
Start-Process -FilePath $cur
Remove-Item -LiteralPath $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
"""
    ps1 = os.path.join(tempfile.gettempdir(), "GlassesValidator_update.ps1")
    with open(ps1, "w", encoding="utf-8") as f:
        f.write(script)

    subprocess.Popen(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-WindowStyle", "Hidden", "-File", ps1],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    return True
