# Glasses Import Validator — Desktop

A shareable Windows desktop app that opens a glasses import Excel file, renders it
like a spreadsheet, and **shades every cell the validator would flag** (red =
error, amber = warning). Hover a cell to see why. A side control panel gives you
summary counts, an "only rows with issues" filter, jump-to-next-issue, a reference
data refresh, and an annotated-Excel export.

It shares its validation logic and reference data with the web validator:
the master and name-master files are fetched live from the
[`Glasses-Import-Validator`](https://github.com/bgal-png/Glasses-Import-Validator)
repo, so any update you push there reaches this app automatically.

## Run from source

> **Windows long-path note:** the Microsoft Store build of Python uses a very
> long per-user package path, and PySide6's deeply nested files overflow the
> 260-char limit (`OSError [Errno 2] … enable-long-paths`). Install into a
> virtual environment at a **short path** to avoid it:

```bash
python -m venv C:\gv
C:\gv\Scripts\python.exe -m pip install -r requirements.txt
C:\gv\Scripts\python.exe main.py
```

(Or enable Windows Long Paths via registry `LongPathsEnabled=1` and use a normal
`pip install`.)

## Build the .exe
```bash
pip install -r requirements.txt
pyinstaller GlassesValidator.spec
```
The app appears at `dist/GlassesValidator.exe` — a single file you can share.

## How updates work
- **Reference data & rules** (master, name master) load from the GitHub repo at
  launch and are cached in `%LOCALAPPDATA%\GlassesValidator`. Push to the repo →
  everyone gets it (within the cache window; use **☁️ Refresh data** to force).
- **The app itself** auto-updates from GitHub Releases of the desktop repo
  (`bgal-png/Glasses-Validator-Desktop`). On startup it checks the latest release;
  if newer, it offers to download the `.exe` asset and relaunch.

### Publishing a new version
1. Bump `__version__` in `version.py` (e.g. `1.0.1`).
2. Build the exe (above).
3. Create a GitHub Release on `Glasses-Validator-Desktop` with tag `v1.0.1` and
   attach `dist/GlassesValidator.exe` as an asset.
4. Users get the update prompt on their next launch.

## First-time GitHub setup
This repo (`Glasses-Validator-Desktop`) doesn't exist yet — create it, then:
```bash
git init
git add .
git commit -m "Initial desktop validator"
git branch -M main
git remote add origin https://github.com/bgal-png/Glasses-Validator-Desktop.git
git push -u origin main
```

## Files
- `main.py` — the desktop UI (PySide6)
- `model.py` — table model + cell colouring + "issues only" filter
- `validator_core.py` — shared validation logic (no UI)
- `remote.py` — fetch/cache master & name-master from the web repo
- `updater.py` — self-update via GitHub Releases
- `export.py` — annotated `.xlsx` export
- `version.py` — app version
- `GlassesValidator.spec` — PyInstaller build config
