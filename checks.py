"""Pure logic for the Image checker, Banned brands and Syntax & duplicates
views — ported from the web validator, no UI.
"""
from __future__ import annotations
import re
from collections import Counter
from pathlib import Path

import banned_brands

IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif",
              ".tif", ".tiff", ".avif", ".jfif")
_EMPTY = ("nan", "", "none")


# --------------------------------------------------------------------------
# Image checker
# --------------------------------------------------------------------------

def normalise_image_name(filename: str) -> str:
    """Filename -> comparable product name (drop extension, '_' means '/')."""
    fname = str(filename).replace('"', "").strip()
    fname = fname.replace("\\", "/").split("/")[-1]
    stem = fname.rsplit(".", 1)[0] if "." in fname else fname
    return stem.replace("_", "/").strip().lower()


def excel_name_rows(user_df, name_col):
    """[(row_idx, original_name, normalised_name)] for non-empty names."""
    out = []
    for idx, val in user_df[name_col].dropna().astype(str).items():
        clean = val.strip()
        if clean.lower() in _EMPTY:
            continue
        out.append((idx, clean, clean.lower()))
    return out


def check_images(user_df, name_col, image_names):
    """Compare the Excel names against a list of image filenames.

    Returns dict with counts and the missing / extra / duplicate breakdowns.
    """
    rows = excel_name_rows(user_df, name_col)
    excel_unique = {n for _i, _o, n in rows}
    excel_counts = Counter(n for _i, _o, n in rows)
    excel_dups = {n: c for n, c in excel_counts.items() if c > 1}

    norm = [normalise_image_name(f) for f in image_names]
    norm = [n for n in norm if n]
    img_counts = Counter(norm)
    img_dups = {n: c for n, c in img_counts.items() if c > 1}
    img_set = set(norm)

    matched = excel_unique & img_set
    missing = [(i, o) for i, o, n in rows if n not in img_set]
    extra = sorted(img_set - excel_unique)

    return {
        "excel_total": len(rows),
        "excel_unique": len(excel_unique),
        "images_total": len(norm),
        "matched": len(matched),
        "missing": missing,            # [(row_idx, original name)]
        "extra": extra,                # normalised image names
        "image_duplicates": img_dups,  # name -> times seen
        "excel_duplicates": excel_dups,
    }


def collect_images(folder, recursive=True):
    """All image files under a folder."""
    return scan_folder(folder, recursive)["images"]


def scan_folder(folder, recursive=True):
    """Scan a folder and report what was found, so an empty result can explain
    itself: {'images', 'skipped', 'ext_counts', 'error'}."""
    p = Path(folder)
    images, skipped = [], []
    ext_counts = Counter()
    error = None
    try:
        it = p.rglob("*") if recursive else p.glob("*")
        for f in it:
            try:
                if not f.is_file():
                    continue
            except OSError:
                continue
            ext = f.suffix.lower()
            ext_counts[ext or "(no extension)"] += 1
            if ext in IMAGE_EXTS:
                images.append(str(f))
            else:
                skipped.append(str(f))
    except Exception as e:  # unreadable path, permissions, …
        error = str(e)
    return {"images": sorted(images), "skipped": skipped,
            "ext_counts": dict(ext_counts), "error": error}


# --------------------------------------------------------------------------
# Banned brands
# --------------------------------------------------------------------------

def find_brand_column(cols):
    c = next((x for x in cols if "brand" in x.lower() and "id" not in x.lower()), None)
    return c or next((x for x in cols if "brand" in x.lower()), None)


def file_brands(user_df, brand_col):
    """Distinct brands in the file (handles pipe-separated values)."""
    out = set()
    for val in user_df[brand_col].dropna().astype(str):
        for b in str(val).split("|"):
            b = b.strip()
            if b and b.lower() not in _EMPTY:
                out.add(b)
    return out


def check_banned_brands(user_df, brand_col):
    """Which brands present in the file are blocked on which sites.

    Returns dict with flagged_by_site, the brand/site axes for a matrix, and
    the list of sites with no issues.
    """
    brands = file_brands(user_df, brand_col)
    flagged_by_site = {}
    for site, cfg in banned_brands.SITE_BANNED.items():
        site_lower = {s.lower() for s in cfg["brands"]}
        if cfg["type"] == "banned":
            flagged_by_site[site] = {b for b in brands if b.lower() in site_lower}
        else:  # allowlist — flag brands NOT allowed there
            flagged_by_site[site] = {b for b in brands if b.lower() not in site_lower}

    flagged_brands = sorted({b for s in flagged_by_site.values() for b in s}, key=str.lower)
    sites_with_flags = [s for s in banned_brands.SITE_BANNED if flagged_by_site[s]]
    clean_sites = [s for s in banned_brands.SITE_BANNED if not flagged_by_site[s]]
    return {
        "brands": sorted(brands, key=str.lower),
        "flagged_by_site": flagged_by_site,
        "flagged_brands": flagged_brands,
        "sites_with_flags": sites_with_flags,
        "clean_sites": clean_sites,
        "total_blocks": sum(len(v) for v in flagged_by_site.values()),
    }


# --------------------------------------------------------------------------
# Syntax & duplicates
# --------------------------------------------------------------------------

def get_skeleton(text):
    """Uppercase->A, lowercase->a, digit->0, everything else kept."""
    if not isinstance(text, str):
        return ""
    out = []
    for ch in text:
        if ch.isupper(): out.append("A")
        elif ch.islower(): out.append("a")
        elif ch.isdigit(): out.append("0")
        else: out.append(ch)
    return "".join(out)


def check_syntax_duplicates(user_df, name_col, name_master_list):
    """Flag names already in the master, repeated within the file, and
    names whose character pattern never appears in the master."""
    valid_names = {n.strip() for n in name_master_list}
    valid_skeletons = {get_skeleton(n) for n in name_master_list}

    series = user_df[name_col].dropna().astype(str).str.strip()
    series = series[~series.str.lower().isin(list(_EMPTY))]
    in_file_counts = Counter(series.tolist())

    report = []
    dup_rows = set()
    for idx, name in user_df[name_col].dropna().astype(str).items():
        clean = name.strip()
        if clean.lower() in _EMPTY:
            continue
        if in_file_counts.get(clean, 0) > 1:
            report.append({"row": idx, "name": clean, "issue": "IN-FILE DUPLICATE",
                           "detail": f"Appears {in_file_counts[clean]}x in this file"})
            dup_rows.add(idx)
        if clean in valid_names:
            report.append({"row": idx, "name": clean, "issue": "DUPLICATE",
                           "detail": "Name already exists in the master file"})
            dup_rows.add(idx)
            continue
        skel = get_skeleton(clean)
        if skel not in valid_skeletons:
            report.append({"row": idx, "name": clean, "issue": "SUSPICIOUS SYNTAX",
                           "detail": f"New pattern: {skel}"})
    return {"report": report, "duplicate_rows": sorted(dup_rows)}
