"""Shared validation logic for the Glasses Import Validator (desktop + web).

Pure functions — no Streamlit, no GUI. Given a user DataFrame and a master
DataFrame, produce a per-cell map of issues that a grid can colour, plus
a flat list of issues and summary counts.

This mirrors Tab 1 of the Streamlit validator exactly:
  - empty required fields (always + sunglasses-only)
  - Meta description format (type prefix must be followed by a space)
  - whitespace anomalies on every column (leading/trailing/double/NBSP/around |)
  - content validation on mapped columns (invalid value / case mismatch vs master)
"""
from __future__ import annotations
import re
import pandas as pd

# ---- Column configuration (kept in sync with the web validator) ----

REQUIRED_COLUMN_KEYWORDS = [
    "Glasses name", "Meta description", "XML description", "Combination", "Barcode",
    "Glasses type ID", "Manufacturer ID", "temple length ID", "lens height ID",
    "lens width ID", "bridge ID", "Glasses shape ID", "frame type ID",
    "Frame Colour ID", "Temple Colour ID", "main material ID", "gendre ID",
    "Items type ID", "Items packing ID", "Glasses contain ID", "Glasses model ID",
    "color code ID", "Brand ID", "HS Code", "Item description", "Case length",
    "Case height", "Case width", "Case weight", "Glasses weight", "origin country",
    "Producing company ID",
]

SUNGLASSES_REQUIRED_KEYWORDS = [
    "lens Colour ID", "lens material ID", "lens effect ID", "Sunglasses filter ID",
]

IDEAL_PAIRS = {
    "Glasses type": "Glasses type ID",
    "Manufacturer": "Manufacturer ID",
    "Glasses size: glasses width": "width ID",
    "Glasses size: temple length": "temple length ID",
    "Glasses size: lens height": "lens height ID",
    "Glasses size: lens width": "lens width ID",
    "Glasses size: bridge": "bridge ID",
    "Glasses shape": "Glasses shape ID",
    "Glasses other info": "other info ID",
    "Glasses frame type": "frame type ID",
    "Glasses frame color": "Frame Colour ID",
    "Glasses temple color": "Temple Colour ID",
    "Glasses main material": "main material ID",
    "Glasses lens color": "lens Colour ID",
    "Glasses lens material": "lens material ID",
    "Glasses lens effect": "lens effect ID",
    "Sunglasses filter": "Sunglasses filter ID",
    "Glasses genre": "Glasses gendre ID",
    "Glasses usable": "Glasses usable ID",
    "Glasses collection": "Glasses collection ID",
    "UV filter": "UV filter ID",
    "Items type": "Items type ID",
    "Items packing": "Items packing ID",
    "Glasses contain": "Glasses contain ID",
    "Sport glasses": "Sports Glasses ID",
    "Glasses frame color effect": "frame color effect ID",
    "Glasses other features": "other features ID",
    "SunGlasses RX lenses": "RX lenses ID",
    "Glasses clip-on lens color": "clip-on lens colour ID",
    "Brand": "Brand ID",
    "Producing company": "Producing company ID",
    "Glasses for your face shape": "face shape ID",
    "Glasses lenses no-orders": "no-orders ID",
}

META_TYPE_PREFIXES = ["Sunglasses", "Eyeglasses"]

# Any whitespace char incl. NBSP, zero-width space, BOM, tab, etc.
WS_CHARS = r"[\s ​﻿]"
_EMPTY = ("nan", "", "none")


def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise column headers the same way the web validator does."""
    df = df.copy()
    df.columns = df.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    return df


def filter_glasses(master_df: pd.DataFrame) -> pd.DataFrame:
    """Filter the master to Items type == 'Glasses' (matches the web loader)."""
    target_col = next((c for c in master_df.columns if "Items type" in c), None)
    if target_col:
        return master_df[master_df[target_col] == "Glasses"]
    return master_df


def _visible_ws(raw_val: str) -> str:
    return raw_val.replace(" ", "[NBSP]").replace("\t", "[TAB]")


def fix_spacing_value(raw: str) -> str:
    """Normalise the whitespace problems the validator flags:
    NBSP/zero-width -> normal space, collapse runs of whitespace, tidy spaces
    around the '|' separator, and trim the ends."""
    s = raw
    # Exotic whitespace -> plain space (NBSP, zero-width space, BOM)
    s = s.replace(" ", " ").replace("​", " ").replace("﻿", " ")
    s = s.replace("\t", " ")
    # Collapse any run of whitespace to a single space
    s = re.sub(r"\s{2,}", " ", s)
    # No spaces around the pipe separator
    s = re.sub(r"\s*\|\s*", "|", s)
    return s.strip()


def fix_spacing(user_df: pd.DataFrame):
    """Apply fix_spacing_value to every cell. Returns (new_df, changes) where
    changes is a list of (row_idx, col_name, before, after)."""
    df = user_df.copy()
    changes = []
    for col in df.columns:
        series = df[col]
        for row_idx in range(len(df)):
            v = series.iat[row_idx]
            if v is None:
                continue
            raw = str(v)
            if raw.lower() in _EMPTY:
                continue
            fixed = fix_spacing_value(raw)
            if fixed != raw:
                changes.append((row_idx, col, raw, fixed))
                df.iat[row_idx, df.columns.get_loc(col)] = fixed
    return df, changes


def validate(user_df: pd.DataFrame, master_df: pd.DataFrame) -> dict:
    """Run the full Tab-1 validation.

    Returns a dict with:
      cell_issues : { (row_idx, col_name): [ {type, severity, message} ] }
      issues      : flat list of issue dicts (Row is the Excel row = idx+2)
      mapping     : {active, required_found, sunglasses_found, unmapped, meta_col}
      counts      : {empty, meta_format, whitespace, invalid, case_mismatch, total}
    """
    user_df = clean_headers(user_df)
    master_df = clean_headers(master_df)
    master_df = filter_glasses(master_df)

    user_cols = list(user_df.columns)
    master_cols = list(master_df.columns)

    # ---- Column resolution ----
    active_map = {}
    unmapped = []
    for mk, uk in IDEAL_PAIRS.items():
        rmc = next((c for c in master_cols if mk in c), None)
        ruc = next((c for c in user_cols if uk in c), None)
        if rmc and ruc:
            active_map[rmc] = ruc
        else:
            unmapped.append((mk, uk))

    required_col_map = {}
    for kw in REQUIRED_COLUMN_KEYWORDS:
        m = next((c for c in user_cols if kw.lower() in c.lower()), None)
        if m:
            required_col_map[kw] = m

    sunglasses_col_map = {}
    for kw in SUNGLASSES_REQUIRED_KEYWORDS:
        m = next((c for c in user_cols if kw.lower() in c.lower()), None)
        if m:
            sunglasses_col_map[kw] = m

    meta_col = next((c for c in user_cols if "meta description" in c.lower()), None)

    # ---- Build case-insensitive master vocab per mapped column ----
    valid_values_ci = {}
    for m_col in active_map.keys():
        raw = master_df[m_col].dropna().astype(str)
        exploded = raw.str.split(r",+").explode().str.strip()
        mapping = {}
        for v in exploded:
            if v and v.lower() not in mapping:
                mapping[v.lower()] = v  # first-seen casing wins
        valid_values_ci[m_col] = mapping

    cell_issues: dict = {}
    issues: list = []

    def add(row_idx, col, itype, severity, message, **extra):
        expected = extra.get("Expected")
        if expected is None and extra.get("Allowed"):
            expected = "e.g. " + ", ".join(extra["Allowed"])
        cell_issues.setdefault((row_idx, col), []).append({
            "type": itype,
            "severity": severity,
            "message": message,
            "value": extra.get("Value"),
            "expected": expected,
        })
        rec = {"Row": row_idx + 2, "Column": col, "Error": message, "type": itype}
        rec.update(extra)
        issues.append(rec)

    # ---- Per-row checks ----
    for idx, row in user_df.iterrows():
        # Meta description type + format
        raw_meta = str(row[meta_col]).strip() if meta_col else ""
        meta_val = raw_meta.lower()
        is_sunglasses = meta_val.startswith("sunglasses")
        for t in META_TYPE_PREFIXES:
            tl = t.lower()
            if meta_val.startswith(tl) and len(meta_val) > len(tl) and meta_val[len(tl)] != " ":
                add(idx, meta_col, "meta_format", "error",
                    f"Missing space after '{t}'", Value=f"Missing space after '{t}'", Content=raw_meta)
                break

        # Empty required (always)
        for kw, u_col in required_col_map.items():
            if str(row[u_col]).strip().lower() in _EMPTY:
                add(idx, u_col, "empty_required", "error", "Empty required field")

        # Empty required (sunglasses only)
        if is_sunglasses:
            for kw, u_col in sunglasses_col_map.items():
                if str(row[u_col]).strip().lower() in _EMPTY:
                    add(idx, u_col, "empty_required_sun", "error", "Empty required field (Sunglasses)")

    # ---- Whitespace check (all columns) ----
    for idx, row in user_df.iterrows():
        for u_col in user_cols:
            raw_val = str(row[u_col])
            if raw_val.lower() in _EMPTY:
                continue
            ws = []
            if re.match(WS_CHARS, raw_val): ws.append("Leading space")
            if re.search(WS_CHARS + r"$", raw_val): ws.append("Trailing space")
            if re.search(WS_CHARS + r"{2,}", raw_val): ws.append("Double spaces")
            if re.search(r"\|\s|\s\|", raw_val): ws.append("Space around separator")
            if " " in raw_val: ws.append("Non-breaking space (NBSP)")
            for w in ws:
                add(idx, u_col, "whitespace", "warning", w, Value=w, Content=_visible_ws(raw_val))

    # ---- Content validation (mapped columns) ----
    for idx, row in user_df.iterrows():
        for m_col, u_col in active_map.items():
            raw_val = str(row[u_col])
            if raw_val.lower() in _EMPTY:
                continue
            ci_map = valid_values_ci[m_col]
            for p in (v.strip() for v in raw_val.strip().split("|")):
                if not p:
                    continue
                if p.lower() not in ci_map:
                    add(idx, u_col, "invalid_content", "error",
                        f"Invalid value: {p}", Value=p, Content=raw_val,
                        Allowed=list(ci_map.values())[:3])
                elif p != ci_map[p.lower()]:
                    add(idx, u_col, "case_mismatch", "warning",
                        f"Case mismatch: '{p}' should be '{ci_map[p.lower()]}'",
                        Value=p, Content=raw_val, Expected=ci_map[p.lower()])

    counts = {
        "empty": sum(1 for i in issues if i["type"].startswith("empty")),
        "meta_format": sum(1 for i in issues if i["type"] == "meta_format"),
        "whitespace": sum(1 for i in issues if i["type"] == "whitespace"),
        "invalid": sum(1 for i in issues if i["type"] == "invalid_content"),
        "case_mismatch": sum(1 for i in issues if i["type"] == "case_mismatch"),
        "total": len(issues),
    }
    mapping = {
        "active": active_map,
        "required_found": required_col_map,
        "sunglasses_found": sunglasses_col_map,
        "unmapped": unmapped,
        "meta_col": meta_col,
    }
    return {"cell_issues": cell_issues, "issues": issues, "mapping": mapping, "counts": counts}
