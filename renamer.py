"""Image matching / renaming / resizing logic, ported verbatim from the web
validator so both stay in sync. Pure functions - no UI.
"""
from __future__ import annotations
import io
import re
from pathlib import Path
from PIL import Image, ImageChops


PHOTO_SUFFIX_RE = re.compile(r"^P\d{1,3}$", re.IGNORECASE)  # P00, P01, ...
PHOTO_NUM_RE = re.compile(r"^\d{1,2}$")                     # 1-2 digit photo index
INVALID_FS_CHARS = set('<>:"/\\|?*')


def list_blob(s):
    """Reduce a string to uppercase alphanumerics only (drop spaces, /, -, &, etc.)."""
    return re.sub(r"[^A-Z0-9]", "", s.upper())


def parse_list_entry(line):
    """A list line is usable if it's non-empty. Returns {'raw': line} or None.
    No brand parsing needed — matching works on the raw text."""
    line = line.strip()
    if not line:
        return None
    return {"raw": line}


def tokenize_source(filename):
    """Strip extension, normalize separators (_ - to space), split into tokens."""
    stem = Path(filename).stem
    return re.sub(r"[_\-]+", " ", stem).split()


def is_photo_token(tok):
    """True for trailing tokens that are photo indices (P00, 01, ...) rather than codes."""
    return bool(PHOTO_SUFFIX_RE.match(tok) or PHOTO_NUM_RE.match(tok))


def code_signature(text):
    """Concatenate the alphanumeric content of tokens that contain a digit
    (model + color codes), dropping pure brand-name words. e.g.
    'Dolce & Gabbana DG4477 252587' -> 'DG4477252587'."""
    parts = []
    for tok in text.split():
        if re.search(r"\d", tok):
            parts.append(re.sub(r"[^A-Z0-9]", "", tok.upper()))
    return "".join(parts)


def match_filename(filename, entries, barcode_map=None):
    """Find the list entry whose alphanumeric core matches the filename's core.
    If the filename is a bare barcode (all digits, >= 8) and barcode_map is given,
    look it up there first and rename to that product's name. Otherwise tries the
    full filename, then strips trailing photo tokens and retries. A match must
    contain a digit (never matches on letters alone) unless it's an exact match.
    Returns matched entry + any leftover photo tokens."""
    tokens = tokenize_source(filename)
    if not tokens:
        return {"status": "error", "reason": "Empty filename after parsing"}

    # Barcode path: the filename is just a barcode (digits only, >= 8 chars).
    if barcode_map:
        stem_digits = re.sub(r"\D", "", Path(filename).stem)
        full_alnum = list_blob(" ".join(tokens))
        if stem_digits and stem_digits == full_alnum and len(stem_digits) >= 8:
            name = barcode_map.get(stem_digits)
            if name:
                return {"status": "matched", "entry": {"raw": name}, "leftover_tokens": []}
            return {"status": "no_match", "tokens": tokens}

    # Split off trailing photo-index tokens (kept for collision suffixes)
    core = list(tokens)
    leftover = []
    while len(core) > 1 and is_photo_token(core[-1]):
        leftover.insert(0, core.pop())

    # Full filename first (in case a 1-2 digit trailing token is really a sub-color),
    # then the photo-stripped core.
    for core_tokens, lo in [(tokens, []), (core, leftover)]:
        f_blob = list_blob(" ".join(core_tokens))
        if not f_blob:
            continue

        # Digit-containing cores (with a model number) match by substring.
        # Digit-less names (e.g. "Nocturna Frames Anima Black Grey") must match
        # a list entry EXACTLY — this avoids latching onto a bare brand word.
        has_digit = bool(re.search(r"\d", f_blob))

        matches = []
        for e in entries:
            e_blob = list_blob(e["raw"])
            if not e_blob:
                continue
            if has_digit:
                hit = f_blob in e_blob or e_blob in f_blob
            else:
                hit = f_blob == e_blob
            if hit:
                matches.append((len(e_blob), e))

        if matches:
            matches.sort(key=lambda x: x[0], reverse=True)
            best_len = matches[0][0]
            top = [e for ln, e in matches if ln == best_len]
            if len(top) > 1:
                return {"status": "ambiguous", "candidates": [e["raw"] for e in top]}
            return {"status": "matched", "entry": top[0], "leftover_tokens": lo}

    # PASS 2 (fallback): the entry's model+color "code signature" appears anywhere
    # in the filename. Handles filenames with extra junk the list lacks — leading
    # zeros, trailing variant codes, brand words missing, e.g.
    # "0DG4477__252587_7009.jpg" vs "Dolce & Gabbana DG4477 252587".
    f_full = list_blob(" ".join(tokens))
    sig_matches = []
    for e in entries:
        sig = code_signature(e["raw"])
        if sig and len(sig) >= 5 and sig in f_full:
            sig_matches.append((len(sig), e))
    if sig_matches:
        sig_matches.sort(key=lambda x: x[0], reverse=True)
        best_len = sig_matches[0][0]
        top = [e for ln, e in sig_matches if ln == best_len]
        if len(top) > 1:
            return {"status": "ambiguous", "candidates": [e["raw"] for e in top]}
        return {"status": "matched", "entry": top[0], "leftover_tokens": leftover}

    return {"status": "no_match", "tokens": tokens}


def resize_centered(image_bytes, target_w=2400, target_h=1800, margin_ratio=0.05, background="auto"):
    """Trim the background around the glasses, scale to fit target_w x target_h
    (preserving aspect ratio), and paste centered on a target-sized PNG canvas.
    background: 'auto' (transparent if source has alpha, else white), 'white', 'transparent'.
    Returns PNG bytes."""
    from PIL import ImageChops
    img = Image.open(io.BytesIO(image_bytes))

    src_has_alpha = img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info)
    if background == "transparent":
        use_alpha = True
    elif background == "white":
        use_alpha = False
    else:  # auto
        use_alpha = src_has_alpha

    if src_has_alpha:
        img = img.convert("RGBA")
        bbox = img.split()[-1].getbbox()  # trim by alpha
    else:
        img = img.convert("RGB")
        bg_ref = Image.new("RGB", img.size, (255, 255, 255))
        diff = ImageChops.difference(img, bg_ref).convert("L")
        mask = diff.point(lambda p: 255 if p > 15 else 0)  # ignore JPEG near-white noise
        bbox = mask.getbbox()

    if bbox:
        img = img.crop(bbox)

    # Scale the glasses to FILL the target minus a margin, preserving aspect
    # ratio. Uses an explicit resize (not thumbnail) so small images are
    # enlarged too — otherwise the margin has no effect on already-small photos.
    avail_w = max(1, int(target_w * (1 - 2 * margin_ratio)))
    avail_h = max(1, int(target_h * (1 - 2 * margin_ratio)))
    w, h = img.size
    if w and h:
        scale = min(avail_w / w, avail_h / h)  # contain; allows upscaling
        img = img.resize((max(1, round(w * scale)), max(1, round(h * scale))), Image.LANCZOS)

    if use_alpha:
        canvas = Image.new("RGBA", (target_w, target_h), (0, 0, 0, 0))
        src = img.convert("RGBA")
        x = (target_w - src.width) // 2
        y = (target_h - src.height) // 2
        canvas.paste(src, (x, y), src)
    else:
        canvas = Image.new("RGB", (target_w, target_h), (255, 255, 255))
        src = img.convert("RGBA") if img.mode == "RGBA" else img.convert("RGB")
        x = (target_w - src.width) // 2
        y = (target_h - src.height) // 2
        if src.mode == "RGBA":
            canvas.paste(src, (x, y), src)  # composite over white
        else:
            canvas.paste(src, (x, y))

    out = io.BytesIO()
    canvas.save(out, format="PNG")
    return out.getvalue()


def safe_name(s):
    return "".join("_" if c in INVALID_FS_CHARS else c for c in s)


def target_name_for(entry, ext):
    base = entry["raw"].replace("/", "_")
    return safe_name(base) + ext


def extract_photo_suffix(filename):
    stem = Path(filename).stem
    m = re.search(r"[_\-]P(\d{2,3})$", stem, re.IGNORECASE)
    if m:
        return f"P{m.group(1)}"
    return None


def derive_photo_suffix(row):
    """Best-effort photo-number suffix for a matched row.
    Priority:
      1. A numeric token in the leftover (e.g. "_01" after the model/color)
      2. A trailing P-suffix in the original filename (P00, P01, ...)
    Returns a string like 'P01' or None."""
    for tok in (row.get("leftover_tokens") or []):
        m = re.match(r"^(\d{1,3})$", tok)
        if m:
            return f"P{int(m.group(1)):02d}"
    return extract_photo_suffix(row["source"])


def resolve_collisions(plan):
    groups = {}
    for row in plan:
        if row["status"] != "matched":
            continue
        groups.setdefault(row["target"], []).append(row)

    for target, rows in groups.items():
        if len(rows) < 2:
            continue
        existing = [derive_photo_suffix(r) for r in rows]
        if all(existing) and len(set(existing)) == len(existing):
            stem, ext = Path(target).stem, Path(target).suffix
            for r, sfx in zip(rows, existing):
                r["target"] = f"{stem} {sfx}{ext}"
                r["collision"] = f"used original suffix {sfx}"
        else:
            stem, ext = Path(target).stem, Path(target).suffix
            for i, r in enumerate(rows):
                r["target"] = f"{stem} P{i:02d}{ext}"
                r["collision"] = f"auto-suffix P{i:02d}"
    return plan
