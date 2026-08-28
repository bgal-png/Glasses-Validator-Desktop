"""Fill rules ported from the Excel VBA macros.

Each rule knows how to derive one target column from other columns. The app
uses them twice:
  * validation — flag cells that are empty (but derivable), differ from the
    rule, or that the rule cannot determine (the macro's "yellow" case);
  * filling    — the "Fill" buttons, applying the rule like the macro does.

Columns are matched by header keyword rather than by spreadsheet letter, so
the rules keep working if columns move.

Source macros: b_HSCodes, c_ItemDescription, d_FillProducingCompany,
e_GlassesModel, f_GlassesColorCode, g_GlassesCollection, h_GlassesUsable,
i_UVFilter, j_FaceShape, k_NoOrder, a_PrivateNames.
"""
from __future__ import annotations
import re

_EMPTY = ("nan", "", "none")


def _s(row, col):
    """Trimmed string value of a column (empty string when missing/blank)."""
    if col is None:
        return ""
    v = row.get(col)
    if v is None:
        return ""
    s = str(v)
    return "" if s.strip().lower() in _EMPTY else s.strip()


def find_col(cols, *keywords):
    """First column whose header contains all the given keywords (case-insensitive)."""
    for c in cols:
        cl = c.lower()
        if all(k.lower() in cl for k in keywords):
            return c
    return None


# --------------------------------------------------------------------------
# Lookup tables (verbatim from the macros)
# --------------------------------------------------------------------------

PRODUCING_COMPANY = {
    "Kering": ["Alexander McQueen", "Balenciaga", "Chloe", "Gucci", "Maui Jim",
               "Montblanc", "Puma", "Saint Laurent"],
    "Marcolin": ["Adidas", "Guess", "Max Mara", "MAX&Co.", "Tom Ford"],
    "Ostalo": ["Arena", "Cebe", "Hawkers", "HEAD", "Lavida", "POC", "Oxydo", "Alpina"],
    "Inspecs": ["Caterpillar", "O'Neill", "Radley", "Superdry"],
    "Marchon": ["Calvin Klein", "Lacoste", "LIU JO", "Nike"],
    "Alensa": ["Alensa"],
    "Adrial": ["Crullé", "Kimikado", "Marisio", "Válle", "LeWish", "Beron"],
    "Luxottica": ["Arnette", "Burberry", "Dolce & Gabbana", "Emporio Armani",
                  "Giorgio Armani", "Armani Exchange", "Michael Kors", "Oakley",
                  "Persol", "Polo Ralph Lauren", "Prada", "Ralph by Ralph Lauren",
                  "Ray-Ban", "Swarovski", "Versace", "Vogue", "Jimmy Choo",
                  "Miu Miu", "Tiffany", "Ralph Lauren"],
    "Safilo": ["Boss by Hugo Boss", "Carolina Herrera", "Carrera", "Chiara Ferragni",
               "David Beckham", "Dsquared2", "Fossil", "Havaianas", "Hugo by Hugo Boss",
               "Kate Spade", "Levi's", "Love Moschino", "Marc Jacobs", "Missoni",
               "Moschino", "Pierre Cardin", "Polaroid", "Tommy Hilfiger", "Under Armour"],
    "GO Eyewear": ["Ana Hickmann"],
    "Strabilia": ["Silhouette"],
    "MCM OPTIK SRL": ["Morel"],
    "Bollé Brands": ["Bollé", "SPY+", "Serengeti"],
    "Noavidet": ["Videt Color Collection", "Videt Style Collection"],
}
# brand (lowercased) -> company
BRAND_TO_COMPANY = {b.lower(): comp for comp, brands in PRODUCING_COMPANY.items() for b in brands}

COLLECTION_KERING = [
    "Alexander McQueen", "Balenciaga", "Gucci", "Saint Laurent", "Chloe", "Dior",
    "Fendi", "Dolce & Gabbana", "Celine", "Miu Miu", "Tom Ford", "Prada",
    "Giorgio Armani", "Beron", "LeWish",
]
COLLECTION_KERING_L = {b.lower() for b in COLLECTION_KERING}
COLLECTION_KERING_VALUE = "Prémiové brýle - Kering"

FASHION_BRANDS = [
    "Botaniq", "Brioni", "Calvin Klein", "Carrera", "Coco song", "Crullé", "Dsquared2",
    "Fossil", "Guess", "Hawkers", "Hugo Boss / BOSS; Boss by Hugo Boss", "Julbo",
    "Kate Spade", "Kimikado", "Lacoste", "Levis", "Marc Jacobs", "Marisio", "Max Mara",
    "Max&Co.", "Michael Kors", "Persol", "Polaroid", "Police", "Puma", "Radley",
    "Ray-Ban", "Seventh Street", "Superdry", "Swarovski", "Swidoo", "Tommy Hilfiger",
    "Vogue", "Under Armour", "Armani Exchange", "Boss by Hugo Boss",
    "Hugo by Hugo Boss", "MAX&Co.",
]
LUXURY_BRANDS = [
    "Alexander McQueen", "Balenciaga", "Bottega Venetta", "Burberry", "Celine",
    "Chiara Ferreagni", "Chloe", "Christian Dior", "Dolce & Gabbana", "Emporio Armani",
    "Fendi", "Givenchy", "Gucci", "Impressio", "Jimmy Choo", "Liu Jo", "Missoni",
    "Moschino / Love Moschino", "Myth", "Pierre Cardin",
    "Polo Ralph Lauren / Ralph by Ralph Lauren", "Prada", "Saint Laurent",
    "Stella McCarteny", "Tiffany", "Tom Ford", "Versace", "Miu Miu", "Beron", "LeWish",
    "Giorgio Armani", "Carolina Herrera", "David Beckham", "Ralph Lauren", "Moschino",
    "Love Moschino",
]
FASHION_L = {b.lower() for b in FASHION_BRANDS}
LUXURY_L = {b.lower() for b in LUXURY_BRANDS}

FACE_SHAPE = {
    "EXTRAVAGANT": "Round face|Oval face|Heart-shaped face|Square face",
    "ROUND": "Oval face|Square face",
    "PILOT": "Round face|Oval face|Heart-shaped face|Square face",
    "CAT EYE": "Oval face|Heart-shaped face",
    "SQUARE": "Round face|Oval face|Heart-shaped face",
    "OVERSIZE": "Round face|Oval face|Heart-shaped face|Square face",
    "RECTANGULAR": "Round face|Oval face|Heart-shaped face",
    "PANTHOS / TEA CUP": "Round face|Oval face|Heart-shaped face|Square face",
    "OVAL / ELIPSE": "Oval face|Heart-shaped face|Square face",
    "SINGLE LENS": "Round face|Oval face|Heart-shaped face|Square face",
    "BUTTERFLY": "Oval face|Square face",
    "HEXAGONAL": "Round face|Oval face",
    "BROWLINE": "Round face|Oval face|Heart-shaped face|Square face",
}


# --------------------------------------------------------------------------
# Column resolution
# --------------------------------------------------------------------------

def resolve_columns(cols):
    """Map logical names to actual headers in the user's file."""
    return {
        "name":        find_col(cols, "glasses name"),
        "meta":        find_col(cols, "meta description"),
        "private":     find_col(cols, "name private") or find_col(cols, "name_private"),
        "type":        find_col(cols, "glasses type"),
        "shape":       find_col(cols, "glasses shape"),
        "frame_type":  find_col(cols, "frame type"),
        "material":    find_col(cols, "main material"),
        "lens_effect": find_col(cols, "lens effect"),
        "usable":      find_col(cols, "glasses usable"),
        "collection":  find_col(cols, "collection"),
        "uv":          find_col(cols, "uv filter"),
        "contain":     find_col(cols, "glasses contain"),
        "model":       find_col(cols, "glasses model"),
        "color_code":  find_col(cols, "color code"),
        "brand":       find_col(cols, "brand"),
        "hs_code":     find_col(cols, "hs code"),
        "item_desc":   find_col(cols, "item description"),
        "producing":   find_col(cols, "producing company"),
        "face_shape":  find_col(cols, "face shape"),
        "no_orders":   find_col(cols, "no-orders"),
    }


# --------------------------------------------------------------------------
# Rule implementations — each returns the derived value, or None when the
# macro would have highlighted the cell yellow (cannot determine).
# --------------------------------------------------------------------------

def r_hs_code(row, C):
    g = _s(row, C["type"]).upper()
    t = _s(row, C["material"]).upper()
    if "SUNGLASSES" in g:
        return "90041091"
    if "PLASTIC" in t and "FRAMES" in g:
        return "90031100"
    if ("METAL" in t or "TITANIUM" in t) and "FRAMES" in g:
        return "90031900"
    return None


def r_item_description(row, C):
    g = _s(row, C["type"]).upper()
    t = _s(row, C["material"]).upper()
    if "SUNGLASSES" in g:
        if t == "PLASTIC":
            return "Sunglasses, plastic frame"
        if t == "METAL":
            return "Sunglasses, metal frame"
        if t in ("METAL|PLASTIC", "PLASTIC|METAL"):
            return "Sunglasses, mixed plastic and metal frame"
        return "Sunglasses"
    if "EYEGLASSES" in g or "FRAMES" in g:
        return "Eyeglasses"
    return None


def r_producing_company(row, C):
    return BRAND_TO_COMPANY.get(_s(row, C["brand"]).lower()) or None


def r_glasses_model(row, C):
    full = _s(row, C["name"])
    brand = _s(row, C["brand"])
    if not full or not brand:
        return None
    if full[:len(brand)].upper() != brand.upper():
        return None
    rest = full[len(brand):].strip()
    if not rest:
        return None
    pos = rest.rfind(" ")
    return rest[:pos].strip() if pos > 0 else rest


def r_color_code(row, C):
    full = _s(row, C["name"])
    pos = full.rfind(" ")
    return full[pos + 1:].strip() if pos > 0 else None


def r_collection(row, C):
    brand = _s(row, C["brand"])
    return COLLECTION_KERING_VALUE if brand.lower() in COLLECTION_KERING_L else None


def r_glasses_usable(row, C):
    brand = _s(row, C["brand"]).lower()
    g = _s(row, C["type"]).upper()
    w = _s(row, C["lens_effect"]).upper()
    brand_cat = "Fashion glasses" if brand in FASHION_L else (
        "Luxury glasses" if brand in LUXURY_L else "")
    use_cat = ""
    if "SUNGLASSES" in g:
        use_cat = "Driving glasses" if "POLARIZED" in w else "Common use"
    if brand_cat and use_cat:
        return f"{brand_cat}|{use_cat}"
    return brand_cat or use_cat or None


def r_uv_filter(row, C):
    return "400" if "SUNGLASSES" in _s(row, C["type"]).upper() else None


def r_face_shape(row, C):
    return FACE_SHAPE.get(_s(row, C["shape"]).upper()) or None


def r_no_orders(row, C):
    q = _s(row, C["frame_type"]).upper()
    ae = _s(row, C["contain"]).upper()
    if "HALF RIM" in q:
        return "CoatingPolarized|Glasses index 1.5"
    if "RIMLESS" in q:
        return "CoatingPolarized|Glasses index 1.5|Glasses index 1.74"
    if "CLIP" in ae:
        return "Glasses index 1.5"
    return None


def make_private_name_rule(eye="", sun="", comp=""):
    """a_PrivateNames — needs the three numbers the VBA form asks for."""
    def rule(row, C):
        v = _s(row, C["meta"]).upper()
        if "COMPUTER GLASSES" in v:
            return f"(Eyeglasses PC {comp})" if comp else None
        if "SUNGLASSES" in v:
            return f"(Sunglasses {sun})" if sun else None
        if "EYEGLASSES" in v:
            return f"(Eyeglasses {eye})" if eye else None
        return None
    return rule


# Rules left unticked in the Fill dialog by default — these derive values from
# the product name/brand, which is usually already correct in the file, so
# re-deriving them is opt-in.
DEFAULT_OFF = {"producing", "model", "color_code"}

# id, label, target logical column, function, source description
RULES = [
    ("hs_code",     "HS Code",           "hs_code",    r_hs_code,           "Glasses type + main material"),
    ("item_desc",   "Item description",  "item_desc",  r_item_description,  "Glasses type + main material"),
    ("producing",   "Producing company", "producing",  r_producing_company, "Brand"),
    ("model",       "Glasses model",     "model",      r_glasses_model,     "Glasses name - Brand - colour code"),
    ("color_code",  "Glasses colour code", "color_code", r_color_code,      "last token of Glasses name"),
    ("collection",  "Glasses collection", "collection", r_collection,       "Brand"),
    ("usable",      "Glasses usable",    "usable",     r_glasses_usable,    "Brand + type + lens effect"),
    ("uv",          "UV filter",         "uv",         r_uv_filter,         "Glasses type"),
    ("face_shape",  "Face shape",        "face_shape", r_face_shape,        "Glasses shape"),
    ("no_orders",   "Lenses no-orders",  "no_orders",  r_no_orders,         "Frame type + Glasses contain"),
]


def evaluate(user_df, rule_ids=None, private_params=None):
    """Run the rules over the sheet.

    Returns a list of dicts:
      {rule, label, row, column, current, derived, status}
    status: 'fill'      — cell empty, rule can supply a value
            'differs'   — cell has a different value than the rule derives
            'undecided' — rule can't determine a value (macro's yellow case)
            'ok'        — cell already matches
    """
    cols = list(user_df.columns)
    C = resolve_columns(cols)

    active = list(RULES)
    if private_params is not None:
        active = active + [("private_name", "Name private", "private",
                            make_private_name_rule(**private_params),
                            "Meta description + entered numbers")]
    if rule_ids is not None:
        active = [r for r in active if r[0] in rule_ids]

    out = []
    for rid, label, target_key, fn, _src in active:
        target = C.get(target_key)
        if not target:
            continue
        for idx, row in user_df.iterrows():
            derived = fn(row, C)
            current = _s(row, target)
            if derived is None:
                status = "undecided" if not current else "ok"
            elif not current:
                status = "fill"
            elif current == derived:
                status = "ok"
            else:
                status = "differs"
            out.append({"rule": rid, "label": label, "row": idx, "column": target,
                        "current": current, "derived": derived, "status": status})
    return out


def apply_rules(user_df, rule_ids=None, private_params=None, overwrite=False):
    """Fill the rule-derived columns. By default only fills empty cells;
    with overwrite=True it also replaces values that differ (like the macros,
    which clear the column first). Returns (new_df, changes)."""
    df = user_df.copy()
    results = evaluate(df, rule_ids=rule_ids, private_params=private_params)
    wanted = {"fill"} | ({"differs"} if overwrite else set())
    changes = []
    for r in results:
        if r["status"] in wanted:
            df.at[r["row"], r["column"]] = r["derived"]
            changes.append(r)
    return df, changes
