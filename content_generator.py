"""
Content generation engine for Backpacks / Trolley Bags / Handbags / Laptop Bags.

Given a category and a mapping of nomenclature field-keys -> actual column
names in the user's file, generates for each SKU row:
  - DN   (Display Name / productDisplayName)
  - LVN  (List View Name / listViewName)
  - PD   (Product Details, bullet-style multi-line text)
  - Size & Fit (sizeAndFitDescription)

Business rules are transcribed from the nomenclature the team supplied.
Fields left blank in the source data are dropped from the generated text
rather than printed empty.
"""

import difflib

WARRANTY_DISCLAIMER = "Warranty provided by brand owner/manufacturer"

CATEGORIES = ["Backpacks", "Trolley Bag", "Handbags", "Laptop Bag"]


def is_blank(v):
    if v is None:
        return True
    s = str(v).strip()
    return s == "" or s.lower() == "nan" or s.lower() == "none"


def clean(v):
    if is_blank(v):
        return ""
    s = str(v).strip()
    # Format whole-number floats without a trailing .0 (e.g. pandas often
    # reads integer counts as floats: 1.0 -> "1")
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except ValueError:
        pass
    return s


def join_nonblank(parts, sep=" "):
    return sep.join([p for p in parts if not is_blank(p)])


# ---------------------------------------------------------------------------
# Field keys required per category, with candidate column names (in priority
# order) used to auto-guess the mapping against the user's actual headers.
# ---------------------------------------------------------------------------

FIELD_CANDIDATES = {
    "Brand": ["Brand", "brand"],
    "Multipack Set": ["Multipack Set", "MultiPack Set"],
    "Print or Pattern Type": ["Print or Pattern Type", "Pattern"],
    "Surface Styling": ["Surface Styling"],
    "Shoulder Strap Type": ["Shoulder Strap Type", "Handles"],
    "Features": ["Features", "Feature"],
    "Technology": ["Technology"],
    "Main Trend": ["Main Trend", "Trends"],
    "Article Type": ["Article Type", "articleType", "AT"],
    "Base Colour": ["Base Colour", "Prominent Colour"],
    "Colour 1": ["Colour 1", "Second Prominent Colour"],
    "Colour 2": ["Colour 2", "Third Prominent Colour"],
    "Number of Main Compartments": ["Number of Main Compartments"],
    "Compartment Closure": ["Compartment Closure"],
    "Number of External Pockets": ["Number of External Pockets"],
    "Number of Inner Pocket": ["Number of Inner Pocket", "Number of Inner Pockets"],
    "Sling Strap": ["Sling Strap"],
    "Add Ons": ["Add Ons", "Add-Ons", "Addons"],
    "Warranty": ["Warranty"],
    "Length": ["Length", "Package Length"],
    "Height": ["Height", "Package Height"],
    "Depth": ["Depth"],
    "Width": ["Width", "Package Breadth"],
    "Volume in Litres": ["Volume in Litres"],
    "Type": ["Type"],
    "Material": ["Material"],
    "Size": ["Size"],
    "Laptop Size": ["Laptop Size"],
    "Laptop Compartment": ["Laptop Compartment"],
    # Trolley-specific
    "Trolley Size": ["Trolley Size"],
    "Bag Type": ["Bag Type"],
    "Number of Wheels": ["Number of Wheels"],
    "Number of Handles": ["Number of Handles", "Handles"],
    "Lock Type": ["Lock Type"],
    "Shell Type": ["Shell Type"],
    "Maximum Carrying Capacity": ["Maximum Carrying Capacity"],
    "Dead Weight": ["Dead Weight"],
    "Number of Internal Side Pockets": ["Number of Internal Side Pockets"],
}

CATEGORY_FIELDS = {
    "Backpacks": [
        "Brand", "Multipack Set", "Print or Pattern Type", "Surface Styling",
        "Shoulder Strap Type", "Features", "Technology", "Main Trend", "Article Type",
        "Base Colour", "Colour 1", "Colour 2", "Number of Main Compartments",
        "Compartment Closure", "Number of External Pockets", "Warranty",
        "Length", "Height", "Depth", "Volume in Litres",
    ],
    "Trolley Bag": [
        "Brand", "Multipack Set", "Print or Pattern Type", "Features",
        "Number of Wheels", "Trolley Size", "Article Type", "Volume in Litres",
        "Base Colour", "Colour 1", "Colour 2", "Bag Type", "Laptop Compartment",
        "Number of Main Compartments", "Number of Handles", "Lock Type", "Shell Type",
        "Number of External Pockets", "Number of Internal Side Pockets",
        "Maximum Carrying Capacity", "Add Ons", "Warranty",
        "Height", "Width", "Depth", "Dead Weight",
    ],
    "Handbags": [
        "Brand", "Multipack Set", "Print or Pattern Type", "Material", "Size",
        "Surface Styling", "Main Trend", "Type", "Base Colour",
        "Number of Main Compartments", "Compartment Closure",
        "Number of External Pockets", "Number of Inner Pocket", "Sling Strap",
        "Add Ons", "Warranty", "Height", "Width", "Depth", "Volume in Litres",
    ],
    "Laptop Bag": [
        "Brand", "Multipack Set", "Print or Pattern Type", "Material",
        "Main Trend", "Features", "Type", "Laptop Size", "Base Colour",
        "Number of Main Compartments", "Laptop Compartment",
        "Number of External Pockets", "Number of Inner Pocket", "Warranty",
        "Height", "Width", "Depth",
    ],
}


def get_required_fields(category):
    return CATEGORY_FIELDS.get(category, [])


def guess_column(field_key, available_columns):
    """Best-effort auto-match of a nomenclature field key to a real column name."""
    candidates = FIELD_CANDIDATES.get(field_key, [field_key])
    lower_map = {c.lower().strip(): c for c in available_columns}
    for cand in candidates:
        if cand.lower().strip() in lower_map:
            return lower_map[cand.lower().strip()]
    # fuzzy fallback
    close = difflib.get_close_matches(field_key, available_columns, n=1, cutoff=0.75)
    if close:
        return close[0]
    return None


def combine_colour(base, c1, c2):
    base, c1, c2 = clean(base), clean(c1), clean(c2)
    if not base:
        return ""
    if c1 and c2:
        return f"{base}- {c1} & {c2}"
    if c1:
        return f"{base}- {c1}"
    return base


def combine_pockets(ext, internal):
    ext, internal = clean(ext), clean(internal)
    parts = []
    if ext:
        parts.append(f"{ext} External Pocket")
    if internal:
        parts.append(f"{internal} Internal Side Pockets")
    return " and ".join(parts)


class RowAccessor:
    """Look up a nomenclature field key's value in a row, via the field mapping."""

    def __init__(self, row, mapping):
        self.row = row
        self.mapping = mapping

    def __call__(self, field_key):
        col = self.mapping.get(field_key)
        if not col or col not in self.row.index:
            return ""
        return clean(self.row[col])


# ---------------------------------------------------------------------------
# DN / LVN builders
# ---------------------------------------------------------------------------

def build_dn(category, g):
    if category == "Backpacks":
        parts = [g("Brand"), g("Multipack Set"), g("Print or Pattern Type"), g("Surface Styling"),
                 g("Shoulder Strap Type"), g("Features"), g("Technology"), g("Main Trend"), g("Article Type")]
        return join_nonblank(parts)

    if category == "Trolley Bag":
        wheels = g("Number of Wheels")
        wheel_txt = f"{wheels} wheel" if wheels else ""
        parts = [g("Brand"), g("Multipack Set"), g("Print or Pattern Type"), g("Features"),
                 wheel_txt, g("Trolley Size"), g("Article Type")]
        dn = join_nonblank(parts)
        litres = g("Volume in Litres")
        if litres:
            dn = f"{dn}- {litres}L"
        return dn

    if category == "Handbags":
        parts = [g("Brand"), g("Multipack Set"), g("Print or Pattern Type"), g("Material"),
                 g("Size"), g("Surface Styling"), g("Main Trend"), g("Type")]
        return join_nonblank(parts)

    if category == "Laptop Bag":
        parts = [g("Brand"), g("Multipack Set"), g("Print or Pattern Type"), g("Material"),
                 g("Main Trend"), g("Features"), g("Type")]
        dn = join_nonblank(parts)
        laptop_size = g("Laptop Size")
        if laptop_size:
            dn = f"{dn} - {laptop_size} Inches"
        return dn

    return ""


LVN_FIELD_ORDER = {
    "Backpacks": ["Multipack Set", "Print or Pattern Type", "Technology", "Features", "Main Trend", "Article Type"],
    "Trolley Bag": ["Multipack Set", "Features", "Trolley Size", "Article Type"],
    "Handbags": ["Multipack Set", "Size", "Print or Pattern Type", "Type"],
    "Laptop Bag": ["Multipack Set", "Material", "Print or Pattern Type", "Features", "Article Type"],
}

LVN_PREFERENCE_ORDER = {
    "Backpacks": ["Multipack Set", "Technology", "Features", "Main Trend", "Article Type", "Print or Pattern Type"],
    "Trolley Bag": ["Multipack Set", "Trolley Size", "Features", "Article Type"],
    "Handbags": ["Multipack Set", "Size", "Type", "Print or Pattern Type"],
    "Laptop Bag": ["Multipack Set", "Material", "Features", "Article Type", "Print or Pattern Type"],
}


def build_lvn(category, g, max_length=None):
    field_order = LVN_FIELD_ORDER.get(category, [])
    values = {f: g(f) for f in field_order}
    ordered_present = [f for f in field_order if values[f]]

    def render(fields):
        return join_nonblank([values[f] for f in fields])

    lvn = render(ordered_present)

    if max_length and len(lvn) > max_length:
        preference = LVN_PREFERENCE_ORDER.get(category, field_order)
        keep = list(ordered_present)
        # drop lowest-priority fields (from the end of preference order) until it fits
        drop_order = [f for f in reversed(preference) if f in keep]
        for f in drop_order:
            if len(render(keep)) <= max_length:
                break
            keep.remove(f)
        lvn = render(keep)

    return lvn


# ---------------------------------------------------------------------------
# PD (Product Details, bullet-style) builders
# ---------------------------------------------------------------------------

def build_pd(category, g):
    lines = []  # list of (label, separator, value)

    def add(label, value, sep=":"):
        if not is_blank(value):
            lines.append((label, sep, value))

    if category == "Backpacks":
        add("Colour", combine_colour(g("Base Colour"), g("Colour 1"), g("Colour 2")))
        add("Number of Main Compartments", g("Number of Main Compartments"))
        add("Compartment Closure", g("Compartment Closure"))
        add("Features", g("Features"))
        add("Technology", g("Technology"))
        add("Number of External Pockets", g("Number of External Pockets"))
        add("Surface Styling", g("Surface Styling"))
        add("Main Trends", g("Main Trend"))
        add("Warranty", g("Warranty"))

    elif category == "Trolley Bag":
        add("Colour", combine_colour(g("Base Colour"), g("Colour 1"), g("Colour 2")))
        add("Trolley Size", g("Trolley Size"))
        add("Bag Type", g("Bag Type"))
        add("Pattern", g("Print or Pattern Type"))
        add("Laptop Compartment", g("Laptop Compartment"))
        add("Number of Main Compartments", g("Number of Main Compartments"))
        add("Number of Handles", g("Number of Handles"))
        add("Number of Wheels", g("Number of Wheels"))
        add("Features", g("Features"))
        add("Lock Type", g("Lock Type"))
        add("Shell Type", g("Shell Type"))
        add("Pockets", combine_pockets(g("Number of External Pockets"), g("Number of Internal Side Pockets")))
        add("Maximum Carrying Capacity", g("Maximum Carrying Capacity"))
        add("Add Ons", g("Add Ons"))
        add("Warranty", g("Warranty"))

    elif category == "Handbags":
        add("Base Colour", g("Base Colour"))
        add("Print or Pattern Type", g("Print or Pattern Type"))
        add("Type", join_nonblank([g("Surface Styling"), g("Type")]))
        add("Number of Main Compartments", g("Number of Main Compartments"))
        add("Compartment Closure", g("Compartment Closure"))
        add("Number of External Pockets", g("Number of External Pockets"), sep="-")
        add("Number of Inner Pocket", g("Number of Inner Pocket"), sep="-")
        add("Sling Strap", g("Sling Strap"))
        add("Addons", g("Add Ons"))
        add("Warranty", g("Warranty"))

    elif category == "Laptop Bag":
        add("Base Colour", g("Base Colour"))
        add("Print or Pattern Type", g("Print or Pattern Type"))
        add("Number of Main Compartments", g("Number of Main Compartments"))
        add("Laptop Compartment", g("Laptop Compartment"))
        add("Number of External Pockets", g("Number of External Pockets"), sep="-")
        add("Number of Inner Pocket", g("Number of Inner Pocket"), sep="-")
        add("Warranty", g("Warranty"))

    items = [f"{label} {sep} {value}" for label, sep, value in lines]
    items.append("Warranty provided by brand owner/manufacturer :")

    li_html = "".join(f"<li>{item}</li>" for item in items)
    return f"<p><ul>{li_html}</ul></p>"


# ---------------------------------------------------------------------------
# Size & Fit builders
# ---------------------------------------------------------------------------

def build_size_fit(category, g):
    lines = []

    def add(label, value):
        if not is_blank(value):
            lines.append(f"{label}: {value}")

    if category == "Backpacks":
        add("Length", g("Length"))
        add("Height", g("Height"))
        add("Depth", g("Depth"))
        add("Volume in Litres", g("Volume in Litres"))

    elif category == "Trolley Bag":
        add("Height", g("Height"))
        add("Width", g("Width"))
        add("Depth", g("Depth"))
        add("Dead Weight", g("Dead Weight"))
        add("Maximum Carrying Capacity", g("Maximum Carrying Capacity"))
        add("Volume in Litres", g("Volume in Litres"))

    elif category == "Handbags":
        add("Height", g("Height"))
        add("Width", g("Width"))
        add("Depth", g("Depth"))
        add("Volume in Litres", g("Volume in Litres"))

    elif category == "Laptop Bag":
        add("Height", g("Height"))
        add("Width", g("Width"))
        add("Depth", g("Depth"))
        laptop_size = g("Laptop Size")
        if laptop_size:
            lines.append(f'Laptop sleeve can hold "{laptop_size}" laptop')

    return f"<p>{'<br>'.join(lines)}</p>"


def generate_all(category, row, mapping, lvn_max_length=None):
    g = RowAccessor(row, mapping)
    return {
        "DN": build_dn(category, g),
        "LVN": build_lvn(category, g, max_length=lvn_max_length),
        "PD": build_pd(category, g),
        "Size & Fit": build_size_fit(category, g),
    }
