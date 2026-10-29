# -*- coding: utf-8 -*-
"""
Client-specific descriptive flexfields (DFFs) for the standard PO template.

WHY THIS IS A SEPARATE FILE
    The 45 core columns are the same for every client, so they live in
    po_columns.py. DFFs are the opposite: they exist only because one client
    configured them, and the attribute column a segment lands in is decided by
    that client's Fusion setup. Hard-coding a DFF into the standard template
    would make the template wrong for the next client. So each client gets a
    small declaration file, and the template and the converter are both built
    from it.

WHERE THE VALUES COME FROM  (this is the part that must not be guessed)
    Setup and Maintenance > Manage Descriptive Flexfields > search the flexfield
    code below > Edit > Manage Contexts > the context's segments. Each segment
    shows a PROMPT (what the user sees) and a COLUMN NAME (ATTRIBUTE5, say).
    The template caption comes from the prompt; `attribute` MUST be the column
    name. They are unrelated: a segment prompted "Project Code" can perfectly
    well sit in ATTRIBUTE12.

    Flexfield codes for a purchase order:
        header        PO_HEADERS          -> PO_HEADERS_INTERFACE
        line          PO_LINES            -> PO_LINES_INTERFACE
        schedule      PO_LINE_LOCATIONS   -> PO_LINE_LOCATIONS_INTERFACE
        distribution  PO_DISTRIBUTIONS    -> PO_DISTRIBUTIONS_INTERFACE

SPEC FORMAT  (dff/<client>.json)
    {
      "client": "Example Trading Co.",
      "confirmed_from": "who confirmed it, from which instance, on what date",
      "fields": [
        {
          "key": "budget_owner",          # internal, unique, snake_case
          "level": "header",              # header | line | schedule | distribution
          "name": "Budget Owner",         # the segment PROMPT - the client sees this
          "context": "MIGRATION",         # DFF context code, or omit for a global-less DFF
          "attribute": "ATTRIBUTE1",      # the segment's COLUMN NAME from the setup screen
          "type": "text",                 # text | number | date
          "required": false,
          "values": ["FIN", "OPS", "IT"],  # optional: makes it a dropdown
          "hint": "Who owns the budget.",  # short, shown under the caption
          "guide": "Longer help text."
        }
      ]
    }

WHAT IS CHECKED BEFORE A TEMPLATE IS BUILT
    Every one of these is a silent-corruption bug if it slips through, because
    Oracle accepts a well-formed row that carries the wrong value in the wrong
    attribute column.
      * the attribute column really exists on that interface
      * the declared type matches the column family, so a number segment cannot
        be pointed at a text column
      * no two segments at the same level claim the same attribute column
      * at most one context per level - a row can only carry one
        ATTRIBUTE_CATEGORY per interface
      * no key or caption collides with a core template field, which would make
        the converter read the wrong column
"""
import io
import json
import os
import re

# level -> (FBDI sheet, flexfield code shown in Manage Descriptive Flexfields)
LEVELS = {
    "header":       ("PO_HEADERS_INTERFACE", "PO_HEADERS"),
    "line":         ("PO_LINES_INTERFACE", "PO_LINES"),
    "schedule":     ("PO_LINE_LOCATIONS_INTERFACE", "PO_LINE_LOCATIONS"),
    "distribution": ("PO_DISTRIBUTIONS_INTERFACE", "PO_DISTRIBUTIONS"),
}

# declared type -> which attribute columns may hold it.
# The prefixes matter: GLOBAL_ is the country/global DFF, SH_ and ITEM_ are
# separate flexfields that happen to be carried on the same interface sheet.
FAMILY = {
    "text":   re.compile(r"^(GLOBAL_|SH_|ITEM_)?ATTRIBUTE\d+$"),
    "number": re.compile(r"^(GLOBAL_|SH_)?ATTRIBUTE_NUMBER\d+$"),
    "date":   re.compile(r"^(GLOBAL_|SH_)?ATTRIBUTE_DATE\d+$"),
}

CONTEXT_COLUMN = "ATTRIBUTE_CATEGORY"

# a caption this long stops fitting the column, and Excel's inline dropdown
# formula gives up past ~255 characters of joined values
MAX_INLINE_LIST = 240


def _norm(s):
    return re.sub(r"\s+", " ", str(s or "").replace("*", " ")).strip().lower()


def load(path, layout, core_columns):
    """Read and validate a client DFF spec.

    Returns (columns, contexts, notes):
      columns  - column dicts ready to append to the template catalogue
      contexts - {fbdi sheet: context code} for the ATTRIBUTE_CATEGORY column
      notes    - human-readable lines for the build log and the conversion report

    Raises SystemExit with every problem listed at once, rather than stopping at
    the first, so a spec can be fixed in one pass.
    """
    with io.open(path, encoding="utf-8") as fh:
        spec = json.load(fh)

    fields = spec.get("fields") or []
    if not isinstance(fields, list):
        raise SystemExit("DFF spec '%s': 'fields' must be a list." % path)

    errors = []
    if not spec.get("confirmed_from"):
        errors.append(
            "'confirmed_from' is missing. Record who confirmed these attribute "
            "columns, from which instance and when - a DFF pointed at the wrong "
            "attribute loads silently and is found months later.")

    core_keys = {c["key"] for c in core_columns}
    core_caps = {_norm(c["name"]) for c in core_columns}

    claimed = {}          # (level, attribute) -> key
    contexts = {}         # level -> context
    seen_keys = set()
    seen_caps = {}
    columns = []

    for i, f in enumerate(fields, start=1):
        where = "field %d (%s)" % (i, f.get("key") or f.get("name") or "unnamed")

        key = f.get("key")
        level = f.get("level")
        name = f.get("name")
        attr = (f.get("attribute") or "").strip().upper()
        ftype = (f.get("type") or "text").strip().lower()

        if not key or not re.match(r"^[a-z][a-z0-9_]*$", str(key)):
            errors.append("%s: 'key' must be lower_snake_case." % where)
            continue
        if not name:
            errors.append("%s: 'name' is required - use the segment prompt." % where)
            continue
        if level not in LEVELS:
            errors.append("%s: level '%s' is not one of %s."
                          % (where, level, ", ".join(sorted(LEVELS))))
            continue

        sheet, flex_code = LEVELS[level]
        sheet_cols = {str(c).strip().upper() for c in layout["sheets"].get(sheet, [])}

        # ---- the attribute column must exist, and hold this kind of value
        if ftype not in FAMILY:
            errors.append("%s: type '%s' is not one of %s."
                          % (where, ftype, ", ".join(sorted(FAMILY))))
        elif not attr:
            errors.append("%s: 'attribute' is required. Take the segment's Column "
                          "Name from Manage Descriptive Flexfields > %s."
                          % (where, flex_code))
        elif attr not in sheet_cols:
            errors.append("%s: '%s' is not a column on %s." % (where, attr, sheet))
        elif "TIMESTAMP" in attr:
            errors.append(
                "%s: ATTRIBUTE_TIMESTAMP columns are not supported yet - only "
                "text, number and date segments are. If the client really has a "
                "timestamp segment, load the date part into an ATTRIBUTE_DATEn "
                "column and raise it with the team." % where)
        elif not FAMILY[ftype].match(attr):
            errors.append(
                "%s: type '%s' cannot go in '%s'. A %s segment belongs in %s."
                % (where, ftype, attr, ftype,
                   {"text": "ATTRIBUTEn", "number": "ATTRIBUTE_NUMBERn",
                    "date": "ATTRIBUTE_DATEn"}[ftype]))
        elif (level, attr) in claimed:
            errors.append("%s: '%s' at %s level is already used by '%s'. Two "
                          "segments cannot share one attribute column."
                          % (where, attr, level, claimed[(level, attr)]))
        else:
            claimed[(level, attr)] = key

        # ---- one context per level: a row carries a single ATTRIBUTE_CATEGORY
        ctx = f.get("context")
        if ctx:
            if CONTEXT_COLUMN not in sheet_cols:
                errors.append("%s: %s has no %s column, so a context cannot be "
                              "set at this level." % (where, sheet, CONTEXT_COLUMN))
            elif level in contexts and contexts[level] != ctx:
                errors.append("%s: context '%s' conflicts with '%s' already "
                              "declared at %s level. One row can only carry one "
                              "context per level - split the load, or use a "
                              "single context with all the segments."
                              % (where, ctx, contexts[level], level))
            else:
                contexts[level] = ctx

        # ---- must not collide with the standard template
        if key in core_keys:
            errors.append("%s: key '%s' is already a core template field." % (where, key))
        if key in seen_keys:
            errors.append("%s: duplicate key '%s'." % (where, key))
        seen_keys.add(key)

        cap = _norm(name)
        if cap in core_caps:
            errors.append("%s: caption '%s' is already a core column. The converter "
                          "matches by caption, so it would read the wrong column."
                          % (where, name))
        if cap in seen_caps:
            errors.append("%s: caption '%s' is already used by '%s'."
                          % (where, name, seen_caps[cap]))
        seen_caps[cap] = key

        values = f.get("values")
        if values is not None and (not isinstance(values, list) or not values):
            errors.append("%s: 'values' must be a non-empty list when present." % where)
            values = None

        columns.append(_column(f, key, name, level, attr, ftype, values))

    if errors:
        raise SystemExit(
            "DFF spec '%s' has %d problem(s):\n  - %s"
            % (path, len(errors), "\n  - ".join(errors)))

    notes = ["client: %s" % (spec.get("client") or "(unnamed)"),
             "confirmed from: %s" % spec.get("confirmed_from")]
    for level in sorted({c["_dff_level"] for c in columns}):
        at = [c for c in columns if c["_dff_level"] == level]
        notes.append("%s: %d field(s) -> %s%s"
                     % (level, len(at), ", ".join(c["tgt"][1] for c in at),
                        ", context %s" % contexts[level] if level in contexts else
                        ", no context (global segments)"))

    return columns, {LEVELS[l][0]: c for l, c in contexts.items()}, notes


def _column(f, key, name, level, attr, ftype, values):
    """Turn one declared segment into a template column definition."""
    sheet, _ = LEVELS[level]
    required = bool(f.get("required"))

    if values:
        dv = "DFF_" + key
    elif ftype == "number":
        dv = "num"
    elif ftype == "date":
        dv = "date"
    else:
        dv = None

    fmt = {"text": "Text", "number": "Number", "date": "Date (YYYY-MM-DD)"}[ftype]
    hint = f.get("hint") or ("Pick from the list." if values else
                             "Client-specific field.")
    guide = f.get("guide") or (
        "%s. Client-specific field held in the %s descriptive flexfield (%s)."
        % (name, level, attr))

    # wide enough for the caption on two lines AND the hint on three, which is
    # what the layout check in build_template.py enforces
    import math
    need_name = math.ceil((len(name) + 4) / 2.0) + 1
    need_hint = math.ceil(len(hint) / 3.0) + 1
    width = max(12, min(30, int(max(need_name, need_hint))))
    return dict(
        key=key, grp="DFF", name=name, req=required, w=width, dv=dv,
        tgt=(sheet, attr), fmt=fmt, ex="",
        hint=hint, guide=guide, adv=False, adv_why="",
        _dff_level=level, _dff_values=values or [],
    )


def example_spec_path(root):
    return os.path.join(root, "dff", "example.json")
