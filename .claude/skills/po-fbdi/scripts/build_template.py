# -*- coding: utf-8 -*-
"""
Builds "templates/Standard - Purchase Order Data Template.xlsx" from
po_columns.COLUMNS.

  Sheet 1  Overview & History  - branded cover, revision history, how-to-fill
  Sheet 2  PO Data             - THE single sheet the client fills in

Row layout of the PO Data sheet
  1  coloured group band
  2  field name    ( *  mandatory,  (*)  mandatory in some cases )
  3  short hint
  4  data type
  5  example row, in yellow, deleted by the client
  6+ client data

Run:  python .claude/skills/po-fbdi/scripts/build_template.py
"""
import math
import os
import sys

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import po_columns as P  # noqa: E402

SCRIPTS = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(SCRIPTS)
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(SKILL)))
TEMPLATES = os.path.join(ROOT, "templates")
LOGO = os.path.join(SKILL, "assets", "speridian-logo.jpeg")
OUT = os.path.join(TEMPLATES, "Standard - Purchase Order Data Template.xlsx")

# ------------------------------------------------------------------ palette
NAVY = "1F3864"          # section bands
BLUE = "2E6DB4"          # table header fill
BLUE_SOFT = "D9E2F3"
INK = "1F2937"
GREY = "595959"

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

REQ_FONT = Font(name="Calibri", size=10.5, bold=True, color="9C0006")
CON_FONT = Font(name="Calibri", size=10.5, bold=True, color="1F4E79")
OPT_FONT = Font(name="Calibri", size=10.5, color="404040")
HINT_FONT = Font(name="Calibri", size=8.5, color="404040")
FMT_FONT = Font(name="Calibri", size=8, italic=True, color="808080")
EX_FONT = Font(name="Calibri", size=10, italic=True, color="7F6000")
EX_FILL = PatternFill("solid", fgColor="FFF9E6")
HINT_FILL = PatternFill("solid", fgColor="FAFAFA")
FMT_FILL = PatternFill("solid", fgColor="F2F2F2")
CHK_FILL = PatternFill("solid", fgColor="F2F2F2")

WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# row heights, sized to the longest text at each column width (checked below)
H_GROUP, H_NAME, H_HINT, H_FMT, H_EX = 20, 30, 40, 13, 18


def a1(key, row):
    return "%s%d" % (get_column_letter(P.col_index(key)), row)


def present(key):
    """Is this field on the sheet in the profile being built?"""
    return any(c["key"] == key for c in P.COLUMNS)


def caption(col):
    """Field name with the mandatory marker."""
    if col["req"] is True:
        return col["name"] + " *"
    if col["req"] == "C":
        return col["name"] + " (*)"
    return col["name"]


def band(ws, row, first, last, text):
    ws.merge_cells(start_row=row, start_column=first, end_row=row, end_column=last)
    c = ws.cell(row, first, text)
    c.font = Font(size=10.5, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for i in range(first, last + 1):
        ws.cell(row, i).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[row].height = 19


def col_range(first_key, last_key):
    """'AL to AV' - real column letters, so instructions cannot drift."""
    return "%s to %s" % (get_column_letter(P.col_index(first_key)),
                         get_column_letter(P.col_index(last_key)))


def split_rule_text():
    """Rule 10 - how to split a line, worded for the profile being built.

    On the lean sheet Schedule No. and Distribution No. are not shown at all;
    the converter works out the numbering from what actually differs between the
    repeated rows. So the client instruction becomes "copy the row and change
    what is different", with no numbering to get wrong.
    """
    if present("schedule_number"):
        sched = col_range("schedule_number", "sched_note_receiver")
        dist = col_range("distribution_number", "expenditure_org")
        qty = get_column_letter(P.col_index("schedule_quantity"))
        dqty = get_column_letter(P.col_index("distribution_quantity"))
        return (
            "Only needed for a line that is delivered in parts, or charged to more than one account. "
            "SELECT THE WHOLE ROW, COPY IT, PASTE IT ON THE NEXT ROW, then change only the columns "
            "listed below - leave everything else exactly as it was, including PO, supplier, item, UOM "
            "and price. Several delivery dates or locations: change columns %s (Schedule No., dates, "
            "ship-to) and put each part's quantity in column %s. Several charge accounts: change columns "
            "%s (Distribution No., account, project) and put each part's quantity in column %s. The "
            "parts must add up to the line total, and column A will tell you if they do not."
            % (sched, qty, dist, dqty))

    date_col = get_column_letter(P.col_index("need_by_date"))
    ship_col = get_column_letter(P.col_index("ship_to_location"))
    qty_col = get_column_letter(P.col_index("quantity"))
    acct_col = get_column_letter(P.col_index("charge_account"))
    return (
        "Only needed for a line delivered in parts, or charged to more than one account. SELECT THE "
        "WHOLE ROW, COPY IT, PASTE IT ON THE NEXT ROW, and keep the same PO Number and Line Number. "
        "Then change only what is actually different: for a part delivered on another date or to "
        "another place, change Need-by Date (column %s) or Ship-to Location (column %s); for a part "
        "charged elsewhere, change Charge Account (column %s). On EVERY copy put the quantity for THAT "
        "PART ONLY in column %s - the parts are added up for you, so do not repeat the full line "
        "quantity on each row. There is nothing to number: the schedule and distribution numbers are "
        "worked out from the rows you provide."
        % (date_col, ship_col, acct_col, qty_col))


SPLIT_RULE = None      # filled in main(), once P.COLUMNS is final


# ------------------------------------------------------------------ sheet 1
def build_overview(ws):
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJ", [23, 20, 20, 12, 44, 3, 20, 18, 14, 18]):
        ws.column_dimensions[col].width = w

    # ---- masthead: logo left, title centred
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 26
    if os.path.exists(LOGO):
        img = XLImage(LOGO)
        img.width, img.height = 172, 44
        img.anchor = "A1"
        ws.add_image(img)

    ws.merge_cells("C1:H1")
    t = ws.cell(1, 3, "Open Purchase Orders")
    t.font = Font(size=16, bold=True, color=INK)
    t.alignment = Alignment(horizontal="center", vertical="bottom")
    ws.merge_cells("C2:H2")
    t2 = ws.cell(2, 3, "Data Template")
    t2.font = Font(size=13, bold=True, color=BLUE)
    t2.alignment = Alignment(horizontal="center", vertical="top")

    ws.merge_cells("A3:J3")
    ws.cell(3, 1).fill = PatternFill("solid", fgColor=BLUE)
    ws.row_dimensions[3].height = 4

    # ---- identification block
    facts = [
        ("Department:", "Procurement"),
        ("Module:", "Oracle Fusion Cloud - Procurement"),
        ("Object:", "Purchase Orders (Open / Historical)"),
        ("Oracle Interface:", "Purchase Order Import (FBDI)"),
        ("Template Version:", "1.0"),
    ]
    for i, (k, v) in enumerate(facts):
        r = 5 + i
        c = ws.cell(r, 1, k)
        c.font = Font(bold=True, color=INK)
        c.fill = PatternFill("solid", fgColor=BLUE_SOFT)
        c.border = BOX
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        d = ws.cell(r, 2, v)
        d.border = BOX
        ws.cell(r, 3).border = BOX

    # ---- revision history + sign off, side by side
    band(ws, 11, 1, 5, "REVISION HISTORY")
    band(ws, 11, 7, 10, "PREPARED AND APPROVED BY")

    def table_head(row, labels, first):
        for i, h in enumerate(labels):
            c = ws.cell(row, first + i, h)
            c.font = Font(bold=True, size=10, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=BLUE)
            c.border = BOX
            c.alignment = CENTER
        ws.row_dimensions[row].height = 18

    table_head(12, ["Date", "Prepared By", "Reviewed By", "Version", "Comments"], 1)
    table_head(12, ["Prepared By", "Designation", "Date", "Approved By"], 7)

    for j in range(3):
        r = 13 + j
        ws.row_dimensions[r].height = 17
        for i in list(range(1, 6)) + list(range(7, 11)):
            ws.cell(r, i).border = BOX
            ws.cell(r, i).alignment = WRAP_TOP
    ws.cell(13, 4, "1.0")
    ws.cell(13, 5, "Initial version - shared with client.")

    # ---- overview
    r = 17
    band(ws, r, 1, 10, "OVERVIEW")
    r += 1
    for t in [
        "This workbook collects every purchase order that must be migrated into Oracle Fusion Cloud.",
        "Everything is entered on ONE sheet - the 'PO Data' sheet. You do not need to know anything about "
        "Oracle: fill the sheet in, and we convert it into the Oracle upload files for you.",
        "Columns are grouped and colour coded: purchase order, supplier, line, shipment, charge account and "
        "legacy reference. Mandatory fields are marked in the field name.",
    ]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        c = ws.cell(r, 1, t)
        c.alignment = WRAP_TOP
        ws.row_dimensions[r].height = 16 if len(t) < 100 else 30
        r += 1

    # ---- how to fill
    r += 1
    band(ws, r, 1, 10, "HOW TO FILL THE 'PO DATA' SHEET")
    r += 1
    rules = [
        ("1.  One row = one PO line",
         "Enter one row for every line of every purchase order. Start entering your data in row 6. The sheet is "
         "ready for %d rows - if you need more, select the last row, copy it, and paste it down as far as you "
         "need so the checks in column A come with it." % P.DATA_ROWS),
        ("2.  Repeat the PO information",
         "A PO with 3 lines takes 3 rows, and the blue PURCHASE ORDER and SUPPLIER columns carry the SAME "
         "values on all 3 rows. Do not merge cells and do not leave them blank."),
        ("3.  Delete the yellow example row",
         "Row 5 is an example so you can see what is expected. Delete that row before sending the file back."),
        ("4.  Watch the Row Check column",
         "Column A checks each row automatically. It stays BLANK until you type something into the row, then "
         "turns green and shows OK, or turns red and lists what is still missing. It also warns you if the "
         "same PO line has been entered twice. Nothing to type in column A. Please return the file only when "
         "every row with data shows OK."),
        ("5.  Goods lines and service lines",
         "For Goods lines fill Quantity, Unit Price and UOM, and leave Line Amount blank. For Fixed Price "
         "Services and Rate Based Services fill Line Amount only."),
        ("6.  Open purchase orders",
         "Include only POs that are still open - approved, not fully received or invoiced. In the Quantity "
         "column enter the quantity STILL TO BE RECEIVED, not the original quantity. Use the grey LEGACY "
         "REFERENCE columns to tell us the original and already received quantities."),
        ("7.  Dates",
         "Always type dates as YYYY-MM-DD, for example 2026-03-10."),
        ("8.  Numbers",
         "No thousands separators, no currency symbols and no words such as 'approx'. Use a full stop as the "
         "decimal separator."),
        ("9.  Names must match Oracle",
         "Supplier, buyer, requester, location, payment terms, category and unit of measure must be spelled "
         "exactly as they are set up in Oracle. If you are unsure, fill in what you know and add a note in "
         "the Remarks column - we will check it."),
        ("10. Split deliveries and split accounts",
         SPLIT_RULE),
        ("11. Do not change the layout",
         "Please do not rename, delete, reorder or hide columns, and do not add sheets. If you need an extra "
         "field, tell us and we will extend the template."),
    ]
    for title, body in rules:
        c = ws.cell(r, 1, title)
        c.font = Font(bold=True, color=NAVY, size=10.5)
        c.alignment = WRAP_TOP
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        c2 = ws.cell(r, 2, body)
        c2.alignment = WRAP_TOP
        ws.row_dimensions[r].height = 15 * (1 + len(body) // 118)
        r += 1

    # ---- legend
    r += 1
    band(ws, r, 1, 10, "HOW TO READ THE COLUMN HEADINGS")
    r += 1
    for text, font in [
        ("Field name *", REQ_FONT),
        ("Field name (*)", CON_FONT),
        ("Field name", OPT_FONT),
    ]:
        c = ws.cell(r, 1, text)
        c.font = font
        c.border = BOX
        c.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        note = {
            "Field name *": "Mandatory. Every row must have a value.",
            "Field name (*)": "Mandatory in certain cases only - the hint under the heading says when.",
            "Field name": "Optional. Fill it in if you track it.",
        }[text]
        d = ws.cell(r, 2, note)
        d.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 16
        r += 1

    r += 1
    for gkey, (label, band_colour, sub) in P.GROUPS.items():
        n = sum(1 for x in P.COLUMNS if x["grp"] == gkey)
        c = ws.cell(r, 1, label)
        c.fill = PatternFill("solid", fgColor=band_colour)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        d = ws.cell(r, 7, "%d column%s" % (n, "" if n == 1 else "s"))
        d.fill = PatternFill("solid", fgColor=sub)
        d.alignment = Alignment(vertical="center", indent=1)
        d.border = BOX
        ws.row_dimensions[r].height = 16
        r += 1

    ws.print_area = "A1:J%d" % (r + 1)


# ------------------------------------------------------------------ Row Check
def duplicate_test(row):
    """Flag a row whose PO + line + schedule + distribution already exists.

    A gap in the numbering (1, 2, 4) is deliberately NOT flagged - Oracle does
    not need line numbers to be contiguous. A repeat is the real error: it means
    the client meant to split a line and forgot to number the split.

    Concatenating the four keys is what makes blanks behave. COUNTIFS with a
    reference to an empty cell as its criteria matches zero and blank alike, so
    it would report false duplicates on every row where Schedule No. is left
    empty - which is most of them.
    """
    last = P.ROW_DATA + P.DATA_ROWS - 1
    # On the wide sheet the client numbers the splits, so the key is the four
    # numbers. On the lean sheet there are no numbers to compare, so the key is
    # what the converter itself uses to tell splits apart - otherwise a genuine
    # split (same line, different delivery date) would read as a duplicate.
    if present("schedule_number") or present("distribution_number"):
        keys = [k for k in ("po_number", "line_number",
                            "schedule_number", "distribution_number") if present(k)]
    else:
        keys = ["po_number", "line_number"]
        keys += [k for k in P.SCHEDULE_KEYS + P.DIST_KEYS if present(k)]

    def rng(key):
        letter = get_column_letter(P.col_index(key))
        return "$%s$%d:$%s$%d" % (letter, P.ROW_DATA, letter, last)

    # COUNTIFS, not SUMPRODUCT. SUMPRODUCT has to build a concatenated string
    # array of every row for every row, which measured 180 ms per keystroke at
    # 1000 rows. COUNTIFS is native and an order of magnitude cheaper.
    # The `&""` on each criterion is what makes blanks behave: a bare reference
    # to an empty cell makes COUNTIFS match zero and blank alike, so rows with
    # no Schedule No. would all look like duplicates of each other.
    args = ", ".join('%s, %s&""' % (rng(k), a1(k, row)) for k in keys)
    return 'IF(COUNTIFS(%s)>1,"Duplicate row; ","")' % args


def check_formula(row):
    """Column-A formula reporting what is missing on `row`."""
    parts = []
    for label, keys in P.MANDATORY_GROUPS.items():
        cells = ",".join(a1(k, row) for k in keys)
        parts.append('IF(%d-COUNTA(%s)>0,"%s; ","")' % (len(keys), cells, label))

    def cond(keys, template, label):
        """Add a conditional test only if every column it needs is on the sheet."""
        if not all(present(k) for k in keys):
            return None
        cells = tuple(a1(k, row) for k in keys)
        return template % (cells + (label,))

    maybe = [
        cond(("item", "category"),
             'IF(AND(%s="",%s=""),"%s; ","")', "Category Name"),
        cond(("line_type", "quantity", "unit_price", "uom"),
             'IF(AND(%s="Goods",OR(%s="",%s="",%s="")),"%s; ","")', "Qty/Price/UOM"),
        cond(("line_type", "line_type", "line_amount"),
             'IF(AND(%s<>"",%s<>"Goods",%s=""),"%s; ","")', "Line Amount"),
        cond(("comm_method", "supplier_email"),
             'IF(AND(UPPER(%s)="EMAIL",%s=""),"%s; ","")', "Supplier Email"),
        cond(("destination_type", "ship_to_org"),
             'IF(AND(UPPER(%s)="INVENTORY",%s=""),"%s; ","")', "Ship-to Organization"),
        cond(("conversion_rate", "rate_type"),
             'IF(AND(%s<>"",%s=""),"%s; ","")', "Rate Type"),
    ]
    parts += [m for m in maybe if m] + [duplicate_test(row)]
    msg = "&".join(parts)
    scan = "%s%d:%s%d" % (get_column_letter(2), row, get_column_letter(len(P.COLUMNS)), row)
    return '=IF(COUNTA({s})=0,"",IF(({m})="","OK","Fix: "&({m})))'.format(s=scan, m=msg)


LONG_LISTS = {}          # list key -> ((col letter, first row), (col letter, last row))


def build_lists(ws):
    """Hidden sheet holding any value set too long for an inline dropdown."""
    ws.sheet_state = "hidden"
    col = 0
    for kind, values in sorted(P.LISTS.items()):
        if len(",".join(values)) <= 240:
            continue
        col += 1
        letter = get_column_letter(col)
        ws.cell(1, col, kind).font = Font(bold=True, size=9)
        for i, v in enumerate(values, start=2):
            ws.cell(i, col, v)
        ws.column_dimensions[letter].width = 24
        LONG_LISTS[kind] = ((letter, 2), (letter, len(values) + 1))
    return bool(LONG_LISTS)


# ------------------------------------------------------------------ sheet 2
def build_data(ws):
    ws.sheet_properties.tabColor = "4F6228"
    ncols = len(P.COLUMNS)
    last_row = P.ROW_DATA + P.DATA_ROWS - 1

    # ---- group band
    start = 1
    for i in range(1, ncols + 2):
        cur = P.COLUMNS[i - 1]["grp"] if i <= ncols else None
        prev = P.COLUMNS[start - 1]["grp"]
        if cur != prev:
            label, band_colour, _sub = P.GROUPS[prev]
            if i - 1 > start:
                ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=i - 1)
            c = ws.cell(1, start, label)
            c.font = Font(size=10.5, bold=True, color="FFFFFF")
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            for k in range(start, i):
                ws.cell(1, k).fill = PatternFill("solid", fgColor=band_colour)
            start = i
    ws.row_dimensions[1].height = H_GROUP

    # ---- headings, hints, formats, example
    for idx, col in enumerate(P.COLUMNS, start=1):
        letter = get_column_letter(idx)
        _label, _band, sub = P.GROUPS[col["grp"]]
        ws.column_dimensions[letter].width = col["w"]

        c = ws.cell(P.ROW_NAME, idx, caption(col))
        c.font = REQ_FONT if col["req"] is True else CON_FONT if col["req"] == "C" else OPT_FONT
        c.fill = PatternFill("solid", fgColor=sub)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BOX

        h = ws.cell(P.ROW_HINT, idx, col["hint"])
        h.font = HINT_FONT
        h.fill = HINT_FILL
        h.alignment = WRAP_TOP
        h.border = BOX

        f = ws.cell(P.ROW_FORMAT, idx, col["fmt"])
        f.font = FMT_FONT
        f.fill = FMT_FILL
        f.alignment = Alignment(vertical="center")
        f.border = BOX

        e = ws.cell(P.ROW_EXAMPLE, idx, col["ex"] if col["ex"] != "" else None)
        e.font = EX_FONT
        e.fill = EX_FILL
        e.border = BOX
        if col["dv"] == "date":
            e.number_format = "yyyy-mm-dd"

    ws.cell(P.ROW_EXAMPLE, 1, "EXAMPLE ROW - DELETE IT")
    ws.cell(P.ROW_EXAMPLE, 1).font = Font(size=9, bold=True, italic=True, color="7F6000")
    ws.row_dimensions[P.ROW_NAME].height = H_NAME
    ws.row_dimensions[P.ROW_HINT].height = H_HINT
    ws.row_dimensions[P.ROW_FORMAT].height = H_FMT
    ws.row_dimensions[P.ROW_EXAMPLE].height = H_EX

    # ---- data area
    money = {"unit_price", "line_amount", "conversion_rate"}
    for row in range(P.ROW_DATA, last_row + 1):
        for idx, col in enumerate(P.COLUMNS, start=1):
            c = ws.cell(row, idx)
            c.border = BOX
            if col["dv"] == "date":
                c.number_format = "yyyy-mm-dd"
            elif col["key"] in money:
                c.number_format = "#,##0.00"
            elif col["dv"] == "num":
                c.number_format = "#,##0.####"
        chk = ws.cell(row, 1, check_formula(row))
        chk.fill = CHK_FILL
        chk.font = Font(size=9, bold=True)
        chk.alignment = WRAP_TOP

    # Freeze the heading ROWS ONLY - no column split.
    #
    # A row+column freeze gives Excel four panes. If the workbook is open while
    # the file is rewritten underneath it, or the freeze is toggled off and on,
    # those panes can end up scrolled independently and Excel draws the heading
    # block up to four times: columns A-B repeated beside themselves and rows
    # 1-5 repeated below themselves. A rows-only freeze has two panes and no
    # vertical split bar, so that cannot happen. It also keeps more of the sheet
    # visible when scrolling right.
    first_data = "A%d" % P.ROW_DATA
    ws.freeze_panes = first_data

    # Match Excel's own XML exactly. Excel never puts topLeftCell on
    # <sheetView> when a pane exists - it pins the view origin to A1, so the
    # scrolling pane restarts at row 1 - and every <selection> needs an active
    # cell inside its own pane, or Excel falls back to A1 and scrolls the
    # scrolling pane up to reveal it. Either one duplicates the headings.
    ws.sheet_view.topLeftCell = None
    corners = {"bottomLeft": first_data, "bottomRight": first_data}
    for sel in ws.sheet_view.selection:
        cell = corners.get(sel.pane)
        if cell:
            sel.activeCell = cell
            sel.sqref = cell

    # ---- validation: dropdowns plus a hover message carrying the full guidance
    for idx, col in enumerate(P.COLUMNS, start=1):
        letter = get_column_letter(idx)
        rng = "%s%d:%s%d" % (letter, P.ROW_DATA, letter, last_row)
        kind = col["dv"]
        if kind in P.LISTS and kind in LONG_LISTS:
            # Excel gives up on an inline list past ~255 characters, so a long
            # value set is parked on the hidden Lists sheet and referenced.
            first, last = LONG_LISTS[kind]
            dv = DataValidation(type="list", allow_blank=True,
                                formula1="='%s'!$%s$%d:$%s$%d"
                                         % (P.LISTS_SHEET, first[0], first[1],
                                            last[0], last[1]),
                                showDropDown=False,
                                errorTitle="Not an allowed value",
                                error="Pick a value from the drop-down list.")
        elif kind in P.LISTS:
            dv = DataValidation(type="list", allow_blank=True,
                                formula1='"%s"' % ",".join(P.LISTS[kind]),
                                showErrorMessage=True,
                                errorTitle="Please pick a value from the list",
                                error="Allowed values: " + ", ".join(P.LISTS[kind]))
        elif kind == "date":
            dv = DataValidation(type="date", operator="greaterThan",
                                formula1="DATE(1990,1,1)", allow_blank=True,
                                showErrorMessage=True, errorTitle="Enter a real date",
                                error="Type the date as YYYY-MM-DD, for example 2026-03-10.")
        elif kind == "num":
            dv = DataValidation(type="decimal", operator="greaterThanOrEqual",
                                formula1="0", allow_blank=True,
                                showErrorMessage=True, errorTitle="Enter a number",
                                error="Enter a number without thousands separators or currency symbols.")
        else:
            dv = DataValidation(allow_blank=True, showErrorMessage=False)
        # No hover prompt. Excel draws it as a floating box that covers four or
        # five cells to the right of whatever you click, and row 3 already
        # carries the same guidance permanently. The long version lives on the
        # Overview sheet and in the mapping guide.
        dv.showInputMessage = False
        ws.add_data_validation(dv)
        dv.add(rng)

    # ---- conditional formatting
    rng = "A%d:A%d" % (P.ROW_DATA, last_row)
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=['LEFT($A%d,2)="OK"' % P.ROW_DATA],
                         fill=PatternFill("solid", fgColor="C6EFCE"),
                         font=Font(color="006100", bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=['AND($A%d<>"",LEFT($A%d,2)<>"OK")' % (P.ROW_DATA, P.ROW_DATA)],
                         fill=PatternFill("solid", fgColor="FFC7CE"),
                         font=Font(color="9C0006", bold=True)))

    po_cell = "$%s%d" % (get_column_letter(P.col_index("po_number")), P.ROW_DATA)
    miss_fill = PatternFill("solid", fgColor="FFEBEB")
    for keys in P.MANDATORY_GROUPS.values():
        for k in keys:
            letter = get_column_letter(P.col_index(k))
            ws.conditional_formatting.add(
                "%s%d:%s%d" % (letter, P.ROW_DATA, letter, last_row),
                FormulaRule(formula=['AND(%s<>"",%s%d="")' % (po_cell, letter, P.ROW_DATA)],
                            fill=miss_fill))

    ws.print_title_rows = "1:%d" % P.ROW_NAME
    ws.page_setup.orientation = "landscape"


# ------------------------------------------------------------------ checks
def check_fits():
    """Warn if any heading or hint needs more lines than its row height allows."""
    problems = []
    for col in P.COLUMNS:
        cpl = max(int(col["w"]) - 1, 4)
        name_lines = math.ceil(len(caption(col)) / cpl)
        hint_lines = math.ceil(len(col["hint"]) / cpl)
        if name_lines > H_NAME // 14:
            problems.append("heading '%s' needs %d lines, row 2 fits %d"
                            % (caption(col), name_lines, H_NAME // 14))
        if hint_lines > H_HINT // 11:
            problems.append("hint of '%s' needs %d lines, row 3 fits %d"
                            % (col["name"], hint_lines, H_HINT // 11))
    return problems


# ------------------------------------------------------------------ main
def main():
    global SPLIT_RULE, OUT
    import argparse
    ap = argparse.ArgumentParser(description="Build the client PO data template.")
    ap.add_argument("--dff", metavar="PATH",
                    help="client descriptive-flexfield spec (dff/<client>.json). "
                         "Adds an 'Additional fields' group and names the output "
                         "after the client.")
    ap.add_argument("--all", action="store_true",
                    help="include the advanced fields Oracle normally derives "
                         "(writes a separate '(all fields)' workbook)")
    args = ap.parse_args()
    if args.all:
        P.set_profile("full")
        OUT = OUT.replace(".xlsx", " (all fields).xlsx")
    else:
        P.set_profile("lean")

    if args.dff:
        import json
        import po_dff
        layout = json.load(open(os.path.join(SKILL, "assets", "fbdi_layout.json"),
                                encoding="utf-8"))
        cols, contexts, notes = po_dff.load(args.dff, layout, P.CATALOGUE)
        P.attach_dff(cols, contexts)
        client = json.load(open(args.dff, encoding="utf-8")).get("client")
        if client:
            safe = "".join(ch for ch in client if ch.isalnum() or ch in " -_").strip()
            OUT = OUT.replace("Standard - ", "%s - " % safe)
        for n in notes:
            print("  dff | " + n)
    SPLIT_RULE = split_rule_text()
    problems = check_fits()
    for p in problems:
        print("WARNING: %s" % p)

    wb = Workbook()
    build_overview(wb.active)
    wb.active.title = "Overview & History"
    # Lists first: build_data() needs LONG_LISTS populated before it can point a
    # drop-down at the hidden sheet.
    lists_ws = wb.create_sheet(P.LISTS_SHEET)
    has_lists = build_lists(lists_ws)
    build_data(wb.create_sheet(P.DATA_SHEET))
    if not has_lists:
        wb.remove(lists_ws)
    else:
        wb.move_sheet(lists_ws, offset=1)
    wb.active = 0
    os.makedirs(TEMPLATES, exist_ok=True)

    # Heal a previous fallback. If an earlier run could not overwrite the real
    # template because it was open in Excel, it left a "(rebuilt)" copy beside
    # it. Two files with almost the same name is how the wrong one gets sent to
    # a client, so collapse them as soon as the real name is writable again.
    fallback = OUT.replace(".xlsx", " (rebuilt).xlsx")
    if os.path.exists(fallback):
        try:
            with open(OUT, "ab"):
                pass
            os.remove(fallback)
            print("removed the stale '%s' left by an earlier run"
                  % os.path.basename(fallback))
        except (PermissionError, OSError):
            pass

    # the template is usually open in Excel while we iterate on it; write beside
    # it rather than failing, and say so
    target = OUT
    try:
        wb.save(target)
    except PermissionError:
        target = OUT.replace(".xlsx", " (rebuilt).xlsx")
        wb.save(target)
        print("NOTE: %s is open in Excel, so it could not be replaced."
              % os.path.basename(OUT))
        print("      Close it and re-run, or use the file written below.")

    print("written: %s" % target)
    print("columns: %d   data rows: %d   layout warnings: %d"
          % (len(P.COLUMNS), P.DATA_ROWS, len(problems)))


if __name__ == "__main__":
    main()
