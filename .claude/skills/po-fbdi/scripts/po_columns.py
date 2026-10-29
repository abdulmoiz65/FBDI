# -*- coding: utf-8 -*-
"""
Single source of truth for the Standard Purchase Order Data Template.

Every column of the client-facing sheet is declared once here, together with the
Oracle FBDI target (interface sheet + column header) it feeds. build_template.py
renders the workbook from this list; convert_to_fbdi.py reads a filled workbook
and uses the same list to populate the Oracle FBDI template.

Grain of the client sheet: ONE ROW = ONE PO LINE / SCHEDULE / DISTRIBUTION
  - header fields are repeated on every row of the same PO Number
  - to split a line across dates or ship-to locations, repeat the line and
    number the schedules 1, 2, 3 ...
  - to split a schedule across accounts, repeat the schedule and number the
    distributions 1, 2, 3 ...
"""

# ---------------------------------------------------------------- layout rows
ROW_GROUP   = 1   # merged, coloured group band
ROW_NAME    = 2   # field name  ( * = mandatory )
ROW_HINT    = 3   # short hint shown under the field name
ROW_GUIDE   = 3   # alias kept for older callers
ROW_FORMAT  = 4   # data type & length
ROW_EXAMPLE = 5   # sample row - client deletes it
ROW_DATA    = 6   # first client row
DATA_ROWS   = 1000   # pre-formatted rows; the client copies the last row down for more

DATA_SHEET  = "PO Data"
LISTS_SHEET = "Lists"

# ---------------------------------------------------------------- groups
# key : (label, band colour, field-name fill)
GROUPS = {
    "CHK": ("CHECK", "404040", "F2F2F2"),
    "PO":  ("1. PURCHASE ORDER  (repeat on every line of the same PO)", "1F4E79", "D9E2F3"),
    "SUP": ("2. SUPPLIER", "31859B", "DAEEF3"),
    "LIN": ("3. PO LINE", "4F6228", "E2EFDA"),
    "SHP": ("4. SHIPMENT / SCHEDULE", "BF8F00", "FFF2CC"),
    "ACC": ("5. CHARGE ACCOUNT / DISTRIBUTION", "7030A0", "E4D9F2"),
    "REF": ("6. LEGACY REFERENCE  (information only - not loaded into Oracle)", "595959", "EDEDED"),
    "DFF": ("7. ADDITIONAL FIELDS  (specific to your organisation)", "C00000", "FCE4E4"),
}

# ---------------------------------------------------------------- value lists
LISTS = {
    "DocType":     ["STANDARD"],
    "LineType":    ["Goods", "Fixed Price Services", "Rate Based Services"],
    "DestType":    ["EXPENSE", "INVENTORY", "DROP SHIP"],
    "CommMethod":  ["NONE", "EMAIL", "PRINT", "FAX"],
    "YN":          ["Y", "N"],
    "MatchOption": ["Order", "Receipt"],
}

# ---------------------------------------------------------------- columns
# key   : internal name used by the converter
# grp   : group key
# name  : caption shown to the client
# req   : True = mandatory, "C" = conditionally mandatory, False = optional
# hint  : short row-3 text, must fit the column width
# guide : full help text, used for the hover message and the docs
# fmt   : row-4 data type
# ex    : row-5 example value
# w     : column width
# dv    : key of LISTS, or "date" / "num", or None
# tgt   : (fbdi sheet, fbdi column header), or None when not loaded directly
COLUMNS = [
    dict(key="check", grp="CHK", name="Row Check", req=False, w=24, dv=None, tgt=None,
         fmt="Automatic", ex="",
         hint="Automatic. Do not type here.",
         guide="Filled in automatically. Shows OK, or the fields that are still missing on this row. Do not type in this column."),

    # ------------------------------------------------------------ 1. PO HEADER
    dict(key="po_number", grp="PO", name="PO Number", req=True, w=16, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Order"), fmt="Text (30)", ex="PO-1001",
         hint="Your PO number. Repeat on every line.",
         guide="Purchase order number from your current system. Repeat the SAME number on every line that belongs to this PO."),
    dict(key="po_description", grp="PO", name="PO Description", req=False, w=30, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Description"), fmt="Text (240)",
         ex="Annual IT hardware purchase",
         hint="What the PO is for.",
         guide="Short description of what the PO is for."),
    dict(key="doc_type", grp="PO", name="Document Type", req=True, w=17, dv="DocType",
         tgt=("PO_HEADERS_INTERFACE", "Document Type Code"), fmt="List", ex="STANDARD",
         hint="Always STANDARD.",
         guide="Always STANDARD for a normal purchase order. Blanket agreements are loaded with a different template."),
    dict(key="procurement_bu", grp="PO", name="Procurement BU", req=True, w=20, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Procurement BU"), fmt="Text (240)", ex="Example BU",
         hint="BU that owns and buys the PO.",
         guide="Business unit that owns / buys the PO. Must match the BU name in Oracle exactly."),
    dict(key="requisitioning_bu", grp="PO", name="Requisitioning BU", req=True, w=20, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Requisitioning BU"), fmt="Text (240)", ex="Example BU",
         hint="BU that requested it. Usually the same.",
         guide="Business unit that requested the goods or services. Usually the same as Procurement BU."),
    dict(key="legal_entity", grp="PO", name="Sold-to Legal Entity", req=True, w=24, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Sold-to Legal Entity"), fmt="Text (240)",
         ex="Example Legal Entity",
         hint="Legal entity the supplier invoices.",
         guide="Legal entity the supplier invoices. Must match the legal entity name in Oracle exactly."),
    dict(key="bill_to_bu", grp="PO", name="Bill-to BU", req=True, w=18, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Bill-to BU"), fmt="Text (240)", ex="Example BU",
         hint="BU that pays the invoice.",
         guide="Business unit that pays the supplier invoice. Usually the same as Procurement BU."),
    dict(key="buyer", grp="PO", name="Buyer", req=True, w=22, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Buyer"), fmt="Text (2000)", ex="Smith, John",
         hint="Buyer name as in Oracle: Last, First.",
         guide="Name of the buyer / procurement agent, exactly as the employee is named in Oracle, for example Smith, John."),
    dict(key="currency", grp="PO", name="Currency", req=False, w=11, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Currency Code"), fmt="Text (3)", ex="SAR",
         hint="Blank = ledger currency.",
         guide="Three letter currency code of the PO: SAR, USD, EUR, AED ... Leave blank for a PO in the ledger currency - Oracle defaults it from the supplier site or the ledger. Fill it in only for a foreign currency PO."),
    dict(key="conversion_rate", grp="PO", name="Conversion Rate", req=False, w=14, dv="num",
         tgt=("PO_HEADERS_INTERFACE", "Rate"), fmt="Number", ex="",
         hint="Foreign currency only.",
         guide="Only when the PO currency is different from the ledger currency and a fixed rate was used. Leave blank otherwise."),
    dict(key="rate_type", grp="PO", name="Rate Type", req="C", w=14, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Rate Type"), fmt="Text (30)", ex="",
         hint="Needed with a Conversion Rate.",
         guide="Conversion rate type, for example User, Corporate or Spot. Mandatory when Conversion Rate is filled in - use User for a rate you supply yourself. Must match a rate type set up in Oracle."),
    dict(key="default_tax_country", grp="PO", name="Default Taxation Country", req=False, w=18, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Default Taxation Country Code"), fmt="Text (2)", ex="",
         hint="2-letter code. Tax setups only.",
         guide="Two letter country code used to work out tax, for example SA or AE. Only needed when the client runs a real Fusion tax configuration. Leave blank if tax is not calculated on purchase orders."),
    dict(key="payment_terms", grp="PO", name="Payment Terms", req=True, w=16, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Payment Terms"), fmt="Text (50)", ex="Net 30",
         hint="For example Net 30, Immediate.",
         guide="Payment terms name as defined in Oracle, for example Immediate, Net 30, Net 60."),
    dict(key="bill_to_location", grp="PO", name="Bill-to Location", req=True, w=18, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Bill-to Location"), fmt="Text (60)", ex="Head Office",
         hint="Where the invoice is sent.",
         guide="Location where the supplier invoice is sent. Must match a location name in Oracle."),
    dict(key="hdr_note_supplier", grp="PO", name="Note to Supplier (PO)", req=False, w=24, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Note to Supplier"), fmt="Text (240)", ex="",
         hint="Message printed on the PO.",
         guide="Message printed on the PO for the supplier. Applies to the whole PO."),
    dict(key="hdr_note_receiver", grp="PO", name="Note to Receiver (PO)", req=False, w=24, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Note to Receiver"), fmt="Text (240)", ex="",
         hint="Internal note for the receiver.",
         guide="Internal message for the person receiving the goods. Applies to the whole PO."),

    # ------------------------------------------------------------ 2. SUPPLIER
    dict(key="supplier", grp="SUP", name="Supplier Name", req=True, w=26, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Supplier"), fmt="Text (360)", ex="ABC Trading Co.",
         hint="Exact name as registered in Oracle.",
         guide="Supplier name exactly as registered in Oracle. Spelling must match or the PO will fail."),
    dict(key="supplier_site", grp="SUP", name="Supplier Site", req=True, w=18, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Supplier Site"), fmt="Text (240)", ex="Head Office",
         hint="Purchasing site of the supplier.",
         guide="Supplier site / address code used for purchasing, for example Head Office, RIYADH."),
    dict(key="supplier_contact", grp="SUP", name="Supplier Contact", req=False, w=22, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Supplier Contact"), fmt="Text (240)", ex="",
         hint="Contact person at the supplier.",
         guide="Contact person at the supplier, as registered on the supplier site in Oracle."),
    dict(key="supplier_order", grp="SUP", name="Supplier Order Number", req=False, w=20, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Supplier Order"), fmt="Text (25)", ex="",
         hint="The supplier's own reference.",
         guide="The supplier's own reference or quotation number for this PO, if you track it."),
    dict(key="comm_method", grp="SUP", name="Communication Method", req=False, w=20, dv="CommMethod",
         tgt=("PO_HEADERS_INTERFACE", "Communication Method"), fmt="List", ex="NONE",
         hint="Use NONE for migrated POs.",
         guide="How the PO is sent to the supplier. Use NONE for migrated / historical POs so Oracle does not e-mail the supplier again."),
    dict(key="supplier_email", grp="SUP", name="Supplier Email", req="C", w=26, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "To Email"), fmt="Text (2000)", ex="",
         hint="Only when the method is EMAIL.",
         guide="Required only when Communication Method = EMAIL. Separate multiple addresses with a comma."),
    dict(key="fob", grp="SUP", name="FOB", req=False, w=12, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "FOB"), fmt="Text (25)", ex="",
         hint="Optional.",
         guide="Optional. Free on board terms, as set up in Oracle."),
    dict(key="freight_terms", grp="SUP", name="Freight Terms", req=False, w=14, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Freight Terms"), fmt="Text (25)", ex="",
         hint="Optional.",
         guide="Optional. Freight terms, as set up in Oracle."),
    dict(key="carrier", grp="SUP", name="Carrier", req=False, w=14, dv=None,
         tgt=("PO_HEADERS_INTERFACE", "Carrier"), fmt="Text (240)", ex="",
         hint="Optional.",
         guide="Optional. Shipping carrier name, as set up in Oracle."),

    # ------------------------------------------------------------ 3. LINE
    dict(key="line_number", grp="LIN", name="Line Number", req=True, w=11, dv="num",
         tgt=("PO_LINES_INTERFACE", "Line"), fmt="Number", ex=1,
         hint="1, 2, 3 ... restart per PO.",
         guide="Line number within the PO: 1, 2, 3 ... Start again at 1 for every new PO Number."),
    dict(key="line_type", grp="LIN", name="Line Type", req=True, w=20, dv="LineType",
         tgt=("PO_LINES_INTERFACE", "Line Type"), fmt="List", ex="Goods",
         hint="Goods, or a service type.",
         guide="Goods = quantity x price. Fixed Price Services / Rate Based Services = an amount, no quantity."),
    dict(key="item", grp="LIN", name="Item Number", req="C", w=18, dv=None,
         tgt=("PO_LINES_INTERFACE", "Item"), fmt="Text (300)", ex="",
         hint="Item code, if a catalogue item.",
         guide="Inventory item code, if the line is for a catalogue item. Leave blank for description based / expense lines."),
    dict(key="item_description", grp="LIN", name="Item Description", req=True, w=34, dv=None,
         tgt=("PO_LINES_INTERFACE", "Item Description"), fmt="Text (240)",
         ex="Laptop computer 14 inch",
         hint="What is being bought. Always fill this in.",
         guide="What is being bought. Always fill this in, also when an Item Number is given."),
    dict(key="category", grp="LIN", name="Category Name", req="C", w=20, dv=None,
         tgt=("PO_LINES_INTERFACE", "Category Name"), fmt="Text (2000)", ex="IT Hardware",
         hint="Needed when Item Number is blank.",
         guide="Purchasing category as defined in Oracle. Mandatory when Item Number is blank."),
    dict(key="uom", grp="LIN", name="UOM", req="C", w=12, dv=None,
         tgt=("PO_LINES_INTERFACE", "UOM"), fmt="Text (60)", ex="Each",
         hint="Each, Box, Hour ...",
         guide="Unit of measure name, for example Each, Box, Hour. Mandatory for Goods lines."),
    dict(key="quantity", grp="LIN", name="Quantity", req="C", w=12, dv="num",
         tgt=("PO_LINES_INTERFACE", "Quantity"), fmt="Number", ex=10,
         hint="Open qty still to receive.",
         guide="Quantity ordered. For open POs enter the quantity STILL TO BE RECEIVED. Mandatory for Goods lines. No thousands separator."),
    dict(key="unit_price", grp="LIN", name="Unit Price", req="C", w=12, dv="num",
         tgt=("PO_LINES_INTERFACE", "Price"), fmt="Number", ex=4500,
         hint="Price of one unit.",
         guide="Price of one unit, in the PO currency. Mandatory for Goods lines. No thousands separator."),
    dict(key="line_amount", grp="LIN", name="Line Amount", req="C", w=14, dv="num",
         tgt=("PO_LINES_INTERFACE", "Amount"), fmt="Number", ex="",
         hint="Services only. Goods: leave blank.",
         guide="Total value of the line. Mandatory for service lines. For Goods lines leave blank - it is calculated as Quantity x Unit Price."),
    dict(key="supplier_item", grp="LIN", name="Supplier Item Number", req=False, w=20, dv=None,
         tgt=("PO_LINES_INTERFACE", "Supplier Item"), fmt="Text (300)", ex="",
         hint="The supplier's part number.",
         guide="Optional. The supplier's own part number for this line."),
    dict(key="line_note_supplier", grp="LIN", name="Note to Supplier (Line)", req=False, w=22, dv=None,
         tgt=("PO_LINES_INTERFACE", "Note to Supplier "), fmt="Text (240)", ex="",
         hint="Note for this line only.",
         guide="Optional. Message for the supplier about this specific line."),
    dict(key="agreement", grp="LIN", name="Source Agreement", req=False, w=18, dv=None,
         tgt=("PO_LINES_INTERFACE", "Source Agreement"), fmt="Text (30)", ex="",
         hint="Contract or blanket number.",
         guide="Optional. Blanket / contract agreement number this line was released against."),
    dict(key="agreement_line", grp="LIN", name="Source Agreement Line", req=False, w=18, dv="num",
         tgt=("PO_LINES_INTERFACE", "Source Agreement Line"), fmt="Number", ex="",
         hint="Line on that agreement.",
         guide="Optional. Line number on the source agreement."),

    # ------------------------------------------------------------ 4. SHIPMENT
    dict(key="schedule_number", grp="SHP", name="Schedule No.", req=False, w=12, dv="num",
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Schedule"), fmt="Number", ex=1,
         hint="Leave 1 normally.",
         guide="Leave blank or 1 in normal cases. Use 2, 3 ... only when ONE line is delivered on several dates or to several locations - repeat the line on a new row for each schedule."),
    dict(key="ship_to_location", grp="SHP", name="Ship-to Location", req=True, w=18, dv=None,
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Ship-to Location"), fmt="Text (60)", ex="Head Office",
         hint="Where goods are delivered.",
         guide="Where the goods are delivered. Must match a location name in Oracle."),
    dict(key="ship_to_org", grp="SHP", name="Ship-to Organization", req="C", w=20, dv=None,
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Ship-to Organization"), fmt="Text (240)", ex="",
         hint="Needed when INVENTORY.",
         guide="Inventory organization code that receives the goods. Mandatory when Destination Type = INVENTORY."),
    dict(key="need_by_date", grp="SHP", name="Need-by Date", req=True, w=14, dv="date",
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Need-by Date"), fmt="Date YYYY-MM-DD", ex="2026-03-10",
         hint="YYYY-MM-DD.",
         guide="Date the goods or services are required. Use the format YYYY-MM-DD, for example 2026-03-10."),
    dict(key="promised_date", grp="SHP", name="Promised Date", req=False, w=14, dv="date",
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Promised Date"), fmt="Date YYYY-MM-DD", ex="2026-03-10",
         hint="YYYY-MM-DD, if known.",
         guide="Date the supplier promised to deliver. Leave blank if not tracked."),
    dict(key="destination_type", grp="SHP", name="Destination Type", req=True, w=16, dv="DestType",
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Destination Type Code"), fmt="List", ex="EXPENSE",
         hint="EXPENSE or INVENTORY.",
         guide="EXPENSE = consumed directly and charged to an account. INVENTORY = received into stock. DROP SHIP = delivered by the supplier to your customer."),
    dict(key="schedule_quantity", grp="SHP", name="Schedule Qty / Amount", req=False, w=20, dv="num",
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Quantity"), fmt="Number", ex="",
         hint="Only when splitting a line.",
         guide="Only needed when a line is split over several schedules. Leave blank for a single schedule - the line quantity or amount is used."),
    dict(key="receipt_required", grp="SHP", name="Receipt Required", req=False, w=15, dv="YN",
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Receipt Required"), fmt="List Y / N", ex="",
         hint="Y or N.",
         guide="Optional. Y when the goods must be received in Oracle before the invoice is paid."),
    dict(key="inspection_required", grp="SHP", name="Inspection Required", req=False, w=16, dv="YN",
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Inspection Required"), fmt="List Y / N", ex="",
         hint="Y or N.",
         guide="Optional. Y when the goods must be inspected after receipt."),
    dict(key="match_option", grp="SHP", name="Invoice Match Option", req=False, w=18, dv="MatchOption",
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Invoice Match Option"), fmt="List", ex="",
         hint="Order or Receipt.",
         guide="Optional. Order = invoice matched to the PO. Receipt = invoice matched to the receipt."),
    dict(key="tax_classification", grp="SHP", name="Tax Classification Code", req=False, w=20, dv=None,
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Tax Classification Code"), fmt="Text (50)", ex="",
         hint="Only if your tax setup needs it.",
         guide="Tax classification code for this shipment, as defined in Oracle tax setup. Only needed when the client runs a real Fusion tax configuration and tax is not derived automatically."),
    dict(key="sched_note_receiver", grp="SHP", name="Note to Receiver (Line)", req=False, w=22, dv=None,
         tgt=("PO_LINE_LOCATIONS_INTERFACE", "Note to receiver"), fmt="Text (240)", ex="",
         hint="Note for the receiver.",
         guide="Optional. Instruction for the person receiving this shipment."),

    # ------------------------------------------------------------ 5. ACCOUNT
    dict(key="distribution_number", grp="ACC", name="Distribution No.", req=False, w=14, dv="num",
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Distribution"), fmt="Number", ex=1,
         hint="Leave 1 normally.",
         guide="Leave blank or 1 in normal cases. Use 2, 3 ... only when ONE shipment is charged to several accounts - repeat the row for each account."),
    dict(key="distribution_quantity", grp="ACC", name="Distribution Qty / Amount", req=False, w=22, dv="num",
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Quantity"), fmt="Number", ex="",
         hint="Only when splitting a schedule.",
         guide="Only needed when a shipment is split over several accounts. Leave blank for a single distribution."),
    dict(key="deliver_to_location", grp="ACC", name="Deliver-to Location", req=False, w=18, dv=None,
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Deliver-to Location"), fmt="Text (60)", ex="Head Office",
         hint="Blank = same as Ship-to.",
         guide="Final internal location the goods are delivered to. Leave blank and the Ship-to Location on the same row is used, which is what almost every PO needs. Fill it in only when goods are shipped to one place and delivered on to another."),
    dict(key="requester", grp="ACC", name="Requester", req=True, w=24, dv=None,
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Requester"), fmt="Text (240)", ex="Smith, John",
         hint="Employee name: Last, First.",
         guide="Employee who requested the goods, exactly as named in Oracle, for example Smith, John."),
    dict(key="subinventory", grp="ACC", name="Subinventory", req=False, w=14, dv=None,
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Subinventory"), fmt="Text (10)", ex="",
         hint="Only for INVENTORY.",
         guide="Stock location inside the inventory organization. Only used when Destination Type = INVENTORY."),
    dict(key="charge_account", grp="ACC", name="Charge Account", req=False, w=30, dv=None,
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "CHARGE_ACCOUNT_SEGMENT*"), fmt="Text, segments split by -",
         ex="1000-101-520100-0000-000",
         hint="Blank = Oracle derives the account.",
         guide="Full GL account the cost is charged to, segments separated by a hyphen, for example 1000-101-520100-0000-000. Leave blank to let Oracle's account generator build the account from the item, category, destination and requester - that is the normal choice on a Fusion implementation. Fill it in only when the legacy account must be preserved exactly. If you do fill it in, every row must use the same number of segments in the same order as the client chart of accounts."),
    dict(key="project_number", grp="ACC", name="Project Number", req=False, w=16, dv=None,
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Project Number"), fmt="Text (25)", ex="",
         hint="Project purchases only.",
         guide="Optional. Only for project related purchases."),
    dict(key="task_number", grp="ACC", name="Task Number", req=False, w=14, dv=None,
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Task Number"), fmt="Text (100)", ex="",
         hint="Projects only.",
         guide="Optional. Project task. Required when a Project Number is given."),
    dict(key="expenditure_type", grp="ACC", name="Expenditure Type", req=False, w=18, dv=None,
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Expenditure Type"), fmt="Text (240)", ex="",
         hint="Projects only.",
         guide="Optional. Project expenditure type. Required when a Project Number is given."),
    dict(key="expenditure_date", grp="ACC", name="Expenditure Item Date", req=False, w=18, dv="date",
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Expenditure Item Date"), fmt="Date YYYY-MM-DD", ex="",
         hint="Projects only.",
         guide="Optional. Project expenditure date. Required when a Project Number is given."),
    dict(key="expenditure_org", grp="ACC", name="Expenditure Organization", req=False, w=22, dv=None,
         tgt=("PO_DISTRIBUTIONS_INTERFACE", "Expenditure Organization"), fmt="Text (240)", ex="",
         hint="Projects only.",
         guide="Optional. Organization owning the project expenditure. Required when a Project Number is given."),

    # ------------------------------------------------------------ 6. REFERENCE
    dict(key="original_qty", grp="REF", name="Original Ordered Qty", req=False, w=18, dv="num", tgt=None,
         fmt="Number", ex=12,
         hint="Information only.",
         guide="For information only. The full quantity originally ordered, before any receipts."),
    dict(key="received_qty", grp="REF", name="Qty Already Received", req=False, w=18, dv="num", tgt=None,
         fmt="Number", ex=2,
         hint="Information only. Proves the open balance.",
         guide="For information only. Quantity already received in the old system. Confirms that the Quantity column holds the open balance."),
    dict(key="invoiced_qty", grp="REF", name="Qty Already Invoiced", req=False, w=18, dv="num", tgt=None,
         fmt="Number", ex=2,
         hint="Information only.",
         guide="For information only. Quantity already invoiced in the old system."),
    dict(key="original_po_date", grp="REF", name="Original PO Date", req=False, w=16, dv="date", tgt=None,
         fmt="Date YYYY-MM-DD", ex="2026-01-15",
         hint="Information only.",
         guide="For information only. Date the PO was raised in the old system."),
    dict(key="remarks", grp="REF", name="Remarks", req=False, w=28, dv=None, tgt=None,
         fmt="Text", ex="",
         hint="Anything we should know about this line.",
         guide="For information only. Anything we should know about this line - open questions, exceptions, approvals."),
]


# ---------------------------------------------------------------- profiles
# Fields a Fusion implementation does NOT normally ask a client for, because
# Oracle derives them from setup, or because they are not part of an open-PO
# balance load. Reviewed with the client's Fusion consultant.
#
# They stay in the catalogue rather than being deleted: the next client may run
# real tax, multi-currency, or multi-account distributions. Build the wide sheet
# with  build_template.py --all  to expose them again.
ADVANCED = {
    # foreign currency - only when the PO is not in the ledger currency
    "conversion_rate":     "Only for a PO not in the ledger currency. Pairs with Rate Type.",
    "rate_type":           "Oracle uses the ledger's default conversion rate type unless a rate is supplied.",
    # tax - derived by the tax engine unless the client overrides it
    "default_tax_country": "Derived from the legal entity and tax setup.",
    "tax_classification":  "Derived by tax rules unless the client overrides tax per line.",
    # supplier-facing detail, held on the supplier site
    "supplier_order":      "The supplier's own reference. Informational, rarely known at migration.",
    "supplier_email":      "Taken from the supplier site contact. Migrated POs should not e-mail suppliers - leave Communication Method NONE.",
    "supplier_item":       "Derived from the item's supplier cross-reference when it is set up.",
    "line_note_supplier":  "Free text. Not part of an open-balance migration.",
    "sched_note_receiver": "Free text. Not part of an open-balance migration.",
    # schedule and distribution structure - derived by the converter
    "schedule_number":     "Derived. One schedule per line, auto-numbered only when a line really is split.",
    "schedule_quantity":   "Derived from the line quantity when the line has a single schedule.",
    "distribution_number": "Derived. One distribution per schedule, auto-numbered only when a schedule really is split.",
    "distribution_quantity": "Derived from the schedule quantity when the schedule has a single distribution.",
    # receiving and matching controls - all default from setup
    "receipt_required":    "Defaults from the item, supplier site and BU receiving options.",
    "inspection_required": "Defaults from the item and supplier site.",
    "match_option":        "Defaults from the supplier site's Invoice Match Option.",
    "subinventory":        "Only for an INVENTORY destination, and defaults from the item and organisation.",
    # legacy reference - never loaded into Oracle at all
    "original_qty":        "Legacy reference only - never loaded into Oracle.",
    "received_qty":        "Legacy reference only - never loaded into Oracle.",
    "invoiced_qty":        "Legacy reference only - never loaded into Oracle.",
    "original_po_date":    "Legacy reference only - never loaded into Oracle.",
    "remarks":             "Legacy reference only - never loaded into Oracle.",
}

# What makes one SCHEDULE of a line different from another: when it is wanted
# and where it goes. What makes one DISTRIBUTION different: what it is charged
# to. Shared by the converter's numbering and the sheet's duplicate formula, so
# the two can never disagree about what counts as a duplicate.
SCHEDULE_KEYS = ("need_by_date", "promised_date", "ship_to_location", "ship_to_org")
DIST_KEYS = ("charge_account", "deliver_to_location", "requester",
             "project", "task", "expenditure_type", "expenditure_org")

# Fields the converter fills in when the client sheet does not carry them.
DERIVED_WHEN_ABSENT = {
    "schedule_number":       "1, or auto-numbered when a line is genuinely split",
    "distribution_number":   "1, or auto-numbered when a schedule is genuinely split",
    "schedule_quantity":     "the line quantity",
    "distribution_quantity": "the schedule quantity",
    "deliver_to_location":   "the Ship-to Location on the same row",
}

CATALOGUE = list(COLUMNS)          # every field we know how to map, never filtered
PROFILE = "lean"


def attach_dff(columns, contexts):
    """Append a client's descriptive flexfield columns to the catalogue.

    Called by the builder and the converter from the same spec file, so the
    sheet the client fills and the mapping that reads it cannot drift apart.
    Value lists arrive per field, so they are registered in LISTS here rather
    than being declared up front like the fixed lists.
    """
    global DFF_CONTEXTS
    for c in columns:
        if c.get("_dff_values"):
            LISTS["DFF_" + c["key"]] = list(c["_dff_values"])
        CATALOGUE.append(c)
    DFF_CONTEXTS = dict(contexts)
    set_profile(PROFILE)
    return CATALOGUE


DFF_CONTEXTS = {}


def set_profile(name):
    """'lean' = the consultant-approved client sheet. 'full' = the catalogue."""
    global PROFILE, COLUMNS
    if name not in ("lean", "full"):
        raise ValueError("profile must be 'lean' or 'full'")
    PROFILE = name
    COLUMNS = [c for c in CATALOGUE if name == "full" or not c["adv"]]
    return COLUMNS


for _c in CATALOGUE:
    _c["adv"] = _c["key"] in ADVANCED
    _c["adv_why"] = ADVANCED.get(_c["key"], "")

set_profile("lean")


# ---------------------------------------------------------------- helpers
def col_index(key):
    """1-based Excel column index of a template column."""
    for i, c in enumerate(COLUMNS, start=1):
        if c["key"] == key:
            return i
    raise KeyError(key)


def by_key(key):
    """The column definition for `key`, from the whole catalogue.

    Metadata lookup must not depend on the active profile: a client mapping may
    legitimately point at a field the lean sheet does not show, and the
    converter still knows how to load it.
    """
    for c in CATALOGUE:
        if c["key"] == key:
            return c
    raise KeyError(key)


# Always-mandatory columns, grouped so the Row Check message stays short.
# Conditionally mandatory columns (quantity, price, amount, category, UOM,
# ship-to org, e-mail, charge account) are checked separately.
# Row Check groups the always-mandatory fields so the message stays short.
# Derived from req=True rather than hand-listed: a field made optional here can
# no longer be demanded by the sheet's formula, which is exactly the drift that
# made Currency and Deliver-to Location fail their own template once.
_GROUP_LABEL = {"PO": "PO Info", "SUP": "Supplier", "LIN": "Line",
                "SHP": "Shipment", "ACC": "Accounting",
                # a client DFF segment marked required is enforced like any
                # other mandatory field, in the sheet and in the converter
                "DFF": "Additional Fields"}


def mandatory_groups():
    out = {}
    for c in COLUMNS:
        if c["req"] is True and c["grp"] in _GROUP_LABEL:
            out.setdefault(_GROUP_LABEL[c["grp"]], []).append(c["key"])
    return out


class _MandatoryGroups(dict):
    """Recomputed on every read, so it follows the active profile."""

    def __iter__(self):
        return iter(mandatory_groups())

    def items(self):
        return mandatory_groups().items()

    def keys(self):
        return mandatory_groups().keys()

    def values(self):
        return mandatory_groups().values()

    def __getitem__(self, k):
        return mandatory_groups()[k]

    def __len__(self):
        return len(mandatory_groups())


MANDATORY_GROUPS = _MandatoryGroups()


# ---------------------------------------------------------------- synonyms
# Header captions seen in the wild that mean the same thing as one of our
# columns. Used by inspect_client_file.py to auto-map a client's own export.
# Matching ignores case, spaces and punctuation, so only list real variants.
SYNONYMS = {
    "po_number": ["po", "po no", "po number", "po num", "purchase order",
                  "purchase order number", "order", "order number",
                  "document number", "doc number", "legacy po", "legacy po number",
                  "po reference", "reference number"],
    "po_description": ["po description", "header description",
                       "purchase order description", "subject", "title", "purpose"],
    "doc_type": ["document type", "doc type", "po type", "order type",
                 "document type code", "type"],
    "procurement_bu": ["procurement bu", "procurement business unit", "bu",
                       "business unit", "purchasing bu", "buying bu", "operating unit"],
    "requisitioning_bu": ["requisitioning bu", "requisition bu", "req bu",
                          "requisitioning business unit", "requesting bu"],
    "legal_entity": ["legal entity", "sold to legal entity", "sold-to legal entity",
                     "le", "company", "entity"],
    "bill_to_bu": ["bill to bu", "bill-to bu", "billing bu", "bill to business unit"],
    "buyer": ["buyer", "buyer name", "purchasing agent", "agent", "procurement agent",
              "buyer full name"],
    "currency": ["currency", "currency code", "curr", "ccy", "po currency"],
    "conversion_rate": ["conversion rate", "exchange rate", "fx rate"],
    "rate_type": ["rate type", "conversion rate type", "exchange rate type",
                  "fx rate type"],
    "default_tax_country": ["default taxation country", "taxation country",
                            "tax country", "default taxation country code",
                            "country of taxation"],
    "payment_terms": ["payment terms", "payment term", "terms", "pay terms"],
    "bill_to_location": ["bill to location", "bill-to location", "billing location",
                         "invoice to location"],
    "hdr_note_supplier": ["note to supplier", "header note to supplier",
                          "note to vendor", "supplier note"],
    "hdr_note_receiver": ["note to receiver", "header note to receiver",
                          "receiver note"],

    "supplier": ["supplier", "supplier name", "vendor", "vendor name", "party name"],
    "supplier_site": ["supplier site", "vendor site", "site", "supplier site code",
                      "vendor site code", "supplier location"],
    "supplier_contact": ["supplier contact", "vendor contact", "contact",
                         "contact person", "contact name"],
    "supplier_order": ["supplier order", "supplier order number", "vendor order",
                       "quotation number", "quote number", "supplier reference"],
    "comm_method": ["communication method", "comm method", "notification method",
                    "send method", "po communication method"],
    "supplier_email": ["supplier email", "to email", "email", "email address",
                       "vendor email", "supplier e-mail"],
    "fob": ["fob", "free on board"],
    "freight_terms": ["freight terms", "freight term", "shipping terms"],
    "carrier": ["carrier", "shipping carrier", "freight carrier"],

    "line_number": ["line", "line number", "line no", "po line", "po line number",
                    "line num", "seq", "sequence", "item line"],
    "line_type": ["line type", "type of line", "po line type", "line type code"],
    "item": ["item", "item number", "item code", "part number", "part no",
             "material", "material code", "sku", "inventory item"],
    "item_description": ["item description", "description", "line description",
                         "material description", "product description", "particulars",
                         "goods description", "narration"],
    "category": ["category", "category name", "purchasing category", "commodity",
                 "commodity code", "item category", "category code"],
    "uom": ["uom", "unit of measure", "unit", "uom name", "uom code", "measure"],
    "quantity": ["quantity", "qty", "order quantity", "ordered quantity",
                 "open quantity", "remaining quantity", "balance quantity",
                 "pending quantity", "qty ordered", "balance qty", "open qty",
                 "remaining qty", "pending qty", "qty balance", "outstanding qty",
                 "outstanding quantity", "undelivered qty"],
    "unit_price": ["unit price", "price", "rate", "unit rate", "unit cost",
                   "price per unit", "item price"],
    "line_amount": ["amount", "line amount", "line total", "total amount",
                    "extended amount", "value", "line value", "net amount"],
    "supplier_item": ["supplier item", "supplier item number", "vendor item",
                      "supplier part number", "vendor part no"],
    "line_note_supplier": ["line note to supplier", "note to supplier line"],
    "agreement": ["source agreement", "agreement", "agreement number",
                  "blanket number", "contract number", "bpa number"],
    "agreement_line": ["source agreement line", "agreement line", "blanket line"],

    "schedule_number": ["schedule", "schedule no", "schedule number", "shipment",
                        "shipment number", "shipment no", "delivery number"],
    "ship_to_location": ["ship to location", "ship-to location", "ship to",
                         "shipping location", "delivery location", "destination",
                         "location", "site location"],
    "ship_to_org": ["ship to organization", "ship-to organization", "ship to org",
                    "inventory organization", "inventory org", "organization code",
                    "receiving org", "warehouse"],
    "need_by_date": ["need by date", "need-by date", "needby date", "required date",
                     "requested date", "delivery date", "due date", "expected date",
                     "need by"],
    "promised_date": ["promised date", "promise date", "supplier promised date",
                      "confirmed date", "eta"],
    "destination_type": ["destination type", "destination type code", "dest type",
                         "delivery type", "dest", "destination code",
                         "expense or inventory"],
    "schedule_quantity": ["schedule quantity", "shipment quantity", "schedule qty",
                          "shipment qty", "schedule amount"],
    "receipt_required": ["receipt required", "receipt reqd", "receiving required"],
    "inspection_required": ["inspection required", "inspection reqd"],
    "match_option": ["invoice match option", "match option", "matching option",
                     "match approval level"],
    "tax_classification": ["tax classification code", "tax classification",
                           "tax code", "tax class", "vat code"],
    "sched_note_receiver": ["line note to receiver", "note to receiver line",
                            "shipment note to receiver"],

    "distribution_number": ["distribution", "distribution no", "distribution number",
                            "dist no", "dist number", "accounting line"],
    "distribution_quantity": ["distribution quantity", "distribution qty",
                              "dist quantity", "distribution amount", "dist amount"],
    "deliver_to_location": ["deliver to location", "deliver-to location", "deliver to",
                            "delivery point", "final destination"],
    "requester": ["requester", "requestor", "deliver to person", "requested by",
                  "requester name", "employee", "end user", "beneficiary"],
    "subinventory": ["subinventory", "sub inventory", "sub-inventory", "store",
                     "stock location", "bin"],
    "charge_account": ["charge account", "account", "gl account", "account code",
                       "account combination", "code combination", "expense account",
                       "gl code", "accounting flexfield", "cost center account"],
    "project_number": ["project number", "project", "project code", "project no"],
    "task_number": ["task number", "task", "task code", "task no"],
    "expenditure_type": ["expenditure type", "exp type"],
    "expenditure_date": ["expenditure item date", "expenditure date", "exp item date"],
    "expenditure_org": ["expenditure organization", "expenditure org", "exp org"],

    "original_qty": ["original quantity", "original ordered qty", "ordered qty",
                     "original qty", "total quantity", "po quantity"],
    "received_qty": ["received quantity", "qty received", "received qty",
                     "delivered quantity", "grn quantity", "receipt quantity"],
    "invoiced_qty": ["invoiced quantity", "qty invoiced", "invoiced qty",
                     "billed quantity"],
    "original_po_date": ["po date", "order date", "creation date", "document date",
                         "original po date", "raised on"],
    "remarks": ["remarks", "comments", "notes", "note", "observation"],
}

# ---------------------------------------------------------------- value aliases
# Legacy values that must become the exact Oracle value. Compared after
# lower-casing and stripping surrounding whitespace.
VALUE_ALIASES = {
    "line_type": {
        "goods": "Goods", "good": "Goods", "item": "Goods", "items": "Goods",
        "material": "Goods", "materials": "Goods", "product": "Goods",
        "stock": "Goods", "inventory": "Goods", "g": "Goods",
        "services": "Fixed Price Services", "service": "Fixed Price Services",
        "fixed price": "Fixed Price Services",
        "fixed price services": "Fixed Price Services",
        "fixed price service": "Fixed Price Services",
        "amount based": "Fixed Price Services", "amount": "Fixed Price Services",
        "expense": "Fixed Price Services", "s": "Fixed Price Services",
        "rate based": "Rate Based Services",
        "rate based services": "Rate Based Services",
        "rate based service": "Rate Based Services",
        "time and material": "Rate Based Services",
    },
    "destination_type": {
        "expense": "EXPENSE", "exp": "EXPENSE", "consumption": "EXPENSE",
        "inventory": "INVENTORY", "inv": "INVENTORY", "stock": "INVENTORY",
        "store": "INVENTORY", "warehouse": "INVENTORY",
        "drop ship": "DROP SHIP", "dropship": "DROP SHIP",
        "drop shipment": "DROP SHIP",
    },
    "comm_method": {
        "none": "NONE", "no": "NONE", "n": "NONE", "not required": "NONE",
        "manual": "NONE", "na": "NONE", "n/a": "NONE",
        "email": "EMAIL", "e-mail": "EMAIL", "mail": "EMAIL",
        "print": "PRINT", "printed": "PRINT", "hard copy": "PRINT",
        "fax": "FAX",
    },
    "doc_type": {
        "standard": "STANDARD", "standard po": "STANDARD",
        "standard purchase order": "STANDARD", "spo": "STANDARD", "po": "STANDARD",
        "purchase order": "STANDARD",
    },
    "receipt_required": {"yes": "Y", "y": "Y", "true": "Y", "1": "Y", "required": "Y",
                         "no": "N", "n": "N", "false": "N", "0": "N",
                         "not required": "N"},
    "inspection_required": {"yes": "Y", "y": "Y", "true": "Y", "1": "Y",
                            "no": "N", "n": "N", "false": "N", "0": "N"},
    "match_option": {"order": "Order", "po": "Order", "purchase order": "Order",
                     "receipt": "Receipt", "receipts": "Receipt", "grn": "Receipt",
                     "2 way": "Order", "2-way": "Order",
                     "3 way": "Receipt", "3-way": "Receipt"},
}


def normalise_value(key, value):
    """Map a legacy value onto the exact Oracle value where we know the alias."""
    table = VALUE_ALIASES.get(key)
    if not table or value is None:
        return value
    hit = table.get(str(value).strip().lower())
    return hit if hit is not None else value
