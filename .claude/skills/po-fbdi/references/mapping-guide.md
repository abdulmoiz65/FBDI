# Mapping guide

How to turn a client's own export into a mapping file the converter can read.
Read this after `inspect_client_file.py` has profiled the file.

## Schema

Save as `inbox/<client file name>.mapping.json`.

```json
{
  "_comment":   "Client, source system, extract date.",
  "_decisions": ["one line per ambiguity you resolved, and why"],

  "sheet":            "PO Dump 2026",
  "data_starts_row":  5,

  "columns":   { "po_number": "A", "quantity": "P", "unit_price": 17 },
  "constants": { "procurement_bu": "Example BU" },
  "overrides": { "7": { "line_amount": 180000 } },
  "charge_account_segments": 5
}
```

| Field | Meaning |
|---|---|
| `sheet` | Worksheet name. Omit for the first sheet. |
| `data_starts_row` | 1-based row of the **first data row**, not the header. |
| `columns` | field key → column letter (`"P"`) or 1-based index (`16`). |
| `constants` | field key → one value used on every row. Applied only where the row is empty. |
| `overrides` | source row number (as a string) → `{field key: value}`. Wins over everything. |
| `charge_account_segments` | how many segments the client chart of accounts has. Pinning it turns a mis-structured account into a hard error instead of a silent mis-map. |
| `_comment`, `_decisions` | ignored by the script, kept as the audit trail. **Always fill `_decisions` in.** |

Two field keys may point at the same column — common for
`ship_to_location` / `deliver_to_location` when the client tracks one location.

## Client-specific attributes

Descriptive flexfields are not in the catalogue below - they belong to one
client, not to the standard. Declare them in `dff/<client>.json` and pass
`--dff` to both scripts, or put `"dff": "../dff/<client>.json"` in the mapping
file. Full instructions in `dff-guide.md`.

## Field keys

The catalogue is every field the converter knows how to load. The **lean sheet**
is what a client actually receives — 45 columns, agreed with the Fusion
consultant. The other 22 are fields Oracle derives from setup, or that are not
part of an open-PO load; they stay in the catalogue and are read whenever a file
carries them, and `build_template.py --all` puts them back on a sheet.

Column letters below are the lean sheet; `—` means the field is not on it.

**CHECK**

| col | key | template field | mandatory | on lean sheet | notes |
|---|---|---|---|---|---|
| `A` | `check` | Row Check |  | yes |  |

**1. PURCHASE ORDER  (repeat on every line of the same PO)**

| col | key | template field | mandatory | on lean sheet | notes |
|---|---|---|---|---|---|
| `B` | `po_number` | PO Number | yes | yes | also matches: po, po no, po number, po num, purchase order |
| `C` | `po_description` | PO Description |  | yes | also matches: po description, header description, purchase order description, subject, title |
| `D` | `doc_type` | Document Type | yes | yes | also matches: document type, doc type, po type, order type, document type code |
| `E` | `procurement_bu` | Procurement BU | yes | yes | also matches: procurement bu, procurement business unit, bu, business unit, purchasing bu |
| `F` | `requisitioning_bu` | Requisitioning BU | yes | yes | also matches: requisitioning bu, requisition bu, req bu, requisitioning business unit, requesting bu |
| `G` | `legal_entity` | Sold-to Legal Entity | yes | yes | also matches: legal entity, sold to legal entity, sold-to legal entity, le, company |
| `H` | `bill_to_bu` | Bill-to BU | yes | yes | also matches: bill to bu, bill-to bu, billing bu, bill to business unit |
| `I` | `buyer` | Buyer | yes | yes | also matches: buyer, buyer name, purchasing agent, agent, procurement agent |
| `J` | `currency` | Currency |  | yes | also matches: currency, currency code, curr, ccy, po currency |
| `—` | `conversion_rate` | Conversion Rate |  | no | Only for a PO not in the ledger currency. Pairs with Rate Type. |
| `—` | `rate_type` | Rate Type | conditional | no | Oracle uses the ledger's default conversion rate type unless a rate is supplied. |
| `—` | `default_tax_country` | Default Taxation Country |  | no | Derived from the legal entity and tax setup. |
| `K` | `payment_terms` | Payment Terms | yes | yes | also matches: payment terms, payment term, terms, pay terms |
| `L` | `bill_to_location` | Bill-to Location | yes | yes | also matches: bill to location, bill-to location, billing location, invoice to location |
| `M` | `hdr_note_supplier` | Note to Supplier (PO) |  | yes | also matches: note to supplier, header note to supplier, note to vendor, supplier note |
| `N` | `hdr_note_receiver` | Note to Receiver (PO) |  | yes | also matches: note to receiver, header note to receiver, receiver note |

**2. SUPPLIER**

| col | key | template field | mandatory | on lean sheet | notes |
|---|---|---|---|---|---|
| `O` | `supplier` | Supplier Name | yes | yes | also matches: supplier, supplier name, vendor, vendor name, party name |
| `P` | `supplier_site` | Supplier Site | yes | yes | also matches: supplier site, vendor site, site, supplier site code, vendor site code |
| `Q` | `supplier_contact` | Supplier Contact |  | yes | also matches: supplier contact, vendor contact, contact, contact person, contact name |
| `—` | `supplier_order` | Supplier Order Number |  | no | The supplier's own reference. Informational, rarely known at migration. |
| `R` | `comm_method` | Communication Method |  | yes | also matches: communication method, comm method, notification method, send method, po communication method |
| `—` | `supplier_email` | Supplier Email | conditional | no | Taken from the supplier site contact. Migrated POs should not e-mail suppliers - leave Communication Method NONE. |
| `S` | `fob` | FOB |  | yes | also matches: fob, free on board |
| `T` | `freight_terms` | Freight Terms |  | yes | also matches: freight terms, freight term, shipping terms |
| `U` | `carrier` | Carrier |  | yes | also matches: carrier, shipping carrier, freight carrier |

**3. PO LINE**

| col | key | template field | mandatory | on lean sheet | notes |
|---|---|---|---|---|---|
| `V` | `line_number` | Line Number | yes | yes | also matches: line, line number, line no, po line, po line number |
| `W` | `line_type` | Line Type | yes | yes | also matches: line type, type of line, po line type, line type code |
| `X` | `item` | Item Number | conditional | yes | also matches: item, item number, item code, part number, part no |
| `Y` | `item_description` | Item Description | yes | yes | also matches: item description, description, line description, material description, product description |
| `Z` | `category` | Category Name | conditional | yes | also matches: category, category name, purchasing category, commodity, commodity code |
| `AA` | `uom` | UOM | conditional | yes | also matches: uom, unit of measure, unit, uom name, uom code |
| `AB` | `quantity` | Quantity | conditional | yes | also matches: quantity, qty, order quantity, ordered quantity, open quantity |
| `AC` | `unit_price` | Unit Price | conditional | yes | also matches: unit price, price, rate, unit rate, unit cost |
| `AD` | `line_amount` | Line Amount | conditional | yes | also matches: amount, line amount, line total, total amount, extended amount |
| `—` | `supplier_item` | Supplier Item Number |  | no | Derived from the item's supplier cross-reference when it is set up. |
| `—` | `line_note_supplier` | Note to Supplier (Line) |  | no | Free text. Not part of an open-balance migration. |
| `AE` | `agreement` | Source Agreement |  | yes | also matches: source agreement, agreement, agreement number, blanket number, contract number |
| `AF` | `agreement_line` | Source Agreement Line |  | yes | also matches: source agreement line, agreement line, blanket line |

**4. SHIPMENT / SCHEDULE**

| col | key | template field | mandatory | on lean sheet | notes |
|---|---|---|---|---|---|
| `—` | `schedule_number` | Schedule No. |  | no | Derived. One schedule per line, auto-numbered only when a line really is split. |
| `AG` | `ship_to_location` | Ship-to Location | yes | yes | also matches: ship to location, ship-to location, ship to, shipping location, delivery location |
| `AH` | `ship_to_org` | Ship-to Organization | conditional | yes | also matches: ship to organization, ship-to organization, ship to org, inventory organization, inventory org |
| `AI` | `need_by_date` | Need-by Date | yes | yes | also matches: need by date, need-by date, needby date, required date, requested date |
| `AJ` | `promised_date` | Promised Date |  | yes | also matches: promised date, promise date, supplier promised date, confirmed date, eta |
| `AK` | `destination_type` | Destination Type | yes | yes | also matches: destination type, destination type code, dest type, delivery type, dest |
| `—` | `schedule_quantity` | Schedule Qty / Amount |  | no | Derived from the line quantity when the line has a single schedule. |
| `—` | `receipt_required` | Receipt Required |  | no | Defaults from the item, supplier site and BU receiving options. |
| `—` | `inspection_required` | Inspection Required |  | no | Defaults from the item and supplier site. |
| `—` | `match_option` | Invoice Match Option |  | no | Defaults from the supplier site's Invoice Match Option. |
| `—` | `tax_classification` | Tax Classification Code |  | no | Derived by tax rules unless the client overrides tax per line. |
| `—` | `sched_note_receiver` | Note to Receiver (Line) |  | no | Free text. Not part of an open-balance migration. |

**5. CHARGE ACCOUNT / DISTRIBUTION**

| col | key | template field | mandatory | on lean sheet | notes |
|---|---|---|---|---|---|
| `—` | `distribution_number` | Distribution No. |  | no | Derived. One distribution per schedule, auto-numbered only when a schedule really is split. |
| `—` | `distribution_quantity` | Distribution Qty / Amount |  | no | Derived from the schedule quantity when the schedule has a single distribution. |
| `AL` | `deliver_to_location` | Deliver-to Location |  | yes | also matches: deliver to location, deliver-to location, deliver to, delivery point, final destination |
| `AM` | `requester` | Requester | yes | yes | also matches: requester, requestor, deliver to person, requested by, requester name |
| `—` | `subinventory` | Subinventory |  | no | Only for an INVENTORY destination, and defaults from the item and organisation. |
| `AN` | `charge_account` | Charge Account |  | yes | also matches: charge account, account, gl account, account code, account combination |
| `AO` | `project_number` | Project Number |  | yes | also matches: project number, project, project code, project no |
| `AP` | `task_number` | Task Number |  | yes | also matches: task number, task, task code, task no |
| `AQ` | `expenditure_type` | Expenditure Type |  | yes | also matches: expenditure type, exp type |
| `AR` | `expenditure_date` | Expenditure Item Date |  | yes | also matches: expenditure item date, expenditure date, exp item date |
| `AS` | `expenditure_org` | Expenditure Organization |  | yes | also matches: expenditure organization, expenditure org, exp org |

**6. LEGACY REFERENCE  (information only - not loaded into Oracle)**

| col | key | template field | mandatory | on lean sheet | notes |
|---|---|---|---|---|---|
| `—` | `original_qty` | Original Ordered Qty |  | no | Legacy reference only - never loaded into Oracle. |
| `—` | `received_qty` | Qty Already Received |  | no | Legacy reference only - never loaded into Oracle. |
| `—` | `invoiced_qty` | Qty Already Invoiced |  | no | Legacy reference only - never loaded into Oracle. |
| `—` | `original_po_date` | Original PO Date |  | no | Legacy reference only - never loaded into Oracle. |
| `—` | `remarks` | Remarks |  | no | Legacy reference only - never loaded into Oracle. |

## Judgment calls

### Open quantity vs ordered quantity

The single most damaging mistake in an open-PO load.

| Client column | Maps to | Why |
|---|---|---|
| `Balance Qty`, `Open Qty`, `Remaining Qty`, `Pending Qty` | `quantity` | what Oracle must still receive |
| `Qty Ordered`, `Original Qty`, `PO Qty` | `original_qty` | reference only, not loaded |
| `Qty Received`, `GRN Qty`, `Delivered Qty` | `received_qty` | reference only, not loaded |

If the file has only one quantity column, ask which it is. If it has an ordered
and a received column but no balance, ask whether to load `ordered − received`
rather than computing it silently.

### Conflicts

When two columns match one field the profiler assigns neither silently — the
stronger one keeps the field, the other is reported as *"also looks like X — you
decide"*. Resolve from the sample values, then write both into `columns`
explicitly.

### Constants vs overrides

- **Constant** — one value true for the entire load: `procurement_bu`,
  `legal_entity`, `bill_to_bu`, `bill_to_location`, `doc_type`, often
  `comm_method: "NONE"` for migrated POs.
- **Override** — a value true for one row only. Use for a figure buried in a
  remark, a one-off inventory org, a corrected date. Every override is printed
  in the conversion report, so nothing is hidden.

A constant that is only true for *most* rows is a bug. Split it into overrides.

### Values that normalise automatically

Handled by `VALUE_ALIASES` in `scripts/po_columns.py`:

| Field | Recognised input | Becomes |
|---|---|---|
| `line_type` | Material, Item, Product, Stock, Goods | `Goods` |
| | Service(s), Fixed Price, Amount Based, Expense | `Fixed Price Services` |
| | Rate Based, Time and Material | `Rate Based Services` |
| `destination_type` | Expense, Consumption | `EXPENSE` |
| | Inventory, Stock, Store, Warehouse | `INVENTORY` |
| | Drop Ship, Dropship | `DROP SHIP` |
| `comm_method` | None, Manual, N/A | `NONE` |
| | Email, E-mail | `EMAIL` |
| `doc_type` | Standard PO, SPO, Purchase Order | `STANDARD` |
| `receipt_required`, `inspection_required` | Yes/No/True/False/1/0 | `Y` / `N` |
| `match_option` | PO, 2-way / Receipt, GRN, 3-way | `Order` / `Receipt` |

Matching is case-insensitive. To add an alias, edit `VALUE_ALIASES` — never
pre-edit the client's file to suit the script.

## Dates and numbers

Accepted date inputs: real Excel dates, `YYYY-MM-DD`, `DD-MM-YYYY`,
`MM-DD-YYYY`, `DD-MON-YYYY`. Anything else is passed through and flagged as a
warning — check those rows, because `03-04-2026` is ambiguous and the converter
will not guess between March and April. If the client's export is ambiguous,
ask, then normalise with overrides or ask for a re-extract.

Numbers are stripped of thousands separators and spaces. A value that is still
not numeric becomes an error, never a zero.

## Charge account

One column holding the full combination, segments separated by `-` (also `.` or
`|`). Split into `CHARGE_ACCOUNT_SEGMENT1..30` in order.

**Confirm the segment order against the client's chart of accounts before the
first production run.** A file whose segments are in a different order loads
without error and posts to the wrong accounts.

## Worked example

`inbox/Najd - Open POs export.mapping.json` in this project is a complete
example against a Focus ERP extract: a title block above the headers, the
client's own captions, mixed-case line types, quantity taken from `Balance Qty`,
seven constants, and two overrides for a service amount and an inventory org.
