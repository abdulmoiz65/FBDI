# Standard PO Template → Oracle FBDI Mapping Specification

Internal document. Describes how **Standard - Purchase Order Data Template.xlsx**
(the one client-facing sheet) maps onto the four sheets of Oracle's
**POPurchaseOrderImportTemplate.xlsm** (Purchase Order Import FBDI).

| | |
|---|---|
| Client template | `templates/Standard - Purchase Order Data Template.xlsx` — 2 sheets: *Overview & History*, *PO Data* |
| Oracle template | `reference/POPurchaseOrderImportTemplate.xlsm` — 4 interface sheets |
| Converter | `.claude/skills/po-fbdi/scripts/convert_to_fbdi.py` |
| Column definitions | `.claude/skills/po-fbdi/scripts/po_columns.py` — single source of truth for both scripts |
| Upload file | `PoImportOrders.zip` → *Load Interface File for Import* → **Import Orders** |

---

## 1. Why one sheet

Oracle's FBDI template needs the client to understand a four-level parent/child
structure (header → line → schedule → distribution) held together by four
artificial key columns. Clients cannot fill that. Our earlier Rewaa template
solved the readability problem but split the data over two sheets, so lines had
to be re-linked to headers by hand and header defaults had to be repeated
anyway.

The standard template collapses all four levels into **one row per PO line** and
generates every key column automatically:

```
client sheet                        Oracle FBDI
─────────────────────────────────   ────────────────────────────────────────────
                       ┌──────────► PO_HEADERS_INTERFACE          (1 per PO)
                       │
one row  ──────────────┼──────────► PO_LINES_INTERFACE            (1 per PO+line)
= one PO line          │
                       ├──────────► PO_LINE_LOCATIONS_INTERFACE   (1 per PO+line+schedule)
                       │
                       └──────────► PO_DISTRIBUTIONS_INTERFACE    (1 per PO+line+schedule+distribution)
```

Rules the client follows:

- **One row = one PO line.** The PO-level columns (blue and teal) carry the same
  value on every row of that PO. No merged cells, no blanks-mean-repeat.
- **Split delivery** — repeat the line on a new row and number *Schedule No.*
  1, 2, 3 …
- **Split account** — repeat the row and number *Distribution No.* 1, 2, 3 …
- Both split columns are left at 1 in normal cases, which is ~95 % of migrated POs.

## 2. Generated (never asked from the client)

| FBDI column | Value | Sheet |
|---|---|---|
| Interface Header Key | `H1`, `H2` … one per distinct **PO Number**, in first-seen order | Headers |
| Interface Line Key | `<header key>L<line>`, e.g. `H1L2` | Lines |
| Interface Line Location Key | `<line key>S<schedule>`, e.g. `H1L2S1` | Line Locations |
| Interface Distribution Key | `<location key>D<distribution>`, e.g. `H1L2S1D1` | Distributions |
| Action | `ORIGINAL` on headers, `ADD` on lines | Headers, Lines |
| Batch ID | `--batch` argument, same on every header row of one run | Headers |
| Schedule | *Schedule No.*, defaulted to 1 when blank | Line Locations |
| Distribution | *Distribution No.*, defaulted to 1 when blank | Distributions |
| Ship-to Location (header) | copied from the first row's line-level *Ship-to Location* | Headers |
| Approval Action | `--approval` argument. Blank ⇒ POs land as **Incomplete** and are approved in Oracle; `SUBMIT` routes them through approval; `BYPASS` brings them straight in as Approved (needs the privilege). | Headers |
| Import Source | `--source` argument, optional | Headers |

## 3. Derived values

| Field | Rule |
|---|---|
| Line **Amount** | Blank + Goods line ⇒ `Quantity × Unit Price`. Service lines take *Line Amount* as entered. |
| Line Quantity / Price / UOM | Dropped for service line types — Oracle rejects them there. |
| Schedule Quantity | *Schedule Quantity* if given, else the line quantity. Service lines get **Amount** instead of Quantity. |
| Distribution Quantity | *Distribution Quantity* if given, else the **schedule** value, else the line value. |
| **Charge Account** | Split on `-` (also `.` and `|`) into `CHARGE_ACCOUNT_SEGMENT1..30`, in order. Verify the segment order matches the client's chart of accounts before the first real run. |
| Dates | Any of `YYYY-MM-DD`, `DD-MM-YYYY`, `DD-MON-YYYY`, or a real Excel date → written as `YYYY/MM/DD`. |
| Numbers | Thousands separators and stray spaces stripped; a value that is still not numeric is reported as an error, never silently zeroed. |

## 4. Validation

The same rule set runs twice — once live in the client's workbook (column **A,
Row Check**, a formula: green *OK* or *Missing: …*), once in the converter,
which refuses to emit CSVs unless every row passes (`--force` overrides).

Checks: mandatory fields per group · Goods needs Quantity + Unit Price + UOM ·
service lines need Line Amount · Category Name required when Item Number is
blank · Ship-to Organization required for INVENTORY · Charge Account required
for EXPENSE · Supplier Email required when Communication Method = EMAIL ·
numbers parse · dates parse · **PO-level fields identical on every row of a PO** ·
no duplicate PO+line+schedule+distribution · **schedule quantities sum to the
line, distribution quantities sum to the schedule**.

The last two catch the failures FBDI is worst at reporting.

## 5. Field-by-field mapping

`M` = mandatory. `Cond.` = mandatory in certain cases (see the guidance row in
the template itself).

**CHECK**

| Col | Template field | M | FBDI sheet | FBDI column |
|---|---|---|---|---|
| A | Row Check | | *not loaded* | — |

**1. PURCHASE ORDER** — repeat on every line of the same PO

| Col | Template field | M | FBDI sheet | FBDI column |
|---|---|---|---|---|
| B | PO Number | Yes | PO_HEADERS | Order |
| C | PO Description | | PO_HEADERS | Description |
| D | Document Type | Yes | PO_HEADERS | Document Type Code |
| E | Procurement BU | Yes | PO_HEADERS | Procurement BU |
| F | Requisitioning BU | Yes | PO_HEADERS | Requisitioning BU |
| G | Sold-to Legal Entity | Yes | PO_HEADERS | Sold-to Legal Entity |
| H | Bill-to BU | Yes | PO_HEADERS | Bill-to BU |
| I | Buyer | Yes | PO_HEADERS | Buyer |
| J | Currency | Yes | PO_HEADERS | Currency Code |
| K | Conversion Rate | | PO_HEADERS | Rate |
| L | Payment Terms | Yes | PO_HEADERS | Payment Terms |
| M | Bill-to Location | Yes | PO_HEADERS | Bill-to Location |
| N | Note to Supplier (PO) | | PO_HEADERS | Note to Supplier |
| O | Note to Receiver (PO) | | PO_HEADERS | Note to Receiver |

**2. SUPPLIER**

| Col | Template field | M | FBDI sheet | FBDI column |
|---|---|---|---|---|
| P | Supplier Name | Yes | PO_HEADERS | Supplier |
| Q | Supplier Site | Yes | PO_HEADERS | Supplier Site |
| R | Supplier Contact | | PO_HEADERS | Supplier Contact |
| S | Supplier Order Number | | PO_HEADERS | Supplier Order |
| T | Communication Method | | PO_HEADERS | Communication Method |
| U | Supplier Email | Cond. | PO_HEADERS | To Email |
| V | FOB | | PO_HEADERS | FOB |
| W | Freight Terms | | PO_HEADERS | Freight Terms |
| X | Carrier | | PO_HEADERS | Carrier |

**3. PO LINE**

| Col | Template field | M | FBDI sheet | FBDI column |
|---|---|---|---|---|
| Y | Line Number | Yes | PO_LINES | Line |
| Z | Line Type | Yes | PO_LINES | Line Type |
| AA | Item Number | Cond. | PO_LINES | Item |
| AB | Item Description | Yes | PO_LINES | Item Description |
| AC | Category Name | Cond. | PO_LINES | Category Name |
| AD | UOM | Cond. | PO_LINES | UOM |
| AE | Quantity | Cond. | PO_LINES | Quantity |
| AF | Unit Price | Cond. | PO_LINES | Price |
| AG | Line Amount | Cond. | PO_LINES | Amount |
| AH | Supplier Item Number | | PO_LINES | Supplier Item |
| AI | Note to Supplier (Line) | | PO_LINES | Note to Supplier |
| AJ | Source Agreement | | PO_LINES | Source Agreement |
| AK | Source Agreement Line | | PO_LINES | Source Agreement Line |

**4. SHIPMENT / SCHEDULE**

| Col | Template field | M | FBDI sheet | FBDI column |
|---|---|---|---|---|
| AL | Schedule No. | | PO_LINE_LOCATIONS | Schedule |
| AM | Ship-to Location | Yes | PO_LINE_LOCATIONS | Ship-to Location |
| AN | Ship-to Organization | Cond. | PO_LINE_LOCATIONS | Ship-to Organization |
| AO | Need-by Date | Yes | PO_LINE_LOCATIONS | Need-by Date |
| AP | Promised Date | | PO_LINE_LOCATIONS | Promised Date |
| AQ | Destination Type | Yes | PO_LINE_LOCATIONS | Destination Type Code |
| AR | Schedule Quantity / Amount | | PO_LINE_LOCATIONS | Quantity / Amount |
| AS | Receipt Required | | PO_LINE_LOCATIONS | Receipt Required |
| AT | Inspection Required | | PO_LINE_LOCATIONS | Inspection Required |
| AU | Invoice Match Option | | PO_LINE_LOCATIONS | Invoice Match Option |
| AV | Note to Receiver (Line) | | PO_LINE_LOCATIONS | Note to receiver |

**5. CHARGE ACCOUNT / DISTRIBUTION**

| Col | Template field | M | FBDI sheet | FBDI column |
|---|---|---|---|---|
| AW | Distribution No. | | PO_DISTRIBUTIONS | Distribution |
| AX | Distribution Quantity / Amount | | PO_DISTRIBUTIONS | Quantity / Amount |
| AY | Deliver-to Location | Yes | PO_DISTRIBUTIONS | Deliver-to Location |
| AZ | Requester | Yes | PO_DISTRIBUTIONS | Requester |
| BA | Subinventory | | PO_DISTRIBUTIONS | Subinventory |
| BB | Charge Account | Cond. | PO_DISTRIBUTIONS | CHARGE_ACCOUNT_SEGMENT1..30 |
| BC | Project Number | | PO_DISTRIBUTIONS | Project Number |
| BD | Task Number | | PO_DISTRIBUTIONS | Task Number |
| BE | Expenditure Type | | PO_DISTRIBUTIONS | Expenditure Type |
| BF | Expenditure Item Date | | PO_DISTRIBUTIONS | Expenditure Item Date |
| BG | Expenditure Organization | | PO_DISTRIBUTIONS | Expenditure Organization |

**6. LEGACY REFERENCE** — information only, never loaded

| Col | Template field | M | FBDI sheet | FBDI column |
|---|---|---|---|---|
| BH | Original Ordered Qty | | *not loaded* | — |
| BI | Qty Already Received | | *not loaded* | — |
| BJ | Qty Already Invoiced | | *not loaded* | — |
| BK | Original PO Date | | *not loaded* | — |
| BL | Remarks | | *not loaded* | — |

Group 6 exists so an **open** PO can be reconciled: the client puts the
*remaining* quantity in column AE and tells us what was originally ordered and
already received. Without it there is no way to tell a short PO from a
partially-received one.

## 6. Oracle DFFs and flexfields

The Oracle template carries ~430 `ATTRIBUTE*`, `SH_ATTRIBUTE*`,
`GLOBAL_ATTRIBUTE*`, `ITEM_ATTRIBUTE*` and `PJC_*` columns. None of them are in
the client template. Where a client genuinely uses a DFF, add the column to
`.claude/skills/po-fbdi/scripts/po_columns.py` with its `tgt` and rebuild — both the template and the
converter pick it up from that one edit.

## 7. Running it

```bash
# rebuild the client template after editing scripts/po_columns.py
python .claude/skills/po-fbdi/scripts/build_template.py

# convert what the client sent back
python .claude/skills/po-fbdi/scripts/convert_to_fbdi.py "Client - PO Data.xlsx" --batch 9002 --xlsm

# regenerate the demo data set
python .claude/skills/po-fbdi/scripts/make_sample_filled.py
```

`--xlsm` additionally fills a copy of Oracle's own template, so the mapping can
be eyeballed sheet by sheet, or the CSVs regenerated with Oracle's own macro.
Output lands in `output/<file name>/`.

## 8. Before the first production run

1. Confirm the **charge account segment order** against the client's chart of accounts.
2. Decide the **Approval Action** (blank / `SUBMIT` / `BYPASS`) with the functional lead.
3. Agree whether *Quantity* holds the **open** or the **original** quantity — the
   template says open; the columns in group 6 are what proves it.
4. Run one PO through *Import Orders* in a test pod before sending the template
   to the client, so the value lists (line types, categories, UOMs, locations)
   are confirmed against the real setup.
