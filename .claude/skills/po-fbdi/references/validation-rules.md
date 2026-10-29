# Validation rules

The same rule set runs twice: as a live formula in column A of the client
template, and in `convert_to_fbdi.py` before anything is written. Both read
`scripts/po_columns.py`, so they cannot drift apart.

No CSV is produced while any error stands. `--force` overrides that and should
only be used when the user explicitly asks.

## Errors — block the conversion

### Always mandatory

| Group | Fields |
|---|---|
| PO Info | PO Number, Document Type, Procurement BU, Requisitioning BU, Sold-to Legal Entity, Bill-to BU, Buyer, Currency, Payment Terms, Bill-to Location |
| Supplier | Supplier Name, Supplier Site |
| Line | Line Number, Line Type, Item Description |
| Shipment | Ship-to Location, Need-by Date, Destination Type |
| Accounting | Deliver-to Location, Requester |

### Conditionally mandatory

| Rule | When |
|---|---|
| Quantity, Unit Price, UOM | Line Type is `Goods` |
| Line Amount | any service line type |
| Category Name | Item Number is empty |
| Ship-to Organization | Destination Type is `INVENTORY` |
| Supplier Email | Communication Method is `EMAIL` |
| Rate Type | a Conversion Rate is given, and both columns are on the sheet |

A rule about a field Oracle derives only fires when that field is actually on
the sheet the client was given. The lean sheet has no Supplier Email, so
Communication Method = `EMAIL` is not an error there — Oracle takes the address
from the supplier site.

**Charge Account is no longer mandatory for EXPENSE.** On a Fusion
implementation the account generator builds the account from item, category,
destination and requester, so blank means "let Oracle derive it" — the normal
choice. Blank EXPENSE rows are reported once as a warning, so the consultant
confirms account generation is configured for the BU before the load.

### Value and type checks

- Line Type must be one of `Goods`, `Fixed Price Services`, `Rate Based Services`
- Destination Type must be one of `EXPENSE`, `INVENTORY`, `DROP SHIP`
- Quantity, Unit Price, Line Amount, schedule and distribution quantities,
  Conversion Rate and Line Number must parse as numbers
- A non-numeric value is an error, never silently zeroed

### Structural checks

These four are the reason the validator exists — Oracle reports them late and
badly, or not at all.

1. **PO-level fields identical on every row of a PO.** A different buyer on line
   4 of a five-line PO is reported with both row numbers and both values.
2. **No duplicate PO + line + schedule + distribution.** Flagged live in the
   sheet as well, in column A. A *gap* in the numbering (1, 2, 4) is not flagged
   — Oracle does not need line numbers to be contiguous.

   On the lean sheet the client never types a schedule or distribution number.
   Repeated rows of one line are read as a split, working out which kind from
   what actually differs: a different Need-by Date or Ship-to Location makes a
   new **schedule**, a different Charge Account or requester makes a new
   **distribution**. A repeated row that differs in *nothing* is still a
   duplicate error. Every line that was read as a split is listed in the report
   under **SPLIT LINES**, so a copy made by accident is visible rather than
   silently doubling a quantity.

   Quantities roll **up**: the figure on a row is that part's, a schedule is the
   sum of its distributions, and the line is the sum of its schedules. One
   client column carries a quantity on a `Goods` line and an amount on a service
   line.
3. **Charge accounts all have the same number of segments.** The client types
   one hyphen-delimited string that is fanned out into
   `CHARGE_ACCOUNT_SEGMENT1..30`. If one row has five segments and another six,
   values land in different Oracle segments per row and the load posts to the
   wrong accounts *without erroring*. Pin the count with
   `charge_account_segments` in the mapping file to make it exact; without it
   the counts still have to agree with each other.
4. **Quantities add up.** Schedule quantities must sum to the line total, and
   distribution quantities to their schedule total, within 0.005. When a line is
   split, every part must carry its own quantity — a blank part is an error, not
   an implied remainder.

## Warnings — reported, do not block

- A date that could not be parsed is passed through unchanged and flagged.
  Check these: `03-04-2026` is ambiguous and the converter does not guess.

## What the validator does not check

It cannot know your Fusion setup. These fail at import, not at conversion:

- whether the supplier, site, buyer, requester, location, category, UOM or
  payment terms actually exist in the target pod
- whether the account combination is valid, enabled, and passes cross-validation
- whether the segment order in the Charge Account column matches the COA
- whether the need-by date falls in an open period

Catching those needs a validation run in the pod. See
`references/oracle-fbdi.md`. The right sequence is: convert clean, load one PO
into a test pod, then load the batch.
