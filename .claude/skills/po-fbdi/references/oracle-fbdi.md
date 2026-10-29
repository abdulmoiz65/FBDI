# Oracle side: Purchase Order Import FBDI

What the generated package is, and what happens to it in Fusion.

## The four interfaces

Source template: `reference/POPurchaseOrderImportTemplate.xlsm` (568 columns
across four sheets). The column order we write is baked into
`assets/fbdi_layout.json`, so a conversion never depends on the .xlsm.

| CSV | Grain | Width |
|---|---|---|
| `PO_HEADERS_INTERFACE.csv` | one per PO | 165 |
| `PO_LINES_INTERFACE.csv` | one per PO + line | 153 |
| `PO_LINE_LOCATIONS_INTERFACE.csv` | one per PO + line + schedule | 125 |
| `PO_DISTRIBUTIONS_INTERFACE.csv` | one per PO + line + schedule + distribution | 125 |

All four are zipped as `PoImportOrders.zip` — the name Oracle's own macro uses.
CSVs carry **no header row**; position is everything, which is why the layout
asset is generated from Oracle's template rather than typed by hand.

## The keys that tie them together

Oracle links the four levels through interface keys the loader must supply.
We generate them; they are meaningful only inside one batch.

```
Interface Header Key          H2
Interface Line Key            H2L1          = header key + L + line number
Interface Line Location Key   H2L1S2        = line key   + S + schedule number
Interface Distribution Key    H2L1S2D1      = location key + D + distribution number
```

`Action` is `ORIGINAL` on headers and `ADD` on lines. `Batch ID` is the same on
every header row of one run and is how you find (and purge) your rows in the
interface tables.

## Loading it in Fusion

1. **Navigator → Tools → Scheduled Processes → Schedule New Process**
2. Run **Load Interface File for Import**
   - *Import Process*: **Import Orders**
   - *Data File*: upload `PoImportOrders.zip`
   - This chains two children: **Transfer File** and **Load File to Interface**.
     The zip lands in UCM under the `prc/po/import` account.
3. Check both children succeeded, then run **Import Orders** if it did not start
   automatically. Parameters that matter:
   - *Batch ID* / *Import Source* — restricts the run to your rows
   - *Default Buyer* — fallback when a header has no buyer
   - *Create or Update Purchase Orders*
   - *Approval Action* — also a column on `PO_HEADERS_INTERFACE` (column **E**,
     values `NONE,BYPASS,SUBMIT`), which is where the converter sets it. It
     defaults to **`BYPASS`**: an open-PO migration wants the orders to land
     already Approved without re-running approval rules or firing thousands of
     notifications. `SUBMIT` pushes every PO into the approval workflow;
     `NONE`/blank leaves them **Incomplete**. Override with `--approval`.
     BYPASS needs the *Bypass Purchase Order Approval* privilege on the run
     account — confirm this before the load or every row rejects.
   - *Communication Method* — set to none for migrated POs so suppliers are not
     re-notified
4. **Review the output and log.** The import produces the
   **Purchasing Interface Errors Report** listing rejected rows with the
   interface key, so an error maps straight back to a row in the client file.
5. Fix, then **purge before reloading**. Re-running without purging duplicates
   or blocks rows. Look for **Purge Purchasing Interface / Delete Purchase Order
   Interface Data** (name varies by release) and purge by Batch ID.

Exact process names and parameter labels shift slightly between releases —
confirm against the pod before a production run.

## Prerequisites — load these first

PO conversion sits near the end of the dependency chain. Every one of these
must already exist in Fusion or rows reject:

- Suppliers and supplier sites, with a purchasing site for the BU
- Buyers and requesters as **HCM persons** with the right roles
- Locations (bill-to, ship-to, deliver-to) and inventory organizations
- Items, if any line carries an item number
- Purchasing categories, units of measure, payment terms
- Chart of accounts and open GL periods covering the need-by dates

## Common rejections

| Error | Cause | Fix |
|---|---|---|
| Supplier / site not found | spelling, or no purchasing site for that BU | verify against the supplier master, not the client's spelling |
| Buyer / requester not found | not an HCM person, or the wrong name format | Fusion usually wants `Last, First` |
| Category / UOM invalid | value not in the client's Fusion setup | reconcile the value list before the load |
| Invalid account combination | wrong segment order, disabled combination, or cross-validation rule | check segment order against the COA |
| Need-by date in a closed period | dates before go-live | agree a date policy with the functional lead |
| Quantity or amount mismatch | schedules do not add up to the line, or distributions to the schedule | our validator catches this before upload |
| Duplicate PO number | already loaded in an earlier run | purge the batch, or the PO already exists |

## Reconciling after the load

Compare against the source extract, per currency:

- PO count, line count
- total ordered value
- the open balance — this is why `original_qty` / `received_qty` are collected
  even though they are never loaded

Report a difference; never quietly absorb one.
