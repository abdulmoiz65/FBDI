---
name: po-fbdi
description: Convert a client-filled purchase order data file into Oracle Fusion Purchase Order Import FBDI files (PO_HEADERS / PO_LINES / PO_LINE_LOCATIONS / PO_DISTRIBUTIONS CSVs plus PoImportOrders.zip). Use whenever someone attaches, drops in inbox/, or points at a purchase order / open PO / PO data workbook and wants it loaded into Oracle, converted to FBDI, validated, or checked before upload. Also use to rebuild or amend the standard client PO template.
---

# Purchase Order → Oracle FBDI

Turn whatever the client sent into a loadable FBDI package, without hand-editing
Oracle's four-sheet template.

**Never hand-build the CSVs and never edit the client's file.** Every decision
lives in a mapping JSON so the run is reproducible and auditable.

## Paths

| | |
|---|---|
| Scripts | `.claude/skills/po-fbdi/scripts/` |
| Client files in | `inbox/` |
| Results out | `output/<file name>/` |
| Blank template to send clients | `templates/Standard - Purchase Order Data Template.xlsx` |
| Oracle originals | `reference/` |

Run scripts with `python .claude/skills/po-fbdi/scripts/<script>.py`.
They need only `openpyxl`.

## Workflow

### 1. Profile the file — always, even if it looks like our template

```bash
python .claude/skills/po-fbdi/scripts/inspect_client_file.py "inbox/<file>"
```

Reports the sheet, header row, data rows, and for every column: what it mapped
to, the confidence, and sample values. Three sections at the bottom need you:

- **CHECK THESE MATCHES** — confidence below 1.0
- **NOT MAPPED BUT HOLDS DATA** — including columns that lost a conflict
- **MANDATORY FIELDS WITH NO SOURCE COLUMN** — must become constants or be asked for

If it says `standard template: yes` and there is nothing in those three
sections, skip to step 3 and convert without `--map`.

### 2. Build the mapping

Write `inbox/<file>.mapping.json`. See `references/mapping-guide.md` for the
full schema; the short version:

```json
{
  "_decisions": ["why you resolved each ambiguity - keep this, it is the audit trail"],
  "sheet": "PO Dump 2026",
  "data_starts_row": 5,
  "columns":   {"po_number": "A", "quantity": "P"},
  "constants": {"procurement_bu": "Example BU"},
  "overrides": {"7": {"line_amount": 180000}},
  "charge_account_segments": 5
}
```

`columns` = read from that column. `constants` = same value on every row, for a
field the client has no column for. `overrides` = patch specific cells by source
row number, so the client's file stays untouched.

**The judgment calls you must make, not guess:**

- **Open quantity vs ordered quantity.** If the file has both an ordered and a
  balance/remaining column, `quantity` takes the **balance** — these are open
  POs and Oracle must receive only what is still outstanding. Put the ordered
  and received figures in `original_qty` / `received_qty` for reconciliation.
  Getting this backwards silently doubles the client's commitments. Say out loud
  which column you chose.
- **Conflicts.** Two columns competing for one field are reported, never
  auto-resolved. Decide from the sample values.
- **Missing mandatory fields.** BU, legal entity, bill-to, document type are
  usually one constant for the whole load — confirm the values with the user
  rather than inventing them. Never invent a supplier, buyer, requester,
  location, category or account.
- **Client attributes (DFFs) are declared, never edited in.** A client with
  descriptive flexfields gets `dff/<client>.json`; the same file builds their
  sheet and drives the converter, so a caption and an `ATTRIBUTE7` cannot drift
  apart. Never read a segment's prompt and assume the attribute number - take
  the Column Name from Manage Descriptive Flexfields. See
  `references/dff-guide.md`.
- **Two template profiles.** The client gets the **lean** sheet: 45 columns,
  reviewed with the Fusion consultant. 22 further fields Oracle derives from
  setup live in the catalogue and are read whenever a file carries them; run
  `build_template.py --all` for a sheet that exposes them (multi-currency, real
  tax, multi-account distributions). Never hand a client the wide sheet by
  default — it asks for things Oracle fills in better than they can.
- **Approval Action defaults to `BYPASS`** so migrated POs land Approved without
  running approval rules. This needs the *Bypass Purchase Order Approval*
  privilege on the account running the import — confirm it, or every row
  rejects. `--approval SUBMIT` routes them through approval instead; `NONE`
  leaves them Incomplete.
- **Schedules and distributions are derived, not asked for.** The client repeats
  a row and changes what differs; the converter numbers it. Check the **SPLIT
  LINES** section of the report against what the client intended.
- **Charge account segments.** Ask how many segments the client chart of
  accounts has and pin it with `charge_account_segments`. Without it a client
  whose COA differs from the example mis-maps silently. Confirm the segment
  ORDER too - the converter cannot check that.
- **Line type and destination values** normalise automatically (`Material` →
  `Goods`, `Expense` → `EXPENSE`). Check `references/mapping-guide.md` before
  adding a new alias.

### 3. Convert

```bash
python .claude/skills/po-fbdi/scripts/convert_to_fbdi.py "inbox/<file>" \
    --map "inbox/<file>.mapping.json" --batch 9100
```

Ask the user for `--batch` if they have a convention; otherwise pick an unused
number and say which you used. Useful flags: `--approval SUBMIT|BYPASS`,
`--source <name>`, `--xlsm` (also fill Oracle's own template for review),
`--force` (write CSVs despite errors — only when the user asks).

Nothing is written while validation fails. That is the point.

### 4. Triage the errors

Sort every error into one of three buckets and act differently:

| Bucket | Example | Action |
|---|---|---|
| Our mapping is wrong | everything in one column is missing | fix the mapping, re-run |
| A known correction | service amount buried in a remark | add an `override`, note it in `_decisions` |
| The client's data is genuinely incomplete | no charge account on 40 expense lines | **stop and list it for the client** — do not invent values |

Never use `--force` to make a red report go away.

### 5. Report back

Tell the user, in this order:

1. Row counts per interface file — headers / lines / schedules / distributions
2. Total value per currency, and the PO count, so they can reconcile against the source
3. Any override applied, and why
4. Anything the client still has to answer
5. Where the zip is, and what to do with it in Oracle (`references/oracle-fbdi.md`)

## Other tasks

**Rebuild the client template** (after changing columns):

```bash
python .claude/skills/po-fbdi/scripts/build_template.py
python .claude/skills/po-fbdi/scripts/make_sample_filled.py
```

**Add or change a field:** edit `scripts/po_columns.py` only. It is the single
source of truth — the template, the profiler's synonyms, the validator and the
converter all read it. Then rebuild the template.

## References

- `references/mapping-guide.md` — mapping schema, all field keys, aliases, worked example
- `references/oracle-fbdi.md` — the four interfaces, generated keys, how to load in Oracle
- `references/validation-rules.md` — every rule, and what the Oracle-side errors mean

## Rules

- Never edit the file the client sent. Corrections go in the mapping.
- Never invent a master-data value. A wrong supplier name fails the import; a
  plausible wrong one corrupts the load.
- Always state which column you took `quantity` from.
- Report what you skipped or assumed. A silent conversion is a failed one.
