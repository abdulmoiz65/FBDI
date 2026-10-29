# Client attributes (DFFs) in the standard template

The 45 core columns are the same for every client. Descriptive flexfields are
the opposite — they exist only because one client configured them, and the
attribute column a segment lands in is decided by that client's own setup.

So DFFs are never edited into the template. Each client gets a small
declaration file, and the template **and** the converter are built from it.

```
dff/<client>.json  ──┬──►  build_template.py --dff …   the sheet the client fills
                     └──►  convert_to_fbdi.py --dff …  the FBDI attribute columns
```

One file, two consumers. They cannot drift apart, which is the whole point: a
caption on a sheet and an `ATTRIBUTE7` in a CSV have to mean the same thing
months apart.

## What Oracle gives you

Per interface, confirmed against `POPurchaseOrderImportTemplate.xlsm`:

| Level | Interface | Text | Number | Date | Context |
|---|---|---|---|---|---|
| header | `PO_HEADERS_INTERFACE` | `ATTRIBUTE1-20` | `ATTRIBUTE_NUMBER1-10` | `ATTRIBUTE_DATE1-10` | `ATTRIBUTE_CATEGORY` |
| line | `PO_LINES_INTERFACE` | `ATTRIBUTE1-20` | `ATTRIBUTE_NUMBER1-10` | `ATTRIBUTE_DATE1-10` | `ATTRIBUTE_CATEGORY` |
| schedule | `PO_LINE_LOCATIONS_INTERFACE` | `ATTRIBUTE1-20` | `ATTRIBUTE_NUMBER1-10` | `ATTRIBUTE_DATE1-10` | `ATTRIBUTE_CATEGORY` |
| distribution | `PO_DISTRIBUTIONS_INTERFACE` | `ATTRIBUTE1-20` | `ATTRIBUTE_NUMBER1-10` | `ATTRIBUTE_DATE1-10` | `ATTRIBUTE_CATEGORY` |

Also present, and **not** the ordinary PO DFF — do not reach for these unless
the client's setup actually points there:

- `GLOBAL_ATTRIBUTE*` — the country/global flexfield. 6 on headers, 1 on
  schedules and distributions.
- `SH_ATTRIBUTE*` on headers and lines, `ITEM_ATTRIBUTE*` on lines — separate
  flexfields carried on the same sheet. Confirm the flexfield code before using
  them; the tool allows them but cannot tell you they are right.
- `PJC_*` on distributions — Project Costing, driven by the project fields, not
  a DFF you fill in by hand.
- `ATTRIBUTE_TIMESTAMP*` — **not supported**. Load the date part into an
  `ATTRIBUTE_DATEn` column and raise it with the team.

## Getting the values (do not guess these)

1. Setup and Maintenance → **Manage Descriptive Flexfields**
2. Search the flexfield code for the level:

   | Level | Flexfield code |
   |---|---|
   | header | `PO_HEADERS` |
   | line | `PO_LINES` |
   | schedule | `PO_LINE_LOCATIONS` |
   | distribution | `PO_DISTRIBUTIONS` |

3. **Edit → Manage Contexts →** open the context the client uses → its segments.
4. Each segment shows two things you need, and they are unrelated:
   - **Prompt** → becomes `name`, the caption the client sees
   - **Column Name** → becomes `attribute`

   A segment prompted "Project Code" can sit in `ATTRIBUTE12`. Reading the
   prompt and assuming `ATTRIBUTE1` is how data ends up in the wrong field.
5. Note the **Context Code** once per level → `context`.

## Writing the spec

Copy `dff/example.json`, which carries the full instructions inline, and delete
what the client does not have.

```json
{
  "client": "Example Trading Co.",
  "confirmed_from": "A Rahman (functional lead), UAT2 pod, 2026-09-08",
  "fields": [
    {
      "key": "budget_owner",
      "level": "header",
      "name": "Budget Owner",
      "context": "MIGRATION",
      "attribute": "ATTRIBUTE1",
      "type": "text",
      "required": true,
      "values": ["Finance", "Operations", "IT"],
      "hint": "Department that owns the budget.",
      "guide": "Longer help text for the mapping guide."
    }
  ]
}
```

| Field | Notes |
|---|---|
| `key` | internal, `lower_snake_case`, unique, must not collide with a core field |
| `level` | `header` \| `line` \| `schedule` \| `distribution` |
| `name` | the segment **prompt** — what the client reads |
| `context` | the DFF **context code**. One per level, at most |
| `attribute` | the segment's **column name** |
| `type` | `text` \| `number` \| `date`, must match the column family |
| `required` | `true` puts it in Row Check and the converter's mandatory list |
| `values` | optional; becomes a dropdown. A long list goes to a hidden Lists sheet automatically |
| `hint`, `guide` | optional; sensible defaults are generated |

## What is refused before a template is built

Every one of these is a silent-corruption bug otherwise, because Oracle happily
accepts a well-formed row carrying the right value in the wrong column.

- an attribute column that does not exist on that interface
- a type that does not match the column family — a `number` segment cannot be
  pointed at `ATTRIBUTE1`
- two segments at the same level claiming the same attribute column
- two different contexts at the same level — a row carries only one
  `ATTRIBUTE_CATEGORY` per interface
- a key or caption that collides with a core template field, which would make
  the converter read the wrong column
- a missing `confirmed_from`, because a DFF pointed at the wrong attribute loads
  silently and is found months later

All problems are reported at once, so a spec is fixed in one pass.

## Running it

```bash
# the client's sheet - 45 core columns plus their own, named after them
python .claude/skills/po-fbdi/scripts/build_template.py --dff dff/acme.json

# converting what they send back - the SAME spec file
python .claude/skills/po-fbdi/scripts/convert_to_fbdi.py "inbox/Acme.xlsx" \
       --dff dff/acme.json --batch 9100
```

For a client on their own layout, put the path in the mapping file instead and
the converter picks it up:

```json
{ "sheet": "Sheet1", "columns": { }, "dff": "../dff/acme.json" }
```

## Two behaviours worth knowing

**The context is only written when the row carries a value.** Stamping
`ATTRIBUTE_CATEGORY` on a row with no attribute values makes Oracle validate
that context's required segments against nothing, and the row rejects. So a
file where nobody filled in the DFF columns loads exactly as it would have
without the spec.

**Long value lists move to a hidden Lists sheet.** Excel's inline dropdown gives
up past about 255 characters of joined values. Anything longer is written to a
hidden `Lists` sheet and the dropdown points at that range instead.

## Adding a client

1. Confirm the segments with the functional lead, from the client's instance.
2. `cp dff/example.json dff/<client>.json`, fill it in, set `confirmed_from`.
3. `build_template.py --dff dff/<client>.json` — fix anything it refuses.
4. Send the client the workbook it names.
5. Convert with the same `--dff`, and check the **client dff** line of the
   report says what you expect before uploading.
