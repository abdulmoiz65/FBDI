# Oracle Fusion — Purchase Order data migration

Standard client template and automated conversion into Oracle Purchase Order
Import (FBDI).

**Drop a client file in `inbox/` and ask Claude to convert it.** The `po-fbdi`
skill profiles the file, proposes a column mapping, validates the data and
writes the FBDI package. Nothing is hand-built and the client's file is never
edited.

## Layout

```
.
├── .claude/
│   ├── settings.json
│   └── skills/po-fbdi/
│       ├── SKILL.md                     workflow Claude follows
│       ├── assets/fbdi_layout.json      Oracle column order, generated
│       ├── references/
│       │   ├── mapping-guide.md         mapping schema, field keys, judgment calls
│       │   ├── oracle-fbdi.md           the four interfaces, how to load in Fusion
│       │   └── validation-rules.md      every rule, and what it will not catch
│       └── scripts/
│           ├── po_columns.py            SINGLE SOURCE OF TRUTH
│           ├── build_template.py        renders the client template
│           ├── inspect_client_file.py   profiles any client file, proposes a mapping
│           ├── convert_to_fbdi.py       validates and writes the FBDI package
│           └── make_sample_filled.py    regenerates the demo data
│
├── templates/     what we send to clients (blank template + filled sample)
├── inbox/         client files come in here, with their mapping JSON
├── output/        one folder per conversion: 4 CSVs, the zip, the report
├── reference/     Oracle's own template and earlier client templates
└── docs/          FBDI Mapping Specification
```

## Manual use

```bash
S=.claude/skills/po-fbdi/scripts

# 1. profile whatever the client sent
python $S/inspect_client_file.py "inbox/Client export.xlsx"

# 2. write inbox/Client export.mapping.json   (skip if it is our template)

# 3. convert
python $S/convert_to_fbdi.py "inbox/Client export.xlsx" \
       --map "inbox/Client export.mapping.json" --batch 9100
```

Output lands in `output/<file name>/`: the four interface CSVs,
`PoImportOrders.zip` to upload, and `validation_report.txt`.

Rebuild the client template after changing columns:

```bash
python $S/build_template.py && python $S/make_sample_filled.py
```

## Changing the field set

Edit `scripts/po_columns.py` and nothing else. The template, the header
synonyms, the validator and the converter all read it. Then rebuild.

## Requirements

Python 3 with `openpyxl`. Excel is not needed for any script.

## Worked example

`inbox/Najd - Open POs export.xlsx` is a fabricated Focus ERP extract — a title
block above the headers, the client's own captions, mixed-case line types — with
`Najd - Open POs export.mapping.json` beside it showing constants, overrides and
the recorded decisions.
