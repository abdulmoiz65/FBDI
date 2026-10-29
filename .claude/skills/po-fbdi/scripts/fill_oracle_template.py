#!/usr/bin/env python
"""Fill Oracle's own POPurchaseOrderImportTemplate.xlsm from generated FBDI CSVs.

Unlike a plain openpyxl save, this patches the worksheet XML *inside* a copy of
the original .xlsm and copies every other zip part byte-for-byte. That keeps the
macro-attached "Generate CSV File" button on the Instructions sheet, which
openpyxl silently drops when it rewrites xl/drawings/drawing1.xml.

The CSVs are written full width (one field per template column) so values are
placed positionally - CSV field i goes to column i. No caption matching, and the
workbook therefore mirrors the zip exactly.

    python fill_oracle_template.py "output/<name>" [output-file-name.xlsm]
"""
import csv, os, re, shutil, sys, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__))))))
TEMPLATE = os.path.join(ROOT, "reference", "POPurchaseOrderImportTemplate.xlsm")
DATA_ROW = 5
# sheet name -> part inside the xlsm (from xl/workbook.xml + its rels)
SHEETS = {
    "PO_HEADERS_INTERFACE":        "xl/worksheets/sheet2.xml",
    "PO_LINES_INTERFACE":          "xl/worksheets/sheet3.xml",
    "PO_LINE_LOCATIONS_INTERFACE": "xl/worksheets/sheet4.xml",
    "PO_DISTRIBUTIONS_INTERFACE":  "xl/worksheets/sheet5.xml",
}
NUMERIC = re.compile(r"-?(0|[1-9]\d*)(\.\d+)?$")     # excludes "0000", "007" etc.


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def esc(v):
    return (v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def cell_xml(ref, value):
    if NUMERIC.match(value):
        return '<c r="%s"><v>%s</v></c>' % (ref, value)
    # inline string: no sharedStrings.xml surgery, so that part stays untouched
    return '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, esc(value))


def build_rows(records):
    out = []
    for i, rec in enumerate(records):
        r = DATA_ROW + i
        cells = "".join(cell_xml("%s%d" % (col_letter(c + 1), r), v)
                        for c, v in enumerate(rec) if v != "")
        out.append('<row r="%d" spans="1:%d">%s</row>' % (r, len(rec), cells))
    return "".join(out)


def patch_sheet(xml, records):
    """Keep everything above DATA_ROW verbatim, replace the sample rows below."""
    m = re.search(r"(<sheetData>)(.*?)(</sheetData>)", xml, re.S)
    if not m:
        raise SystemExit("no <sheetData> found")
    kept = "".join(row for row in re.findall(r"<row .*?</row>|<row [^>]*/>", m.group(2), re.S)
                   if int(re.match(r'<row r="(\d+)"', row).group(1)) < DATA_ROW)
    return xml[:m.start()] + "<sheetData>" + kept + build_rows(records) + "</sheetData>" + xml[m.end():]


def fill(outdir, name="POPurchaseOrderImportTemplate_filled.xlsm", quiet=True):
    dst = os.path.join(outdir, name)
    if os.path.exists(dst):
        try:
            os.remove(dst)
        except PermissionError:
            raise SystemExit("%s is open in Excel - close it and re-run." % dst)

    new = {}
    for sheet, part in SHEETS.items():
        path = os.path.join(outdir, sheet + ".csv")
        if not os.path.exists(path):
            raise SystemExit("missing %s - run convert_to_fbdi.py first" % path)
        with open(path, encoding="utf-8-sig", newline="") as fh:
            new[part] = list(csv.reader(fh))

    src = zipfile.ZipFile(TEMPLATE)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info in src.infolist():                       # order and names preserved
            data = src.read(info.filename)
            if info.filename in new:
                data = patch_sheet(data.decode("utf-8"), new[info.filename]).encode("utf-8")
            z.writestr(info, data)
    src.close()

    if not quiet:
        print("filled  : %s" % dst)
        for sheet, part in SHEETS.items():
            print("  %-30s %d row(s) at row %d" % (sheet, len(new[part]), DATA_ROW))
        print("macros and the Generate CSV File button are preserved (parts copied verbatim)")
    return dst


if __name__ == "__main__":
    fill(sys.argv[1] if len(sys.argv) > 1 else ".",
         sys.argv[2] if len(sys.argv) > 2 else "POPurchaseOrderImportTemplate_filled.xlsm",
         quiet=False)
