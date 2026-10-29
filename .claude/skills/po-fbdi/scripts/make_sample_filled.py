# -*- coding: utf-8 -*-
"""
Fills the standard template with realistic sample data so the whole chain
(template -> converter -> FBDI csv) can be tested end to end.

Run:  python tools/make_sample_filled.py
Out:  Sample - Filled PO Data.xlsx   (in the project root)
"""
import os
import shutil
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import po_columns as P  # noqa: E402

SCRIPTS = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(SCRIPTS)
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(SKILL)))
TEMPLATES = os.path.join(ROOT, "templates")
SRC = os.path.join(TEMPLATES, "Standard - Purchase Order Data Template.xlsx")
DST = os.path.join(TEMPLATES, "Sample - Filled PO Data.xlsx")

HDR = dict(doc_type="STANDARD", procurement_bu="Example BU",
           requisitioning_bu="Example BU", legal_entity="Example Legal Entity",
           bill_to_bu="Example BU", buyer="Smith, John", currency="SAR",
           payment_terms="Net 30", bill_to_location="Head Office",
           supplier="ABC Trading Co.", supplier_site="Head Office",
           comm_method="NONE")

ROWS = [
    # PO-1001 : two goods lines, expense destination
    dict(HDR, po_number="PO-1001", po_description="IT hardware refresh",
         line_number=1, line_type="Goods", item_description="Laptop computer 14 inch",
         category="IT Hardware", uom="Each", quantity=10, unit_price=4500,
         ship_to_location="Head Office", need_by_date="2026-03-10",
         promised_date="2026-03-10", destination_type="EXPENSE",
         deliver_to_location="Head Office", requester="Smith, John",
         charge_account="1000-101-520100-0000-000",
         original_qty=12, received_qty=2),
    dict(HDR, po_number="PO-1001", po_description="IT hardware refresh",
         line_number=2, line_type="Goods", item_description="Docking station USB-C",
         category="IT Hardware", uom="Each", quantity=10, unit_price=650,
         ship_to_location="Head Office", need_by_date="2026-03-15",
         destination_type="EXPENSE", deliver_to_location="Head Office",
         requester="Smith, John",
         charge_account="1000-101-520100-0000-000"),

    # PO-1002 : one line delivered in two schedules
    dict(HDR, po_number="PO-1002", po_description="Office furniture",
         line_number=1, line_type="Goods", item_description="Ergonomic office chair",
         category="Furniture", uom="Each", quantity=40, unit_price=900,
         schedule_number=1, schedule_quantity=25,
         ship_to_location="Head Office", need_by_date="2026-04-01",
         destination_type="EXPENSE", deliver_to_location="Head Office",
         requester="Smith, John",
         charge_account="1000-102-520200-0000-000"),
    dict(HDR, po_number="PO-1002", po_description="Office furniture",
         line_number=1, line_type="Goods", item_description="Ergonomic office chair",
         category="Furniture", uom="Each", quantity=40, unit_price=900,
         schedule_number=2, schedule_quantity=15,
         ship_to_location="Branch Office", need_by_date="2026-05-01",
         destination_type="EXPENSE", deliver_to_location="Branch Office",
         requester="Smith, John",
         charge_account="1000-102-520200-0000-000"),

    # PO-1003 : fixed price service line, split over two charge accounts
    dict(HDR, po_number="PO-1003", po_description="Annual audit fees",
         line_number=1, line_type="Fixed Price Services",
         item_description="External audit - FY26", category="Professional Services",
         line_amount=180000, distribution_number=1, distribution_quantity=120000,
         ship_to_location="Head Office", need_by_date="2026-06-30",
         destination_type="EXPENSE", deliver_to_location="Head Office",
         requester="Smith, John",
         charge_account="1000-101-540300-0000-000"),
    dict(HDR, po_number="PO-1003", po_description="Annual audit fees",
         line_number=1, line_type="Fixed Price Services",
         item_description="External audit - FY26", category="Professional Services",
         line_amount=180000, distribution_number=2, distribution_quantity=60000,
         ship_to_location="Head Office", need_by_date="2026-06-30",
         destination_type="EXPENSE", deliver_to_location="Head Office",
         requester="Smith, John",
         charge_account="1000-102-540300-0000-000"),

    # PO-1004 : inventory destination with an item number
    dict(HDR, po_number="PO-1004", po_description="Stock replenishment",
         line_number=1, line_type="Goods", item="INV-110",
         item_description="A4 paper ream 80gsm", uom="Box", quantity=200,
         unit_price=18, ship_to_location="Head Office", ship_to_org="EXAMPLE-ORG",
         need_by_date="2026-03-20", destination_type="INVENTORY",
         deliver_to_location="Head Office", subinventory="STORES",
         requester="Smith, John"),
]


def main():
    shutil.copyfile(SRC, DST)
    wb = openpyxl.load_workbook(DST)
    ws = wb[P.DATA_SHEET]
    for i, rec in enumerate(ROWS):
        r = P.ROW_DATA + i
        for key, val in rec.items():
            # The sample is written against whichever profile the template was
            # built with, so a field the lean sheet does not carry is skipped.
            try:
                col = P.col_index(key)
            except KeyError:
                continue
            ws.cell(r, col, val)
    wb.save(DST)
    print("written: %s  (%d rows)" % (DST, len(ROWS)))


if __name__ == "__main__":
    main()
