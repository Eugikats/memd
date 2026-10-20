"""Extract the supply chain tables from Business Central into data/raw (staged as received).

Usage:  python -m pulse.extract [--days 90]
"""
import argparse
import gzip
import json
import time
from datetime import date, datetime, timedelta

from .bc_client import DATA_DIR, ROOT, BCClient

RAW = DATA_DIR / "raw"


def save(name, rows):
    RAW.mkdir(parents=True, exist_ok=True)
    with gzip.open(RAW / f"{name}.json.gz", "wt", encoding="utf-8") as f:
        json.dump(rows, f)
    return len(rows)


def run(days=90):
    bc = BCClient()
    since = (date.today() - timedelta(days=days)).isoformat()
    print(f"Company: {bc.company_name}  ({bc.env})")
    jobs = [
        ("items", lambda: bc.api("items", {"$select": "number,displayName,type,itemCategoryCode,"
                                           "inventoryPostingGroupCode,baseUnitOfMeasureCode,unitCost,blocked,inventory"})),
        ("locations", lambda: bc.odata("Location_Card_Excel", {
            "$select": "Code,Name,Production_Floor,Use_As_In_Transit,FinishedGoodsInventory,Bonded_Warehouse"})),
        ("vendors", lambda: bc.api("vendors", {"$select": "number,displayName"})),
        # Open entries carry the remaining quantity: their sum is stock on hand by item and location.
        ("ile_open", lambda: bc.odata("ItemLedgerEntries", {
            "$filter": "Open eq true",
            "$select": "Entry_No,Item_No,Location_Code,Posting_Date,Remaining_Quantity,Entry_Type"})),
        # Movements over the reporting window drive consumption and demand rates.
        ("ile_movements", lambda: bc.odata("ItemLedgerEntries", {
            "$filter": f"Posting_Date ge {since}",
            "$select": "Entry_No,Entry_Type,Item_No,Location_Code,Posting_Date,Quantity,Document_No"})),
        ("purchase_lines", lambda: bc.odata("Purchase_Lines_Excel", {
            "$filter": "Document_Type eq 'Order' and Type eq 'Item' and Outstanding_Quantity gt 0",
            "$select": "Document_No,Line_No,Buy_from_Vendor_No,Buy_from_Vendor_Name,No,Description,"
                       "Location_Code,Quantity,Quantity_Received,Outstanding_Quantity,Expected_Receipt_Date,"
                       "Status,Unit_Cost_LCY,Outstanding_Amount_LCY"})),
        ("bom_headers", lambda: bc.odata("Production_BOM_Excel", {"$select": "No,Description,Unit_of_Measure_Code,Status,ActiveVersionCode"})),
        ("bom_lines", lambda: bc.odata("Production_BOM_ExcelProdBOMLine", {
            "$select": "Production_BOM_No,Version_Code,Line_No,Type,No,Description,Quantity_per,"
                       "Unit_of_Measure_Code,Scrap_Percent"})),
    ]
    log = {"started": datetime.now().isoformat(timespec="seconds"), "company": bc.company_name,
           "environment": bc.env, "window_days": days, "since": since, "tables": {}}
    for name, fn in jobs:
        t = time.time()
        print(f"  {name} ...", flush=True)
        n = save(name, fn())
        log["tables"][name] = {"rows": n, "seconds": round(time.time() - t, 1)}
        print(f"  {name}: {n:,} rows in {time.time() - t:.0f}s", flush=True)
    log["finished"] = datetime.now().isoformat(timespec="seconds")
    (RAW / "run_log.json").write_text(json.dumps(log, indent=2))
    return log


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=90)
    run(ap.parse_args().days)
