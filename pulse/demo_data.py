"""Create a clearly fictional Business Central-shaped snapshot for offline demos."""

import gzip
import json
import random
from datetime import date, datetime, timedelta
from pathlib import Path


def create_snapshot(raw_dir: Path, days: int = 90) -> dict:
    """Write a deterministic synthetic snapshot accepted by ``pulse.model.build``."""
    raw_dir.mkdir(parents=True, exist_ok=True)
    today = date.today()
    now = datetime.now().replace(microsecond=0).isoformat()
    rng = random.Random(42)

    item_rows = [
        ("RM-1001", "Synthetic virgin resin", "VIRGIN", 5200, 2100),
        ("RM-1002", "Synthetic master batch", "MASTER BATCH", 13800, 180),
        ("RM-1003", "Synthetic packaging film", "PACKAGING FILM", 7100, 450),
        ("RM-1004", "Synthetic additive", "ADDITIVES", 9600, 95),
        ("RM-1005", "Synthetic regrind", "REGRIND", 2100, 900),
        ("WIP-2001", "Synthetic crate shell", "SEMI FINISHED", 4200, 240),
        ("WIP-2002", "Synthetic chair frame", "SEMI FINISHED", 6800, 105),
        ("FG-3001", "Synthetic storage crate", "FINISHED CRATES", 18500, 320),
        ("FG-3002", "Synthetic household chair", "FINNONCRAT", 27600, 86),
        ("FG-3003", "Synthetic garden stool", "FINNONCRAT", 14200, 210),
        ("FG-3004", "Synthetic cleaning brush set", "RESALE", 8300, 72),
    ]
    items = [{
        "number": number, "displayName": description, "type": "Inventory",
        "itemCategoryCode": "DEMO", "inventoryPostingGroupCode": group,
        "baseUnitOfMeasureCode": "PCS" if group in {"SEMI FINISHED", "FINISHED CRATES", "FINNONCRAT", "RESALE"} else "KG",
        "unitCost": cost, "blocked": False, "inventory": stock,
    } for number, description, group, cost, stock in item_rows]

    locations = [
        {"Code": "SHIP1RM", "Name": "Synthetic Site 1 Raw Materials", "Use_As_In_Transit": False, "Production_Floor": False},
        {"Code": "SHIP1PR", "Name": "Synthetic Site 1 Production", "Use_As_In_Transit": False, "Production_Floor": True},
        {"Code": "SHIP1FG", "Name": "Synthetic Site 1 Finished Goods", "Use_As_In_Transit": False, "Production_Floor": False},
        {"Code": "SHIP2RM", "Name": "Synthetic Site 2 Raw Materials", "Use_As_In_Transit": False, "Production_Floor": False},
        {"Code": "SHIP2PR", "Name": "Synthetic Site 2 Production", "Use_As_In_Transit": False, "Production_Floor": True},
        {"Code": "SHIP2FG", "Name": "Synthetic Site 2 Finished Goods", "Use_As_In_Transit": False, "Production_Floor": False},
        {"Code": "MIT", "Name": "Synthetic In Transit", "Use_As_In_Transit": True, "Production_Floor": False},
    ]

    stock_by_item = {
        "RM-1001": {"SHIP1RM": 1600, "SHIP2RM": 500},
        "RM-1002": {"SHIP1RM": 90, "SHIP2RM": 90},
        "RM-1003": {"SHIP1RM": 350, "SHIP2RM": 100},
        "RM-1004": {"SHIP1RM": 40, "SHIP2RM": 55},
        "RM-1005": {"SHIP1RM": 700, "SHIP2RM": 200},
        "WIP-2001": {"SHIP1PR": 160, "SHIP2PR": 80},
        "WIP-2002": {"SHIP1PR": 30, "SHIP2PR": 75},
        "FG-3001": {"SHIP1FG": 240, "SHIP2FG": 80},
        "FG-3002": {"SHIP1FG": 40, "SHIP2FG": 46},
        "FG-3003": {"SHIP1FG": 130, "SHIP2FG": 80},
        "FG-3004": {"SHIP1FG": 42, "SHIP2FG": 30},
    }
    # Keep the displayed item inventory equal to its synthetic ledger balance.
    for row in items:
        row["inventory"] = sum(stock_by_item[row["number"]].values())

    open_entries = []
    entry = 1
    for number, stock_map in stock_by_item.items():
        for location, total in stock_map.items():
            split = round(total * 0.35, 3)
            open_entries.append({"Entry_No": entry, "Item_No": number, "Location_Code": location,
                                 "Posting_Date": today.isoformat(), "Remaining_Quantity": split,
                                 "Entry_Type": "Positive Adjmt."})
            entry += 1
            open_entries.append({"Entry_No": entry, "Item_No": number, "Location_Code": location,
                                 "Posting_Date": today.isoformat(), "Remaining_Quantity": round(total - split, 3),
                                 "Entry_Type": "Purchase"})
            entry += 1
    movement_entries = []
    entry = 1000
    consumables = ["RM-1001", "RM-1002", "RM-1003", "RM-1004", "RM-1005", "WIP-2001", "WIP-2002"]
    finished = ["FG-3001", "FG-3002", "FG-3003", "FG-3004"]
    for offset in range(min(days, 90), 0, -7):
        posting = (today - timedelta(days=offset)).isoformat()
        for number in consumables:
            location = "SHIP1RM" if number.startswith("RM") else "SHIP1PR"
            qty = -rng.randint(35, 95) if number.startswith("RM") else -rng.randint(12, 35)
            movement_entries.append({"Entry_No": entry, "Entry_Type": "Consumption", "Item_No": number,
                                     "Location_Code": location, "Posting_Date": posting, "Quantity": qty,
                                     "Document_No": f"DEMO-CONS-{entry}"})
            entry += 1
        for number in finished:
            location = "SHIP1FG" if number in {"FG-3001", "FG-3002"} else "SHIP2FG"
            movement_entries.append({"Entry_No": entry, "Entry_Type": "Sale", "Item_No": number,
                                     "Location_Code": location, "Posting_Date": posting,
                                     "Quantity": -rng.randint(18, 64), "Document_No": f"DEMO-SALE-{entry}"})
            entry += 1
        for number in ["FG-3001", "FG-3002", "WIP-2001", "WIP-2002"]:
            location = "SHIP1PR"
            movement_entries.append({"Entry_No": entry, "Entry_Type": "Output", "Item_No": number,
                                     "Location_Code": location, "Posting_Date": posting,
                                     "Quantity": rng.randint(25, 110), "Document_No": f"DEMO-OUTPUT-{entry}"})
            entry += 1

    vendor_rows = [
        {"number": "DEMO-V001", "displayName": "Synthetic Resin Supplier"},
        {"number": "DEMO-V002", "displayName": "Synthetic Packaging Supplier"},
        {"number": "DEMO-V003", "displayName": "Synthetic General Supplier"},
    ]
    po_specs = [
        ("RM-1001", "DEMO-V001", "Synthetic Resin Supplier", 500, 6),
        ("RM-1002", "DEMO-V001", "Synthetic Resin Supplier", 160, 18),
        ("RM-1003", "DEMO-V002", "Synthetic Packaging Supplier", 240, -12),
        ("RM-1004", "DEMO-V003", "Synthetic General Supplier", 90, -118),
        ("RM-1005", "DEMO-V001", "Synthetic Resin Supplier", 300, 24),
        ("WIP-2001", "DEMO-V003", "Synthetic General Supplier", 60, -4),
        ("FG-3001", "DEMO-V002", "Synthetic Packaging Supplier", 70, 12),
        ("FG-3003", "DEMO-V003", "Synthetic General Supplier", 55, -35),
    ]
    purchases = []
    for line, (number, vendor_no, vendor_name, outstanding, days_from_today) in enumerate(po_specs, start=1):
        cost = next(row["unitCost"] for row in items if row["number"] == number)
        purchases.append({
            "Document_No": f"DEMO-PO-{1000 + line}", "Line_No": line * 10000,
            "Buy_from_Vendor_No": vendor_no, "Buy_from_Vendor_Name": vendor_name,
            "No": number, "Description": next(row["displayName"] for row in items if row["number"] == number),
            "Location_Code": "SHIP1RM" if number.startswith("RM") else "SHIP1FG",
            "Quantity": outstanding + 20, "Quantity_Received": 20, "Outstanding_Quantity": outstanding,
            "Expected_Receipt_Date": (today + timedelta(days=days_from_today)).isoformat(),
            "Status": "Open", "Unit_Cost_LCY": cost,
            "Outstanding_Amount_LCY": round(outstanding * cost, 2),
        })

    bom_headers = [
        {"No": "FG-3001", "Description": "Synthetic crate BOM", "Unit_of_Measure_Code": "PCS", "Status": "Certified", "ActiveVersionCode": "DEMO"},
        {"No": "FG-3002", "Description": "Synthetic chair BOM", "Unit_of_Measure_Code": "PCS", "Status": "Certified", "ActiveVersionCode": "DEMO"},
        {"No": "WIP-2001", "Description": "Synthetic crate-shell BOM", "Unit_of_Measure_Code": "PCS", "Status": "Certified", "ActiveVersionCode": "DEMO"},
        {"No": "WIP-2002", "Description": "Synthetic chair-frame BOM", "Unit_of_Measure_Code": "PCS", "Status": "Certified", "ActiveVersionCode": "DEMO"},
    ]
    bom_lines = []
    for parent, components in {
        "FG-3001": [("WIP-2001", 1.0), ("RM-1002", 0.15), ("RM-1003", 0.08)],
        "FG-3002": [("WIP-2002", 1.0), ("RM-1001", 0.75), ("RM-1004", 0.04)],
        "WIP-2001": [("RM-1001", 1.2), ("RM-1005", 0.3)],
        "WIP-2002": [("RM-1001", 0.85), ("RM-1002", 0.12)],
    }.items():
        for line, (component, qty) in enumerate(components, start=1):
            bom_lines.append({"Production_BOM_No": parent, "Version_Code": "DEMO", "Line_No": line * 10000,
                              "Type": "Item", "No": component,
                              "Description": next(row["displayName"] for row in items if row["number"] == component),
                              "Quantity_per": qty, "Unit_of_Measure_Code": "PCS", "Scrap_Percent": 2.0})

    tables = {
        "items": items, "locations": locations, "vendors": vendor_rows, "ile_open": open_entries,
        "ile_movements": movement_entries, "purchase_lines": purchases,
        "bom_headers": bom_headers, "bom_lines": bom_lines,
    }
    counts = {}
    for name, rows in tables.items():
        path = raw_dir / f"{name}.json.gz"
        with gzip.open(path, "wt", encoding="utf-8") as stream:
            json.dump(rows, stream)
        counts[name] = len(rows)
    log = {
        "started": now, "finished": now, "company": "Pulse Offline Demo",
        "environment": "SYNTHETIC / OFFLINE", "window_days": min(days, 90),
        "since": (today - timedelta(days=min(days, 90))).isoformat(),
        "data_mode": "SYNTHETIC", "synthetic": True,
        "tables": {name: {"rows": count, "seconds": 0} for name, count in counts.items()},
    }
    (raw_dir / "run_log.json").write_text(json.dumps(log, indent=2), encoding="utf-8")
    return log
