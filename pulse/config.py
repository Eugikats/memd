"""Business configuration for the supply chain model. Reviewed with the Supply Chain team during the proof period."""

# Item class from the Business Central Inventory Posting Group.
CLASS_BY_POSTING_GROUP = {
    # Raw material (incl. packaging and auxiliary material consumed in production)
    "VIRGIN": "RM", "MASTER BATCH": "RM", "ADDITIVES": "RM", "REGRIND": "RM", "RECYCLE": "RM",
    "DEFECTS": "RM", "CRATE SHELLS": "RM", "PACKAGING FILM": "RM", "PACKAGING PAPER": "RM",
    "AUXILIARY PEN": "RM", "AUXILIARY OTHERS": "RM", "AUXILIARY TOOTHBRUSH": "RM",
    # Work in progress
    "SEMI FINISHED": "WIP",
    # Finished goods
    "FINNONCRAT": "FG", "FINISHED CRATES": "FG", "RESALE": "FG",
}
# Anything not listed (spares, factory consumables, stationery, fuel, PPE) is "Other" and kept out of cover views.
CLASS_LABELS = {"RM": "Raw Material", "WIP": "Work in Progress", "FG": "Finished Goods", "OTHER": "Other / MRO"}

# Ledger entry types that represent usage of an item.
CONSUMPTION_TYPES = {"Consumption", "Assembly Consumption"}   # RM and WIP usage
DEMAND_TYPES = {"Sale"}                                        # FG demand (expected demand = recent outward sales)

# Cover thresholds, in days.
COVER_CRITICAL = 7
COVER_LOW = 14
COVER_TARGET = 30
COVER_EXCESS = 180

PROJECTION_WEEKS = 12

# Locations whose stock is physically moving between sites rather than held.
# (Also any location flagged "Use As In-Transit" in Business Central.)
IN_TRANSIT_LOCATIONS = {"MIT"}


def site_of(location_code: str) -> str:
    """Group locations into sites: SHIP4RM, SHIP4PR, SHIP4FG -> SHIP 4. Others stay as themselves."""
    code = (location_code or "").upper()
    if code.startswith("SHIP") and len(code) > 4 and code[4].isdigit():
        return f"SHIP {code[4]}"
    return code or "(blank)"
