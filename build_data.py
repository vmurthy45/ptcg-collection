#!/usr/bin/env python3
"""
Build data/collection.json for the Pokémon TCG Collection website
from the "PTCG Purchases Tracker.xlsx" spreadsheet.

Workflow for updates:
    1. Keep the xlsx updated.
    2. Run:  python3 build_data.py  [optional/path/to/tracker.xlsx]
    3. Commit & push  data/collection.json  (bump sw.js cache name).

Reads two sheets:
    "Card Purchases" -> one row per card line (name, qty, all-in cost per card)
    "Needs"          -> quantity + card name of cards still wanted

Cards are grouped by a normalised BASE NAME so that different printings /
variants of the same card sit together, while each distinct printed name
(including its [RH]/[PP]/[Holo]/[FA]/(pattern)/set-code tags) stays a
separate variant row.
"""
import json
import os
import re
import sys
from collections import OrderedDict

DEFAULT_XLSX = "/Users/vighnesh/Library/CloudStorage/OneDrive-Personal/01. Documents/05. Games & Books/TCG/PTCG Purchases Tracker.xlsx"
OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "collection.json")

# --- variant token expansion (display only; unknown tokens shown verbatim) ---
TOKEN_MAP = {
    "RH": "Reverse Holo",
    "PP": "Prize Pack",
    "Holo": "Holo",
    "FA": "Full Art",
    "PB": "Poké Ball",
}


def clean_ws(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def parse_name(full: str):
    """Return (base_name, [variant_tags]) for a printed card name."""
    tags = []
    work = full

    # 1) parenthetical tags anywhere e.g. (Ascended Heroes), (Poke Ball), (Energy Symbol Pattern)
    for m in re.findall(r"\(([^)]*)\)", work):
        t = clean_ws(m)
        if t:
            tags.append(t)
    work = re.sub(r"\([^)]*\)", " ", work)

    # 2) set code e.g. " - 141/167", " - 025"
    setcodes = re.findall(r"-\s*(\d+(?:/\w+)?)\b", work)
    for sc in setcodes:
        tags.append("#" + sc)
    work = re.sub(r"-\s*\d+(?:/\w+)?\b", " ", work)

    # 3) bracket tags e.g. [RH], [PP], [Holo], [Corbeau]
    for m in re.findall(r"\[([^\]]*)\]", work):
        t = clean_ws(m)
        if t:
            tags.append(t)
    work = re.sub(r"\[[^\]]*\]", " ", work)

    # 4) trailing bare " FA" (Full Art) e.g. "Oricorio ex FA", "Tapu Lele GX FA"
    if re.search(r"\bFA\s*$", work):
        tags.append("FA")
        work = re.sub(r"\bFA\s*$", " ", work)

    base = clean_ws(work)
    if not base:  # safety: never blank
        base = clean_ws(full)
    return base, tags


def variant_label(tags):
    """Human-readable variant chip. Empty tags -> 'Standard'."""
    if not tags:
        return "Standard"
    out = []
    for t in tags:
        if t.startswith("#"):
            out.append("No. " + t[1:])
        else:
            out.append(TOKEN_MAP.get(t, t))
    return " · ".join(out)


def main():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX
    if not os.path.exists(xlsx):
        sys.exit(f"ERROR: spreadsheet not found: {xlsx}")
    try:
        import openpyxl
    except ImportError:
        sys.exit("ERROR: openpyxl not installed. Run: python3 -m pip install openpyxl --break-system-packages")

    wb = openpyxl.load_workbook(xlsx, data_only=True)

    # ---- Card Purchases -> variants ----
    cp = wb["Card Purchases"]
    rows = list(cp.iter_rows(values_only=True))
    header = [str(h).strip() if h else "" for h in rows[0]]
    col = {name: i for i, name in enumerate(header)}
    ci_name = col["Card Name"]
    ci_qty = col["Quantity"]
    ci_total = col["Total Cost"]        # all-in (incl. shipping/tax/fees allocated)
    ci_price = col["Price per Card"]    # sticker price per card (pre-allocation)

    # aggregate by exact printed name (variant identity)
    variants = OrderedDict()
    for r in rows[1:]:
        name = r[ci_name]
        if not name:
            continue
        name = clean_ws(str(name))
        qty = r[ci_qty] or 0
        total = r[ci_total] or 0.0
        price = r[ci_price] or 0.0
        v = variants.get(name)
        if v is None:
            base, tags = parse_name(name)
            v = variants[name] = {
                "name": name,
                "base": base,
                "variant": variant_label(tags),
                "qty": 0,
                "allInCost": 0.0,       # sum of all-in Total Cost
                "stickerCost": 0.0,     # sum of price*qty (sticker only)
            }
        v["qty"] += qty
        v["allInCost"] += total
        v["stickerCost"] += price * qty

    # ---- group variants by base name ----
    groups = OrderedDict()
    for v in variants.values():
        g = groups.get(v["base"])
        if g is None:
            g = groups[v["base"]] = {"base": v["base"], "total": 0, "variants": []}
        g["total"] += v["qty"]
        g["variants"].append({
            "name": v["name"],
            "variant": v["variant"],
            "qty": v["qty"],
            "cost": round(v["allInCost"], 2),   # all-in total (shipping/tax/fees allocated)
            "avgAllIn": round(v["allInCost"] / v["qty"], 2) if v["qty"] else 0.0,
        })

    group_list = list(groups.values())
    # sort variants within a group: Standard first, then by name
    for g in group_list:
        g["variants"].sort(key=lambda x: (x["variant"] != "Standard", x["name"].lower()))
    group_list.sort(key=lambda g: g["base"].lower())

    total_cards = sum(v["qty"] for v in variants.values())
    total_all_in = sum(v["allInCost"] for v in variants.values())

    # ---- Needs ----
    needs = []
    if "Needs" in wb.sheetnames:
        for r in list(wb["Needs"].iter_rows(values_only=True))[1:]:
            if not r or not r[1]:
                continue
            needs.append({"qty": r[0] or 0, "name": clean_ws(str(r[1]))})
        needs.sort(key=lambda x: x["name"].lower())

    from datetime import datetime, timezone
    data = {
        "updated": datetime.now(timezone.utc).strftime("%Y-%m-%d"),
        "currency": "NZD",
        "stats": {
            "totalCards": total_cards,
            "uniqueVariants": len(variants),
            "uniqueNames": len(groups),
            "totalSpent": round(total_all_in, 2),
            "avgCostPerCard": round(total_all_in / total_cards, 2) if total_cards else 0.0,
            "needsCount": sum(n["qty"] for n in needs),
        },
        "groups": group_list,
        "needs": needs,
    }

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)

    print(f"Wrote {OUT_PATH}")
    print(f"  {total_cards} cards | {len(variants)} variants | {len(groups)} names | ${total_all_in:.2f} {data['currency']}")
    print(f"  needs: {len(needs)} lines / {data['stats']['needsCount']} cards")


if __name__ == "__main__":
    main()
