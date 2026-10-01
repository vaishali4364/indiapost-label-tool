import math
import numpy as np
import pandas as pd


# 1. Exact base weights (grams) and declared customs values (INR) from Sharmex Global dispatch log
def get_exact_item_specs(title, item_id=""):
    t = str(title).lower()

    # --- Fragrances & Perfumes (Bella Vita, Wild Stone, Fogg) ---
    if any(k in t for k in ["bella vita", "bellavita", "oud", "perfume", "edp"]):
        hs = "33030040"
        if "combo" in t or "2 pack" in t or "2 x 100ml" in t or "ceo + goat" in t:
            return {
                "desc": "BV CEO & GOAT 100ml 2Pk",
                "hs": hs,
                "wt": 848,
                "val": 700,
            }
        if "set" in t or "4 x 20ml" in t or "4x20ml" in t:
            if "2" in t and (
                "pack of 2" in t or "2 bella vita" in t or "set 2" in t
            ):
                return {
                    "desc": "BV Gift Set 4x20ml 2Pk",
                    "hs": hs,
                    "wt": 848,
                    "val": 700,
                }
            return {
                "desc": "BV Perfume Gift Set 4x20ml",
                "hs": hs,
                "wt": 398,
                "val": 350,
            }
        if (
            "2 x 20ml" in t
            or "2x20ml" in t
            or "white and honey" in t
            or "fresh + white" in t
            or "date and senorita" in t
            or "ceo + white" in t
        ):
            return {
                "desc": "BV Perfume Duo 2x20ml",
                "hs": hs,
                "wt": 198,
                "val": 360,
            }
        if "20ml" in t:
            return {
                "desc": "BV Pocket Perfume 20ml",
                "hs": hs,
                "wt": 98,
                "val": 180,
            }
        if "dark oud" in t or "oud dark" in t or "oud gold" in t:
            return {
                "desc": "BV Dark Oud EDP 100ml",
                "hs": hs,
                "wt": 498,
                "val": 486,
            }
        if "ceo" in t and "intense" in t:
            return {
                "desc": "BV CEO Intense EDP 100ml",
                "hs": hs,
                "wt": 398,
                "val": 450,
            }
        if "goat" in t or "date" in t or "ceo" in t or "white oud" in t:
            return {
                "desc": "BV Eau De Parfum 100ml",
                "hs": hs,
                "wt": 398,
                "val": 350,
            }
        return {
            "desc": "BV Eau De Parfum 100ml",
            "hs": hs,
            "wt": 398,
            "val": 350,
        }

    # --- Skincare, Serums & Face Masks ---
    if "detan" in t or "moroccon" in t or "face pack" in t or "sayy" in t:
        return {
            "desc": "Sayy Moroccan DeTan Mask",
            "hs": "33049090",
            "wt": 140,
            "val": 210,
        }
    if "elvive" in t or "serum" in t or "loreal" in t:
        if "combo" in t:
            return {
                "desc": "Loreal Elvive Hair Combo",
                "hs": "33059090",
                "wt": 548,
                "val": 506,
            }
        return {
            "desc": "Loreal Paris Elvive Serum",
            "hs": "33059090",
            "wt": 140,
            "val": 434,
        }
    if "gluta hya" in t or "vaseline" in t:
        if "bundle" in t or "pack of 3" in t or "(3)" in t:
            return {
                "desc": "Vaseline Gluta Hya 3Pk",
                "hs": "33049090",
                "wt": 645,
                "val": 826,
            }
        if "sun protect" in t or "400ml" in t:
            return {
                "desc": "Vaseline Body Lotion 400ml",
                "hs": "33049090",
                "wt": 480,
                "val": 340,
            }
        return {
            "desc": "Vaseline Gluta Hya Lotion",
            "hs": "33049090",
            "wt": 249,
            "val": 255,
        }
    if "chemist at play" in t or "roll on" in t or "rexona" in t:
        return {
            "desc": "Deodorant Underarm Roll On",
            "hs": "33072000",
            "wt": 98,
            "val": 291,
        }
    if "betadine" in t:
        return {
            "desc": "Betadine Antiseptic 100ml",
            "hs": "30049099",
            "wt": 140,
            "val": 90,
        }
    if "supradyn" in t:
        return {
            "desc": "Supradyn Multivitamin 60s",
            "hs": "21069099",
            "wt": 98,
            "val": 250,
        }

    # --- Shoe Polish & Sponges ---
    if "sponge" in t:
        return {
            "desc": "Kiwi Shoe Shine Sponge",
            "hs": "34051000",
            "wt": 98,
            "val": 170,
        }
    if "kiwi" in t or "cherry blossom" in t or "polish" in t:
        if "instant" in t or "liquid" in t:
            return {
                "desc": "Kiwi Instant Shoe Polish",
                "hs": "34051000",
                "wt": 98,
                "val": 106,
            }
        return {
            "desc": "Kiwi Shoe Polish Wax Tin",
            "hs": "34051000",
            "wt": 120,
            "val": 106,
        }

    # --- Copperware ---
    if "copper" in t:
        if "pot" in t or "jug" in t:
            return {
                "desc": "Pure Copper Water Pot",
                "hs": "74198090",
                "wt": 326,
                "val": 945,
            }
        return {
            "desc": "Ayurvedic Copper Ball",
            "hs": "74198090",
            "wt": 49,
            "val": 162,
        }

    # --- Shaving Razors & Blades ---
    if any(k in t for k in ["mach 3", "fusion", "vector", "guard", "razor"]):
        return {
            "desc": "Gillette Shaving Razor",
            "hs": "82121010",
            "wt": 98,
            "val": 260,
        }

    # --- Toothpaste & Toothbrushes ---
    if "toothbrush" in t or "oral b" in t:
        return {
            "desc": "Oral Care Toothbrush Pack",
            "hs": "96032100",
            "wt": 98,
            "val": 213,
        }
    if "paste" in t or "dabur" in t or "colgate" in t:
        return {
            "desc": "Dental Toothpaste Tube",
            "hs": "33061020",
            "wt": 198,
            "val": 210,
        }

    # Default fallback
    return {
        "desc": str(title)[:28],
        "hs": "33049090",
        "wt": 140,
        "val": 250,
    }


# ========================================================
# BUILD ARTICLEDETAILS & SUBPIECES ACCURATELY
# ========================================================
article_rows = []
subpiece_rows = []
order_serial = 1

for order_id, order_group in valid_orders.groupby("Order number", sort=False):
    # Only keep line items that have a real product title
    item_rows = order_group[
        order_group["Item title"].fillna("").str.strip() != ""
    ]
    if item_rows.empty:
        item_rows = order_group.head(1)

    first_row = order_group.iloc[0]

    parcel_total_weight = 0
    parcel_total_declared_val = 0
    current_order_subpieces = []

    # Process EVERY unique item in this order
    for lsn, (_, item) in enumerate(item_rows.iterrows(), start=1):
        item_title = str(item.get("Item title", ""))
        item_id = str(item.get("Item number", "")).strip()

        # Parse Quantity
        try:
            qty = int(float(item.get("Quantity", 1)))
            if qty < 1:
                qty = 1
        except:
            qty = 1

        # Get exact specs
        specs = get_exact_item_specs(item_title, item_id)

        # 1. Total Weight for this line item (unit weight * quantity)
        item_total_weight = specs["wt"] * qty

        # 2. Total Declared Value for this line item (unit value * quantity)
        unit_declared_val = specs["val"]
        item_total_declared_val = unit_declared_val * qty

        # 3. FOB calculation (in GBP/USD converted or proportional)
        try:
            sold_price_raw = (
                str(item.get("Sold for", "0"))
                .replace("£", "")
                .replace("$", "")
                .strip()
            )
            unit_fob = float(sold_price_raw)
        except:
            unit_fob = round(unit_declared_val / 128.0, 2)
        total_fob = round(unit_fob * qty, 2)

        parcel_total_weight += item_total_weight
        parcel_total_declared_val += item_total_declared_val

        # SubPiece entry (One row per unique item line)
        current_order_subpieces.append(
            {
                "SERIAL NUMBER REF ": order_serial,
                "HS CODE": specs["hs"],
                "CTH CODE": specs["hs"],
                "HS DESCRIPTION": specs["desc"],
                "SP UNIT CD": "PC",
                "SP COUNT": qty,
                "SP WEIGHT TOTAL": item_total_weight,  # Exact total line weight
                "SP NET WEIGHT": item_total_weight,
                "SP ORIGIN COUNTRY CODE": "IN",
                "SP ORIGIN CURRENCY CODE": "INR",
                "SP COMM INVOICE NO": order_id,
                "SP COM INVOICE DATE(DD-MM-YYYY)": pd.Timestamp.now().strftime(
                    "%Y-%m-%d"
                ),
                "SP INVOICE LSN": lsn,
                "SP INV CURRENCY CODE": "GBP",
                "SP INV EXCHANGE RATE": 128,
                "SP ASBL FOB VALUE": total_fob,
                "SP ASBL INR VAL": item_total_declared_val,  # MUST match total value
                "SP TAX INVOICE NO": order_id,
                "SP TAX INVOICE DATE": pd.Timestamp.now().strftime("%Y-%m-%d"),
                "SP INV VALUE PU": unit_declared_val,  # Unit price
                "SP INV VALUE TOTAL": item_total_declared_val,  # Total price = Unit * Qty
                "ecommerce_url": "Ebay.com",
                "ecommerce_sku": item_id,
            }
        )

    # Article Details entry (One row per parcel)
    article_rows.append(
        {
            "SERIAL NUMBER": order_serial,
            "ARTICLE NUMBER": np.nan,
            "DESTINATION COUNTRY CODE": first_row.get("Country Code", "GB"),
            "DESTINATION COUNTRY NAME": first_row.get(
                "Country Name", "United Kingdom"
            ),
            "MAIL NATURE TYPE": 31,
            "MAIL TRANSPORT TYPE": 1,
            "PHYSICAL WEIGHT": parcel_total_weight,  # Sum of all item weights
            "DECLARED VALUE": parcel_total_declared_val,  # Exact sum of all SubPiece totals
            "NON DELIVERY INSTRUCTIONS": 2,
            "SENDER NAME": "Vaishali Sharma",
            "SENDER COMPANY": "Sharmex Global",
            "RECEIVER NAME": first_row.get("Buyer Name", ""),
            "RECEIVER ADD LINE 1": first_row.get("Buyer Address 1", ""),
            "RECEIVER ADD LINE 2": first_row.get("Buyer Address 2", ""),
            "RECEIVER CITY": first_row.get("Buyer City", ""),
            "RECEIVER STATE": first_row.get("Buyer State", ""),
            "RECEIVER ZIPCODE": first_row.get("Buyer Postcode", ""),
            "RECEIVER MOBILE NO": first_row.get("Buyer Phone", "0000000000"),
            "PBE TYPE": 3,
            "PBE FILING": "N",
        }
    )

    subpiece_rows.extend(current_order_subpieces)
    order_serial += 1

df_art_final = pd.DataFrame(article_rows)
df_sub_final = pd.DataFrame(subpiece_rows)