import math
import numpy as np
import pandas as pd

# Load master catalog
catalog_df = pd.read_excel("product_catalog.xlsx")
# Ensure item_id is string and stripped
catalog_df["item_id"] = catalog_df["item_id"].astype(str).str.strip()


def lookup_product(item_num, title):
    """Matches product by 12-digit Item Number first, then fallback to title keywords."""
    clean_num = str(item_num).strip()
    match = catalog_df[catalog_df["item_id"] == clean_num]
    if not match.empty:
        r = match.iloc[0]
        return {
            "hs_code": str(r.get("hs_code", "33049990")),
            "cth_code": str(r.get("cth_code", "33049990")),
            "hs_desc": str(r.get("hs_description", title[:28])),
            "weight": int(r.get("weight_grams", 250)),
            "declared_val": int(r.get("declared_value", 350)),
        }

    # Fallback to keyword matching if ID not found
    t = str(title).lower()
    for _, r in catalog_df.iterrows():
        kw = str(r.get("title_keywor", "")).strip().lower()
        if kw and kw in t:
            return {
                "hs_code": str(r.get("hs_code", "33049990")),
                "cth_code": str(r.get("cth_code", "33049990")),
                "hs_desc": str(r.get("hs_description", title[:28])),
                "weight": int(r.get("weight_grams", 250)),
                "declared_val": int(r.get("declared_value", 350)),
            }

    # Safe defaults
    return {
        "hs_code": "33049990",
        "cth_code": "33049990",
        "hs_desc": str(title)[:28],
        "weight": 250,
        "declared_val": 350,
    }


# ========================================================
# BUILD ARTICLEDETAILS & SUBPIECES
# ========================================================
article_rows = []
subpiece_rows = []

# Group orders by Order Number
order_serial = 1

for order_id, order_group in valid_orders.groupby("Order number", sort=False):
    # Filter rows that actually have an Item Title / Item Number (skipping empty transaction summary rows)
    item_rows = order_group[
        order_group["Item title"].fillna("").str.strip() != ""
    ]
    if item_rows.empty:
        item_rows = order_group.head(1)

    first_row = order_group.iloc[0]

    order_total_weight = 0
    order_total_declared_val = 0
    order_subpieces = []

    # Loop through EVERY item in this order
    for lsn, (_, item) in enumerate(item_rows.iterrows(), start=1):
        # Extract item details
        item_title = str(item.get("Item title", ""))
        item_id = str(item.get("Item number", "")).strip()

        # Parse Quantity safely
        try:
            qty = int(float(item.get("Quantity", 1)))
            if qty < 1:
                qty = 1
        except:
            qty = 1

        # Look up product in catalog
        prod = lookup_product(item_id, item_title)

        # Calculate exact weights and values with Quantity
        unit_weight = prod["weight"]
        total_item_weight = unit_weight * qty

        unit_declared_val = prod["declared_val"]
        total_item_declared_val = (
            unit_declared_val * qty
        )  # Multiplied by quantity!

        # eBay Sold Price (for FOB value reference)
        try:
            sold_price_raw = str(item.get("Sold for", "0")).replace("£", "").replace("$", "").strip()
            unit_fob = float(sold_price_raw)
        except:
            unit_fob = round(unit_declared_val / 128.0, 2)
        total_fob = round(unit_fob * qty, 2)

        order_total_weight += total_item_weight
        order_total_declared_val += total_item_declared_val

        # Append SubPiece row
        order_subpieces.append(
            {
                "SERIAL NUMBER REF ": order_serial,
                "HS CODE": prod["hs_code"],
                "CTH CODE": prod["cth_code"],
                "HS DESCRIPTION": prod["hs_desc"],
                "SP UNIT CD": "PC",
                "SP COUNT": qty,
                "SP WEIGHT TOTAL": total_item_weight,
                "SP NET WEIGHT": total_item_weight,
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
                "SP ASBL INR VAL": total_item_declared_val,  # MUST equal total value
                "SP TAX INVOICE NO": order_id,
                "SP TAX INVOICE DATE": pd.Timestamp.now().strftime("%Y-%m-%d"),
                "SP INV VALUE PU": unit_declared_val,  # Per unit
                "SP INV VALUE TOTAL": total_item_declared_val,  # Total = Unit * Qty
                "ecommerce_url": "Ebay.com",
                "ecommerce_sku": item_id,
            }
        )

    # 1 Article row per parcel/order
    article_rows.append(
        {
            "SERIAL NUMBER": order_serial,
            "ARTICLE NUMBER": np.nan,
            "DESTINATION COUNTRY CODE": first_row.get(
                "Country Code", "GB"
            ),  # e.g. GB
            "DESTINATION COUNTRY NAME": first_row.get(
                "Country Name", "United Kingdom"
            ),
            "MAIL NATURE TYPE": 31,  # Sale of Goods
            "MAIL TRANSPORT TYPE": 1,  # Air
            "PHYSICAL WEIGHT": order_total_weight,  # Exact sum of all item weights
            "DECLARED VALUE": order_total_declared_val,  # Exact sum of all subpiece values
            "NON DELIVERY INSTRUCTIONS": 2,  # Return to sender
            # Sender Details...
            "SENDER NAME": "Vaishali Sharma",
            "SENDER COMPANY": "Sharmex Global",
            # Receiver Details...
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

    # Add all subpieces for this article
    subpiece_rows.extend(order_subpieces)

    order_serial += 1

df_art_final = pd.DataFrame(article_rows)
df_sub_final = pd.DataFrame(subpiece_rows)