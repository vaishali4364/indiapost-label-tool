import io
import re
import numpy as np
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="India Post Label Generator - Sharmex Global",
    page_icon="📦",
    layout="wide",
)

st.title("📦 India Post Bulk Label Generator")
st.caption(
    "Updated with new 78-column SubPieces template, strict invoice valuation, and dynamic UK (128) / US (96) exchange rates."
)


# ========================================================
# 1. PRODUCT SPECIFICATION & WEIGHT MAPPING ENGINE
# ========================================================
def get_item_specs(title):
    t = str(title).lower()

    # --- Fragrances & Perfumes (Bella Vita, Wild Stone, Fogg) ---
    if any(k in t for k in ["bella vita", "bellavita", "perfume", "oud", "edp"]):
        hs = "33030040"
        if any(
            k in t
            for k in ["combo", "2 pack", "2 x 100ml", "ceo + goat", "ceo and goat"]
        ):
            return {
                "desc": "BV CEO & GOAT 100ml 2Pk",
                "hs": hs,
                "wt": 848,
                "val": 700,
            }
        if any(k in t for k in ["gift set", "4 x 20ml", "4x20ml", "all star"]):
            if any(
                k in t for k in ["pack of 2", "2 bella vita", "set 2", "set of 2"]
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
        if any(
            k in t
            for k in [
                "2 x 20ml",
                "2x20ml",
                "white and honey",
                "white oud & honey",
                "fresh + white",
                "date and senorita",
                "ceo + white",
                "date & glam",
            ]
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
        if any(k in t for k in ["dark oud", "oud dark", "oud gold"]):
            return {
                "desc": "BV Dark Oud EDP 100ml",
                "hs": hs,
                "wt": 498,
                "val": 486,
            }
        if "intense" in t:
            return {
                "desc": "BV CEO Intense EDP 100ml",
                "hs": hs,
                "wt": 398,
                "val": 450,
            }
        return {
            "desc": "BV Eau De Parfum 100ml",
            "hs": hs,
            "wt": 398,
            "val": 350,
        }

    # --- Skincare, Serums & Face Packs ---
    if any(k in t for k in ["detan", "moroccon", "face pack", "sayy"]):
        return {
            "desc": "Sayy Moroccan DeTan Mask",
            "hs": "33049090",
            "wt": 140,
            "val": 210,
        }
    if any(k in t for k in ["elvive", "serum", "loreal paris", "biolage"]):
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
    if any(k in t for k in ["gluta hya", "gluta-hya", "vaseline"]):
        if any(k in t for k in ["bundle", "pack of 3", "(3)"]):
            return {
                "desc": "Vaseline Gluta Hya 3Pk",
                "hs": "33049090",
                "wt": 645,
                "val": 826,
            }
        if "400ml" in t or "sun protect" in t:
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
    if any(k in t for k in ["roll on", "roll-on", "chemist at play", "rexona"]):
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
    if any(k in t for k in ["kiwi", "cherry blossom", "shoe polish"]):
        if any(k in t for k in ["instant", "liquid"]):
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
        if any(k in t for k in ["pot", "jug"]):
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
    if any(k in t for k in ["guard", "mach 3", "fusion", "vector", "razor"]):
        return {
            "desc": "Gillette Shaving Razor",
            "hs": "82121010",
            "wt": 98,
            "val": 260,
        }

    # --- Oral Care ---
    if any(k in t for k in ["toothbrush", "oral b", "oral-b"]):
        return {
            "desc": "Oral Care Toothbrush Pack",
            "hs": "96032100",
            "wt": 98,
            "val": 213,
        }
    if any(k in t for k in ["paste", "dabur", "colgate"]):
        return {
            "desc": "Dental Toothpaste Tube",
            "hs": "33061020",
            "wt": 198,
            "val": 210,
        }

    # Default fallback
    return {
        "desc": str(title)[:28],
        "hs": "33049990",
        "wt": 140,
        "val": 250,
    }


# ========================================================
# 2. FILE UPLOADER & PROCESSING PIPELINE
# ========================================================
uploaded_file = st.file_uploader(
    "Upload eBay Orders CSV Report", type=["csv"]
)

if uploaded_file is not None:
    try:
        content = uploaded_file.getvalue().decode("utf-8-sig", errors="ignore")
        lines = content.splitlines()

        header_idx = -1
        for i, line in enumerate(lines[:10]):
            if "Order Number" in line or "Order number" in line:
                header_idx = i
                break

        if header_idx == -1:
            st.error("Could not find 'Order Number' header in the CSV.")
            st.stop()

        df_raw = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])))
        df_raw.columns = [str(c).strip() for c in df_raw.columns]

        order_col = (
            "Order Number" if "Order Number" in df_raw.columns else "Order number"
        )
        valid_orders = df_raw[
            df_raw[order_col]
            .astype(str)
            .str.contains(r"^\d{2}-\d{5}-\d{5}$", regex=True)
        ].copy()

        if valid_orders.empty:
            st.warning("No valid eBay orders found in the uploaded file.")
            st.stop()

        st.success(
            f"Found {len(valid_orders[order_col].unique())} unique orders to process!"
        )

        article_rows = []
        subpiece_rows = []
        order_serial = 1

        for order_id, order_group in valid_orders.groupby(
            order_col, sort=False
        ):
            title_col = (
                "Item Title"
                if "Item Title" in order_group.columns
                else "Item title"
            )
            item_rows = order_group[
                order_group[title_col].fillna("").str.strip() != ""
            ]
            if item_rows.empty:
                item_rows = order_group.head(1)

            first_row = order_group.iloc[0]

            # --- DYNAMIC COUNTRY & EXCHANGE RATE DETECTION ---
            country_raw = str(
                first_row.get(
                    "Ship to Country",
                    first_row.get(
                        "Country", first_row.get("Ship To Country", "")
                    ),
                )
            ).strip()

            if any(
                c in country_raw.lower()
                for c in ["us", "usa", "united states", "america"]
            ):
                dest_code = "US"
                dest_name = "United States"
                currency_code = "USD"
                ex_rate = 96.0  # Updated to 96 as requested
            else:
                dest_code = "GB"
                dest_name = "United Kingdom"
                currency_code = "GBP"
                ex_rate = 128.0

            parcel_total_weight = 0
            parcel_total_declared_val = 0
            current_order_subpieces = []

            for lsn, (_, item) in enumerate(item_rows.iterrows(), start=1):
                item_title = str(item.get(title_col, ""))
                num_col = (
                    "Item Number"
                    if "Item Number" in item.index
                    else "Item number"
                )
                item_id = str(item.get(num_col, "")).strip()

                qty_col = (
                    "Quantity" if "Quantity" in item.index else "Item Quantity"
                )
                try:
                    qty = int(float(item.get(qty_col, 1)))
                    if qty < 1:
                        qty = 1
                except:
                    qty = 1

                specs = get_item_specs(item_title)

                item_total_weight = specs["wt"] * qty
                unit_declared_val = specs["val"]
                item_total_declared_val = unit_declared_val * qty

                sold_col = "Sold For" if "Sold For" in item.index else "Sold for"
                try:
                    raw_p = (
                        str(item.get(sold_col, "0"))
                        .replace("£", "")
                        .replace("$", "")
                        .replace(",", "")
                        .strip()
                    )
                    unit_fob = float(raw_p)
                except:
                    unit_fob = round(unit_declared_val / ex_rate, 2)
                total_fob = round(unit_fob * qty, 2)

                parcel_total_weight += item_total_weight
                parcel_total_declared_val += item_total_declared_val

                # New 78-column SubPieces row
                current_order_subpieces.append(
                    {
                        "SERIAL NUMBER REF ": order_serial,
                        "HS CODE": specs["hs"],
                        "CTH CODE": specs["hs"],
                        "HS DESCRIPTION": specs["desc"],
                        "SP UNIT CD": "PC",
                        "SP COUNT": qty,
                        "SP WEIGHT TOTAL": item_total_weight,
                        "SP NET WEIGHT": item_total_weight,
                        "SP ORIGIN COUNTRY CODE": "IN",
                        "SP ORIGIN CURRENCY CODE": "INR",
                        "SP COMM INVOICE NO": order_id,
                        "SP COM INVOICE DATE(DD-MM-YYYY)": pd.Timestamp.now().strftime(
                            "%Y-%m-%d"
                        ),
                        "SP INVOICE LSN": lsn,
                        "SP INV CURRENCY CODE": currency_code,
                        "SP INV EXCHANGE RATE": int(ex_rate),
                        "SP ASBL FOB VALUE": total_fob,
                        "SP ASBL INR VAL": item_total_declared_val,
                        "SP TAX INVOICE NO": order_id,
                        "SP TAX INVOICE DATE": pd.Timestamp.now().strftime(
                            "%Y-%m-%d"
                        ),
                        "SP INV VALUE PU": unit_declared_val,
                        "SP INV VALUE TOTAL": item_total_declared_val,
                        "ecommerce_url": "Ebay.com",
                        "ecommerce_paytranid": np.nan,
                        "ecommerce_sku": item_id,
                    }
                )

            buyer_name = str(
                first_row.get(
                    "Buyer Name", first_row.get("Ship To Name", "")
                )
            ).strip()
            add1 = str(
                first_row.get(
                    "Buyer Address 1", first_row.get("Ship To Address 1", "")
                )
            ).strip()
            add2 = str(
                first_row.get(
                    "Buyer Address 2", first_row.get("Ship To Address 2", "")
                )
            ).strip()
            city = str(
                first_row.get(
                    "Buyer City", first_row.get("Ship To City", "")
                )
            ).strip()
            state = str(
                first_row.get(
                    "Buyer State", first_row.get("Ship To State", "")
                )
            ).strip()
            zipcode = str(
                first_row.get(
                    "Buyer Postcode",
                    first_row.get(
                        "Ship To Zip", first_row.get("Ship To Postcode", "")
                    ),
                )
            ).strip()
            phone = str(
                first_row.get(
                    "Buyer Phone", first_row.get("Ship To Phone", "0000000000")
                )
            ).strip()

            # New 46-column ArticleDetails row (Country Name then Code, plus INCOTERMS)
            article_rows.append(
                {
                    "SERIAL NUMBER": order_serial,
                    "ARTICLE NUMBER": np.nan,
                    "DESTINATION COUNTRY NAME": dest_name,
                    "DESTINATION COUNTRY CODE": dest_code,
                    "MAIL NATURE TYPE": 31,
                    "MAIL TRANSPORT TYPE": 1,
                    "PHYSICAL WEIGHT": parcel_total_weight,
                    "DECLARED VALUE": parcel_total_declared_val,
                    "NON DELIVERY INSTRUCTIONS": 2,
                    "SENDER NAME": "NEELA SHARMA",
                    "SENDER COMPANY": "SHARMEX GLOBAL",
                    "SENDER ADD LINE 1": "1037, NEW MODEL TOWN",
                    "SENDER ADD LINE 2": "PINJORE",
                    "SENDER ADD LINE 3": np.nan,
                    "SENDER CITY": "PANCHKULA",
                    "SENDER STATE": "HR",
                    "SENDER COUNTRY NAME": "India",
                    "SENDER COUNTRY CODE": "IN",
                    "SENDER PINCODE": 134102,
                    "SENDER EMAILID": "sharmexglobal@gmail.com",
                    "SENDER MOBILE": 8629056095,
                    "RECEIVER NAME": buyer_name,
                    "RECEIVER ADD LINE 1": add1,
                    "RECEIVER ADD LINE 2": add2,
                    "RECEIVER CITY": city,
                    "RECEIVER STATE": state,
                    "RECEIVER ZIPCODE": zipcode,
                    "RECEIVER MOBILE NO": phone if phone != "nan" else "0000000000",
                    "DROP OFF PINCODE": 136118,
                    "PBE TYPE": "PBEIII",
                    "PBE FILING": "SELF",
                    "INCOTERMS": "DAP",
                }
            )

            subpiece_rows.extend(current_order_subpieces)
            order_serial += 1

        df_art_final = pd.DataFrame(article_rows)
        df_sub_final = pd.DataFrame(subpiece_rows)

        # Build Output Excel matching exact template column order
        output_buffer = io.BytesIO()

        # Load column order from template.xlsx to preserve all 78 SubPiece and 46 Article columns
        tpl_path = "template.xlsx"
        try:
            wb_tpl = openpyxl.load_workbook(tpl_path, data_only=True)
            ref_art_cols = [
                cell.value for cell in wb_tpl["ArticleDetails"][1] if cell.value
            ]
            ref_sub_cols = [
                cell.value for cell in wb_tpl["SubPieces"][1] if cell.value
            ]
        except:
            ref_art_cols = list(df_art_final.columns)
            ref_sub_cols = list(df_sub_final.columns)

        # Reindex to ensure strict alignment with official template headers
        df_art_export = df_art_final.reindex(columns=ref_art_cols)
        df_sub_export = df_sub_final.reindex(columns=ref_sub_cols)

        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            df_art_export.to_excel(
                writer, sheet_name="ArticleDetails", index=False
            )
            df_sub_export.to_excel(writer, sheet_name="SubPieces", index=False)

            # Copy reference meta sheets
            try:
                for sheet in ["PickupAddress", "Information"]:
                    if sheet in wb_tpl.sheetnames:
                        df_extra = pd.read_excel(tpl_path, sheet_name=sheet)
                        df_extra.to_excel(writer, sheet_name=sheet, index=False)
            except:
                pass

        st.subheader("📋 Output Verification")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Total Parcels (Articles)", len(df_art_export))
            st.dataframe(
                df_art_export[
                    [
                        "SERIAL NUMBER",
                        "RECEIVER NAME",
                        "DESTINATION COUNTRY NAME",
                        "DESTINATION COUNTRY CODE",
                        "PHYSICAL WEIGHT",
                        "DECLARED VALUE",
                    ]
                ]
            )
        with c2:
            st.metric("Total SubPieces", len(df_sub_export))
            st.dataframe(
                df_sub_export[
                    [
                        "SERIAL NUMBER REF ",
                        "HS DESCRIPTION",
                        "SP COUNT",
                        "SP INV CURRENCY CODE",
                        "SP INV EXCHANGE RATE",
                        "SP INV VALUE TOTAL",
                    ]
                ]
            )

        export_date = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="📥 Download India Post Excel File",
            data=output_buffer.getvalue(),
            file_name=f"IndiaPost_BulkUpload_{export_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error processing orders: {str(e)}")