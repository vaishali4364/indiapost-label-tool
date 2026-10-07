import io
import re
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="India Post Label Generator - Sharmex Global",
    page_icon="📦",
    layout="wide",
)

st.title("📦 India Post Bulk Label Generator")
st.caption(
    "Automated bulk booking file creator with strict SubPiece value matching, "
    "Ship-To address resolution, and pink-highlighted multi-packs."
)


# Helper function to prevent literal 'nan' strings in Excel cells
def clean_str(val, fallback=""):
  if pd.isna(val) or str(val).strip().lower() == "nan":
    return fallback
  return str(val).strip()


# ========================================================
# 1. PRODUCT SPECIFICATION & WEIGHT MAPPING ENGINE
# ========================================================
def get_item_specs(title):
  t = str(title).lower()

  # Detect pack multiplier from title (e.g. "Pack of 2", "(2)", "2x", etc.)
  pack_mult = 1
  m_pack = re.search(r"pack\s*(?:of)?\s*(\d+)", t)
  m_lead = re.match(r"^(\d+)\s+", t)
  m_paren = re.search(r"[\(\[]\s*(\d+)\s*[\)\]]", t)
  m_x = re.search(r"(\d+)\s*x\s*(?!20ml|100ml)", t)

  if m_pack:
    pack_mult = int(m_pack.group(1))
  elif m_lead and int(m_lead.group(1)) in [2, 3, 4, 5, 6, 8]:
    pack_mult = int(m_lead.group(1))
  elif m_paren and int(m_paren.group(1)) in [2, 3, 4, 5, 6, 8]:
    pack_mult = int(m_paren.group(1))
  elif m_x and int(m_x.group(1)) in [2, 3, 4, 5, 6, 8]:
    pack_mult = int(m_x.group(1))

  # --- Olivia Herbal Bleach Cream ---
  if "olivia" in t and "bleach" in t:
    return {
        "desc": "Olivia Herbal Bleach Cream",
        "hs": "33049910",
        "wt": 49 * pack_mult,
        "val": 55 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Betadine Antiseptic ---
  if "betadine" in t:
    return {
        "desc": "Betadine Povidone-iodine",
        "hs": "30049099",
        "wt": 248 if pack_mult == 2 else 140 * pack_mult,
        "val": 90 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Fragrances & Perfumes (Bellavita, Wild Stone, Fogg) ---
  if any(k in t for k in ["bella vita", "bellavita", "perfume", "oud", "edp"]):
    hs = "33030040"
    if "intense" in t and ("ceo" in t or "100" in t):
      return {
          "desc": "Bellavita CEO Intense 100ml",
          "hs": hs,
          "wt": 448 * pack_mult,
          "val": 450 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    if any(
        k in t
        for k in ["combo", "2 pack", "2 x 100ml", "ceo + goat", "ceo and goat"]
    ):
      return {
          "desc": "Bellavita CEO & GOAT 2Pk",
          "hs": hs,
          "wt": 848,
          "val": 700,
          "is_multi": True,
      }
    if any(k in t for k in ["gift set", "4 x 20ml", "4x20ml", "all star"]):
      if (
          pack_mult == 2
          or "pack of 2" in t
          or "2 bella vita" in t
          or "set 2" in t
      ):
        return {
            "desc": "Bellavita Gift Set 4x20ml 2Pk",
            "hs": hs,
            "wt": 848,
            "val": 700,
            "is_multi": True,
        }
      return {
          "desc": "Bellavita Gift Set 4x20ml",
          "hs": hs,
          "wt": 398,
          "val": 350,
          "is_multi": False,
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
          "desc": "Bellavita Perfume Duo 2x20ml",
          "hs": hs,
          "wt": 198,
          "val": 360,
          "is_multi": True,
      }
    if "20ml" in t:
      return {
          "desc": "Bellavita Pocket Perfume 20ml",
          "hs": hs,
          "wt": 98 * pack_mult,
          "val": 180 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    if any(k in t for k in ["dark oud", "oud dark", "oud gold"]):
      return {
          "desc": "Bellavita Dark Oud EDP 100ml",
          "hs": hs,
          "wt": 498 * pack_mult,
          "val": 486 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    return {
        "desc": "Bellavita Eau De Parfum 100ml",
        "hs": hs,
        "wt": 398 * pack_mult,
        "val": 350 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Skincare, Serums & Face Packs ---
  if any(k in t for k in ["detan", "moroccon", "face pack", "sayy"]):
    return {
        "desc": "Sayy Moroccan DeTan Mask",
        "hs": "33049090",
        "wt": 140 * pack_mult,
        "val": 210 * pack_mult,
        "is_multi": pack_mult > 1,
    }
  if any(k in t for k in ["elvive", "serum", "loreal paris", "biolage"]):
    if "combo" in t:
      return {
          "desc": "Loreal Elvive Hair Combo",
          "hs": "33059090",
          "wt": 548,
          "val": 506,
          "is_multi": True,
      }
    if pack_mult > 1:
      wt_map = {2: 240, 3: 348, 4: 490}
      return {
          "desc": "Loreal Paris Elvive Serum",
          "hs": "33059090",
          "wt": wt_map.get(pack_mult, 140 * pack_mult),
          "val": 434 * pack_mult,
          "is_multi": True,
      }
    return {
        "desc": "Loreal Paris Elvive Serum",
        "hs": "33059090",
        "wt": 140,
        "val": 434,
        "is_multi": False,
    }
  if any(k in t for k in ["gluta hya", "gluta-hya", "vaseline"]):
    if any(k in t for k in ["bundle", "pack of 3", "(3)"]) or pack_mult == 3:
      return {
          "desc": "Vaseline Gluta Hya 3Pk",
          "hs": "33049090",
          "wt": 645,
          "val": 826,
          "is_multi": True,
      }
    if "400ml" in t or "sun protect" in t:
      return {
          "desc": "Vaseline Body Lotion 400ml",
          "hs": "33049090",
          "wt": 480 * pack_mult,
          "val": 340 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    if pack_mult > 1:
      return {
          "desc": "Vaseline Gluta Hya Lotion",
          "hs": "33049090",
          "wt": 398 if pack_mult == 2 else 249 * pack_mult,
          "val": 255 * pack_mult,
          "is_multi": True,
      }
    return {
        "desc": "Vaseline Gluta Hya Lotion",
        "hs": "33049090",
        "wt": 249,
        "val": 255,
        "is_multi": False,
    }
  if any(k in t for k in ["roll on", "roll-on", "chemist at play", "rexona"]):
    return {
        "desc": "Deodorant Underarm Roll On",
        "hs": "33072000",
        "wt": 98 * pack_mult,
        "val": 291 * pack_mult,
        "is_multi": pack_mult > 1,
    }
  if "supradyn" in t:
    return {
        "desc": "Supradyn Multivitamin 60s",
        "hs": "21069099",
        "wt": 98 * pack_mult,
        "val": 250 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Shoe Polish & Sponges ---
  if "sponge" in t:
    return {
        "desc": "Kiwi Shoe Shine Sponge",
        "hs": "34051000",
        "wt": 198 if pack_mult == 2 else 98 * pack_mult,
        "val": 170 * pack_mult,
        "is_multi": pack_mult > 1,
    }
  if any(k in t for k in ["kiwi", "cherry blossom", "shoe polish"]):
    if any(k in t for k in ["instant", "liquid"]):
      return {
          "desc": "Kiwi Instant Shoe Polish",
          "hs": "34051000",
          "wt": 198 if pack_mult == 2 else 98 * pack_mult,
          "val": 106 * pack_mult,
          "is_multi": pack_mult > 1,
      }
    return {
        "desc": "Kiwi Shoe Polish Wax Tin",
        "hs": "34051000",
        "wt": 248 if pack_mult == 2 else 120 * pack_mult,
        "val": 106 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Copperware ---
  if "copper" in t:
    if any(k in t for k in ["pot", "jug"]):
      return {
          "desc": "Pure Copper Water Pot",
          "hs": "74198090",
          "wt": 326,
          "val": 945,
          "is_multi": False,
      }
    wt_balls = 49 if pack_mult <= 2 else (98 if pack_mult <= 4 else 198)
    return {
        "desc": "Ayurvedic Copper Ball",
        "hs": "74198090",
        "wt": wt_balls,
        "val": 162 * (pack_mult // 2 if pack_mult > 1 else 1),
        "is_multi": pack_mult > 1,
    }

  # --- Shaving Razors & Blades ---
  if any(k in t for k in ["guard", "mach 3", "fusion", "vector", "razor"]):
    return {
        "desc": "Gillette Shaving Razor",
        "hs": "82121010",
        "wt": 198 if pack_mult == 2 else 98 * pack_mult,
        "val": 260 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # --- Oral Care ---
  if any(k in t for k in ["toothbrush", "oral b", "oral-b"]):
    return {
        "desc": "Oral Care Toothbrush Pack",
        "hs": "96032100",
        "wt": 198 if pack_mult == 2 else 98 * pack_mult,
        "val": 213 * pack_mult,
        "is_multi": pack_mult > 1,
    }
  if any(k in t for k in ["paste", "dabur", "colgate"]):
    return {
        "desc": "Dental Toothpaste Tube",
        "hs": "33061020",
        "wt": 198 * pack_mult,
        "val": 210 * pack_mult,
        "is_multi": pack_mult > 1,
    }

  # Default fallback
  return {
      "desc": str(title)[:28],
      "hs": "33049990",
      "wt": 140 * pack_mult,
      "val": 250 * pack_mult,
# --- ROBUST CASE-INSENSITIVE EBAY FIELD EXTRACTOR ---
      row_dict = {str(k).lower().strip(): v for k, v in first_row.items()}

      def get_field(candidate_keys, fallback=""):
        for k in candidate_keys:
          val = row_dict.get(k.lower().strip())
          s = clean_str(val)
          if s != "":
            return s
        return fallback

      # 1. Country & Exchange Rate
      country_raw = get_field([
          "ship to country",
          "shipping country",
          "buyer country",
          "country",
          "country/region",
      ])
      if any(
          c in country_raw.lower()
          for c in ["us", "usa", "united states", "america"]
      ):
        dest_code = "US"
        dest_name = "United States of America"
        currency_code = "USD"
        ex_rate = 96.0
      else:
        dest_code = "GB"
        dest_name = "United Kingdom"
        currency_code = "GBP"
        ex_rate = 128.0

      # 2. Receiver Address Fields (Handles all eBay formats)
      ship_name = get_field([
          "ship to name",
          "shipping name",
          "recipient name",
          "buyer name",
          "buyer full name",
          "contact name",
      ])
      add1 = get_field([
          "ship to address 1",
          "shipping address 1",
          "buyer address 1",
          "delivery address 1",
          "address 1",
          "street 1",
      ])
      add2 = get_field([
          "ship to address 2",
          "shipping address 2",
          "buyer address 2",
          "delivery address 2",
          "address 2",
          "street 2",
      ])
      city = get_field(
          ["ship to city", "shipping city", "buyer city", "city", "town"]
      )
      state = get_field([
          "ship to state",
          "shipping state",
          "buyer state",
          "state",
          "province",
          "county",
      ])
      zipcode = get_field([
          "ship to zip",
          "ship to postcode",
          "shipping zip",
          "shipping postcode",
          "buyer zip",
          "buyer postcode",
          "postal code",
          "postcode",
          "zip",
      ]).upper()

      # US ZIP code: strip extension after '-' (e.g. 90210-1234 -> 90210)
      if dest_code == "US" and "-" in zipcode:
        zipcode = zipcode.split("-")[0].strip()

      raw_phone = get_field([
          "ship to phone",
          "shipping phone",
          "buyer phone",
          "buyer phone number",
          "phone number",
          "phone",
      ])
      clean_phone = re.sub(r"[^\d]", "", raw_phone)
      if not clean_phone:
        clean_phone = "0000000000"

      
      parcel_total_weight = 0
      parcel_total_declared_val = 0
      order_is_multi = len(item_rows) > 1
      current_order_subpieces = []

      # Loop over EVERY distinct item line in this order
      for lsn, (_, item) in enumerate(item_rows.iterrows(), start=1):
        item_title = str(item.get(title_col, ""))
        num_col = (
            "Item Number" if "Item Number" in item.index else "Item number"
        )
        item_id = str(item.get(num_col, "")).strip()

        qty_col = "Quantity" if "Quantity" in item.index else "Item Quantity"
        try:
          qty = int(float(item.get(qty_col, 1)))
          if qty < 1:
            qty = 1
        except:
          qty = 1

        specs = get_item_specs(item_title)
        if specs["is_multi"] or qty > 1:
          order_is_multi = True

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

        # SubPiece entry
        current_order_subpieces.append({
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
            "SP TAX INVOICE DATE": pd.Timestamp.now().strftime("%Y-%m-%d"),
            "SP INV VALUE PU": unit_declared_val,
            "SP INV VALUE TOTAL": item_total_declared_val,
            "ecommerce_url": "Ebay.com",
            "ecommerce_paytranid": np.nan,
            "ecommerce_sku": item_id,
        })

      # --- SHIP TO ADDRESS FIELDS (PRIORITIZED OVER BUYER ADDRESS) ---
      ship_name = clean_str(
          first_row.get(
              "Ship To Name",
              first_row.get(
                  "Ship to Name",
                  first_row.get(
                      "Recipient Name", first_row.get("Buyer Name", "")
                  ),
              ),
          )
      )
      add1 = clean_str(
          first_row.get(
              "Ship To Address 1",
              first_row.get(
                  "Ship to Address 1",
                  first_row.get(
                      "Shipping Address 1",
                      first_row.get("Buyer Address 1", ""),
                  ),
              ),
          )
      )
      add2 = clean_str(
          first_row.get(
              "Ship To Address 2",
              first_row.get(
                  "Ship to Address 2",
                  first_row.get(
                      "Shipping Address 2",
                      first_row.get("Buyer Address 2", ""),
                  ),
              ),
          )
      )
      city = clean_str(
          first_row.get(
              "Ship To City",
              first_row.get(
                  "Ship to City",
                  first_row.get(
                      "Shipping City", first_row.get("Buyer City", "")
                  ),
              ),
          )
      )
      state = clean_str(
          first_row.get(
              "Ship To State",
              first_row.get(
                  "Ship to State",
                  first_row.get(
                      "Shipping State", first_row.get("Buyer State", "")
                  ),
              ),
          )
      )
      zipcode = clean_str(
          first_row.get(
              "Ship To Zip",
              first_row.get(
                  "Ship to Zip",
                  first_row.get(
                      "Ship To Postcode",
                      first_row.get(
                          "Ship to Postcode",
                          first_row.get(
                              "Shipping Postcode",
                              first_row.get("Buyer Postcode", ""),
                          ),
                      ),
                  ),
              ),
          )
      ).upper()

      raw_phone = clean_str(
          first_row.get(
              "Ship To Phone",
              first_row.get(
                  "Ship to Phone",
                  first_row.get(
                      "Shipping Phone", first_row.get("Buyer Phone", "")
                  ),
              ),
          )
      )
      clean_phone = re.sub(r"[^\d]", "", raw_phone)
      if not clean_phone:
        clean_phone = "0000000000"

      if order_is_multi:
        highlight_articles.append(order_serial)

      # Article Details entry
      article_rows.append({
          "SERIAL NUMBER": order_serial,
          "ARTICLE NUMBER": np.nan,
          "DESTINATION COUNTRY NAME": dest_name,
          "DESTINATION COUNTRY CODE": dest_code,
          "MAIL NATURE TYPE": 11,
          "MAIL TRANSPORT TYPE": "AMS",
          "PHYSICAL WEIGHT": parcel_total_weight,
          "DECLARED VALUE": parcel_total_declared_val,
          "NON DELIVERY INSTRUCTIONS": "P",
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
          "RECEIVER NAME": ship_name,
          "RECEIVER ADD LINE 1": add1,
          "RECEIVER ADD LINE 2": add2 if add2 != "" else np.nan,
          "RECEIVER CITY": city,
          "RECEIVER STATE": state,
          "RECEIVER ZIPCODE": zipcode,
          "RECEIVER MOBILE NO": clean_phone,
          "BULK REF": str(order_id),
          "DROP OFF PINCODE": 136118,
          "PBE TYPE": "PBEIII",
          "PBE FILING": "SELF",
          "INCOTERMS": "DAP",
      })

      subpiece_rows.extend(current_order_subpieces)
      order_serial += 1

    df_art_final = pd.DataFrame(article_rows)
    df_sub_final = pd.DataFrame(subpiece_rows)

    output_buffer = io.BytesIO()
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

    df_art_export = df_art_final.reindex(columns=ref_art_cols)
    df_sub_export = df_sub_final.reindex(columns=ref_sub_cols)

    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
      df_art_export.to_excel(writer, sheet_name="ArticleDetails", index=False)
      df_sub_export.to_excel(writer, sheet_name="SubPieces", index=False)

      try:
        for sheet in ["PickupAddress", "Information"]:
          if sheet in wb_tpl.sheetnames:
            df_extra = pd.read_excel(tpl_path, sheet_name=sheet)
            df_extra.to_excel(writer, sheet_name=sheet, index=False)
      except:
        pass

    # Apply Pink Highlighting to PHYSICAL WEIGHT for multi-packs
    wb_out = openpyxl.load_workbook(output_buffer)
    ws_out_art = wb_out["ArticleDetails"]
    pink_fill = PatternFill(
        start_color="FFB6C1", end_color="FFB6C1", fill_type="solid"
    )

    header_cols = {
        ws_out_art.cell(1, c).value: c
        for c in range(1, ws_out_art.max_column + 1)
    }
    sno_col = header_cols.get("SERIAL NUMBER", 1)
    wt_col = header_cols.get("PHYSICAL WEIGHT", 7)

    for r in range(2, ws_out_art.max_row + 1):
      s_val = ws_out_art.cell(r, sno_col).value
      if s_val in highlight_articles:
        ws_out_art.cell(r, wt_col).fill = pink_fill

    final_buffer = io.BytesIO()
    wb_out.save(final_buffer)

    st.subheader("📋 Output Verification")
    c1, c2 = st.columns(2)
    with c1:
      st.metric("Total Parcels (Articles)", len(df_art_export))
      st.caption("Multi-packs are highlighted in pink in the Excel.")
      st.dataframe(
          df_art_export[[
              "SERIAL NUMBER",
              "RECEIVER NAME",
              "DESTINATION COUNTRY NAME",
              "DESTINATION COUNTRY CODE",
              "PHYSICAL WEIGHT",
              "DECLARED VALUE",
          ]]
      )
    with c2:
      st.metric("Total SubPieces", len(df_sub_export))
      st.dataframe(
          df_sub_export[[
              "SERIAL NUMBER REF ",
              "HS DESCRIPTION",
              "SP COUNT",
              "SP INV CURRENCY CODE",
              "SP INV EXCHANGE RATE",
              "SP INV VALUE TOTAL",
          ]]
      )

    export_date = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    st.download_button(
        label="📥 Download India Post Excel File",
        data=final_buffer.getvalue(),
        file_name=f"IndiaPost_BulkUpload_{export_date}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  except Exception as e:
    st.error(f"Error processing orders: {str(e)}")