import io
import re
from datetime import datetime
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="India Post Bulk Upload Generator", layout="wide"
)

st.title("📦 Sharmex Global - India Post Label Generator")
st.write(
    "Upload your eBay Awaiting Dispatch CSV to automatically generate India"
    " Post bulk booking templates."
)

# ----------------- SENDER CONFIGURATION -----------------
SENDER_DETAILS = {
    "NAME": "NEELA SHARMA",
    "COMPANY": "SHARMEX GLOBAL",
    "ADD_1": "1037, NEW MODEL TOWN",
    "ADD_2": "PINJORE",
    "ADD_3": "",
    "CITY": "PANCHKULA",
    "STATE": "HR",
    "COUNTRY_NAME": "India",
    "COUNTRY_CODE": "IN",
    "PINCODE": 134102,
    "EMAIL": "sharmexglobal@gmail.com",
    "MOBILE": 8629056095,
    "KYC": "DIXPS8173N",
    "DROP_OFF_PINCODE": 136118,
    "PBE_TYPE": "PBEIII",
    "PBE_FILING": "SELF",
}

EXCHANGE_RATES = {
    "GBP": 128.0,
    "USD": 83.5,
    "EUR": 91.0,
}


def clean_uk_postcode(pc):
  if not pc or pd.isna(pc):
    return ""
  pc = str(pc).strip().upper().replace(" ", "")
  if len(pc) > 4:
    return f"{pc[:-3]} {pc[-3:]}"
  return pc


def clean_phone(phone_val):
  if not phone_val or pd.isna(phone_val):
    return ""
  p = str(phone_val).split(".")[0].strip()
  p = re.sub(r"[^\d]", "", p)
  if p.startswith("44") and len(p) > 10:
    p = p[2:]
  elif p.startswith("1") and len(p) > 10:
    p = p[1:]
  return p[-10:] if len(p) >= 10 else p


def parse_currency_and_val(val_str):
  if not val_str or pd.isna(val_str):
    return "GBP", 0.0, 128.0
  s = str(val_str).strip()
  curr = "GBP"
  if "$" in s:
    curr = "USD"
  elif "€" in s or "EUR" in s:
    curr = "EUR"
  elif "£" in s or "GB" in s:
    curr = "GBP"

  num_match = re.search(r"[\d]+(?:\.\d+)?", s.replace(",", ""))
  amt = float(num_match.group()) if num_match else 0.0
  rate = EXCHANGE_RATES.get(curr, 128.0)
  return curr, amt, rate


# ----------------- FILE LOADERS -----------------
col1, col2 = st.columns(2)

with col1:
  try:
    template_wb = openpyxl.load_workbook("template.xlsx")
    st.success("✅ Loaded local 'template.xlsx' automatically")
  except Exception:
    uploaded_template = st.file_uploader(
        "Upload template.xlsx", type=["xlsx"]
    )
    template_wb = (
        openpyxl.load_workbook(uploaded_template) if uploaded_template else None
    )

with col2:
  try:
    catalog_df = pd.read_excel("product_catalog.xlsx", dtype=str)
    st.success("✅ Loaded local 'product_catalog.xlsx' automatically")
  except Exception:
    uploaded_cat = st.file_uploader(
        "Upload product_catalog.xlsx", type=["xlsx"]
    )
    catalog_df = (
        pd.read_excel(uploaded_cat, dtype=str) if uploaded_cat else None
    )

uploaded_ebay = st.file_uploader(
    "Upload eBay Orders CSV (Awaiting Dispatch)", type=["csv"]
)

if uploaded_ebay and template_wb and catalog_df is not None:
  if st.button("🚀 Process & Generate Labels", type="primary"):
    content = uploaded_ebay.getvalue().decode("utf-8-sig", errors="ignore")
    lines = content.splitlines()

    # Locate the header row safely
    header_idx = None
    for idx, line in enumerate(lines[:10]):
      if "order number" in line.lower():
        header_idx = idx
        break

    if header_idx is None:
      st.error(
          "Could not detect eBay Order Number header. Please check your CSV"
          " file."
      )
      st.stop()

    raw_ebay = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])), dtype=str)
    raw_ebay.columns = [str(c).strip() for c in raw_ebay.columns]

    # Map variations between UK and US eBay export formats
    col_map = {
        "Order Number": "Order number",
        "Item Title": "Item title",
        "Item Number": "Item number",
        "Buyer Name": "Buyer name",
        "Ship To Name": "Post to name",
        "Ship To Address 1": "Post to address 1",
        "Ship To Address 2": "Post to address 2",
        "Ship To City": "Post to city",
        "Ship To State": "Post to county",
        "Ship To Zip": "Post to postcode",
        "Ship To Country": "Post to country",
        "Ship To Phone": "Post to phone",
        "Buyer Address 1": "Post to address 1",
        "Buyer Address 2": "Post to address 2",
        "Buyer City": "Post to city",
        "Buyer State": "Post to county",
        "Buyer Zip": "Post to postcode",
        "Buyer Country": "Post to country",
        "Buyer Phone": "Post to phone",
        "Sold For": "Sold for",
    }
    for old_k, new_k in col_map.items():
      if old_k in raw_ebay.columns and new_k not in raw_ebay.columns:
        raw_ebay[new_k] = raw_ebay[old_k]

    # Filter out empty and footer rows
    valid_orders = raw_ebay[
        raw_ebay["Order number"]
        .astype(str)
        .str.strip()
        .str.contains(r"^\d{2}-\d{5}-\d{5}$", regex=True)
    ].copy()

    # Drop parent bundle lines that lack item titles
    if "Item title" in valid_orders.columns:
      valid_orders = valid_orders[
          valid_orders["Item title"].fillna("").str.strip() != ""
      ].copy()

    today_str = datetime.today().strftime("%Y-%m-%d")

    article_rows = []
    subpiece_rows = []
    serial_no = 1

    # Group by order number to bundle multi-item orders
    for order_id, order_group in valid_orders.groupby("Order number", sort=False):
      first_row = order_group.iloc[0]

      dest_country = str(
          first_row.get("Post to country", "United Kingdom")
      ).strip()
      country_cd = "GB" if "united kingdom" in dest_country.lower() else "US"

      rec_name = str(
          first_row.get("Post to name", first_row.get("Buyer name", ""))
      ).strip().title()
      rec_add1 = str(first_row.get("Post to address 1", "")).strip()
      rec_add2 = str(first_row.get("Post to address 2", "")).strip()
      rec_city = str(first_row.get("Post to city", "")).strip().title()
      rec_state = str(first_row.get("Post to county", "")).strip().title()
      rec_zip = clean_uk_postcode(first_row.get("Post to postcode", ""))
      rec_phone = clean_phone(first_row.get("Post to phone", ""))

      order_weight_total = 0
      order_val_inr_total = 0

      for lsn, (_, item_row) in enumerate(order_group.iterrows(), start=1):
        title = str(item_row.get("Item title", "")).strip()
        item_id = str(item_row.get("Item number", "")).strip()
        qty = int(float(item_row.get("Quantity", 1)))
        sold_for_raw = item_row.get("Sold for", "")

        curr_code, item_price, ex_rate = parse_currency_and_val(sold_for_raw)
        item_inr = int(round(item_price * ex_rate))

        # Lookup in product catalog: try item_id first, then keyword matching
        matched = None
        for _, cat in catalog_df.iterrows():
          cat_id = str(cat.get("item_id", "")).strip()
          cat_kw = str(cat.get("title_keywor", "")).strip().lower()

          if cat_id and cat_id in item_id:
            matched = cat
            break
          elif cat_kw and cat_kw in title.lower():
            matched = cat
            break

        if matched is not None:
          hs_code = str(matched.get("hs_code", "33030090")).split(".")[0]
          cth_code = str(matched.get("cth_code", hs_code)).split(".")[0]
          hs_desc = str(matched.get("hs_description", title[:30]))
          wt_val = int(float(matched.get("weight_grams", 250)))
        else:
          hs_code = "33030090"
          cth_code = "33030090"
          hs_desc = title[:30]
          wt_val = 250

        sp_wt = wt_val * qty
        order_weight_total += sp_wt
        order_val_inr_total += item_inr

        subpiece_rows.append({
            "SERIAL NUMBER REF ": serial_no,
            "HS CODE": hs_code,
            "CTH CODE": cth_code,
            "HS DESCRIPTION": hs_desc,
            "SP UNIT CD": "PC",
            "SP COUNT": qty,
            "SP WEIGHT TOTAL": sp_wt,
            "SP NET WEIGHT": sp_wt,
            "SP ORIGIN COUNTRY CODE": "IN",
            "SP ORIGIN CURRENCY CODE": "INR",
            "SP COMM INVOICE NO": order_id,
            "SP COM INVOICE DATE(DD-MM-YYYY)": today_str,
            "SP INVOICE LSN": lsn,
            "SP INV CURRENCY CODE": curr_code,
            "SP INV EXCHANGE RATE": int(ex_rate),
            "SP ASBL FOB VALUE": item_price,
            "SP ASBL INR VAL": item_inr,
            "SP TAX INVOICE NO": order_id,
            "SP TAX INVOICE DATE": today_str,
            "SP INV VALUE PU": item_inr,
            "SP INV VALUE TOTAL": item_inr,
            "ecommerce_url": "Ebay.com",
            "ecommerce_sku": item_id,
        })

      article_rows.append({
          "SERIAL NUMBER": serial_no,
          "DESTINATION COUNTRY CODE": country_cd,
          "DESTINATION COUNTRY NAME": dest_country,
          "MAIL NATURE TYPE": 11,
          "MAIL TRANSPORT TYPE": "AMS",
          "PHYSICAL WEIGHT": order_weight_total,
          "DECLARED VALUE": order_val_inr_total,
          "NON DELIVERY INSTRUCTIONS": "N",
          "SENDER NAME": SENDER_DETAILS["NAME"],
          "SENDER COMPANY": SENDER_DETAILS["COMPANY"],
          "SENDER ADD LINE 1": SENDER_DETAILS["ADD_1"],
          "SENDER ADD LINE 2": SENDER_DETAILS["ADD_2"],
          "SENDER ADD LINE 3": SENDER_DETAILS["ADD_3"],
          "SENDER CITY": SENDER_DETAILS["CITY"],
          "SENDER STATE": SENDER_DETAILS["STATE"],
          "SENDER COUNTRY NAME": SENDER_DETAILS["COUNTRY_NAME"],
          "SENDER COUNTRY CODE": SENDER_DETAILS["COUNTRY_CODE"],
          "SENDER PINCODE": SENDER_DETAILS["PINCODE"],
          "SENDER EMAILID": SENDER_DETAILS["EMAIL"],
          "SENDER MOBILE": SENDER_DETAILS["MOBILE"],
          "SENDER KYC": SENDER_DETAILS["KYC"],
          "RECEIVER NAME": rec_name,
          "RECEIVER ADD LINE 1": rec_add1,
          "RECEIVER ADD LINE 2": rec_add2,
          "RECEIVER CITY": rec_city,
          "RECEIVER STATE": rec_state,
          "RECEIVER ZIPCODE": rec_zip,
          "RECEIVER EMAILID": SENDER_DETAILS["EMAIL"],
          "RECEIVER MOBILE NO": rec_phone,
          "POD FLAG": False,
          "PICKUP ADDRESS FLAG": False,
          "DROP OFF PINCODE": SENDER_DETAILS["DROP_OFF_PINCODE"],
          "PBE TYPE": SENDER_DETAILS["PBE_TYPE"],
          "PBE FILING": SENDER_DETAILS["PBE_FILING"],
      })

      serial_no += 1

    # Load output into template sheets
    ws_art = template_wb["ArticleDetails"]
    ws_sub = template_wb["SubPieces"]

    # Clear pre-existing data rows (keep headers)
    if ws_art.max_row > 1:
      ws_art.delete_rows(2, ws_art.max_row)
    if ws_sub.max_row > 1:
      ws_sub.delete_rows(2, ws_sub.max_row)

    # Write ArticleDetails
    art_headers = [cell.value for cell in ws_art[1]]
    for r in article_rows:
      row_vals = [r.get(h, None) for h in art_headers]
      ws_art.append(row_vals)

    # Write SubPieces
    sub_headers = [cell.value for cell in ws_sub[1]]
    for r in subpiece_rows:
      row_vals = [r.get(h, None) for h in sub_headers]
      ws_sub.append(row_vals)

    # Export to memory
    output_stream = io.BytesIO()
    template_wb.save(output_stream)
    output_stream.seek(0)

    st.success(
        f"🎉 Successfully prepared {len(article_rows)} consignments with"
        f" {len(subpiece_rows)} product sub-pieces!"
    )
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    st.download_button(
        label="📥 Download India Post Excel File",
        data=output_stream,
        file_name=f"IndiaPost_BulkUpload_{timestamp_str}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )