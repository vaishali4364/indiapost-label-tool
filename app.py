import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
from datetime import datetime
import os

st.set_page_config(page_title="India Post Label Generator", layout="wide", page_icon="📦")

# --- STATIC SENDER CONFIGURATION ---
STATIC_SENDER = {
    'SENDER NAME': 'NEELA SHARMA',
    'SENDER COMPANY': 'SHARMEX GLOBAL',
    'SENDER ADD LINE 1': '1037, NEW MODEL TOWN',
    'SENDER ADD LINE 2': 'PINJORE',
    'SENDER ADD LINE 3': '',
    'SENDER CITY': 'PANCHKULA',
    'SENDER STATE': 'HR',
    'SENDER COUNTRY NAME': 'India',
    'SENDER COUNTRY CODE': 'IN',
    'SENDER PINCODE': 134102,
    'SENDER EMAILID': 'sharmexglobal@gmail.com',
    'SENDER MOBILE': 8629056095,
    'SENDER ALT CONTACT': '',
    'SENDER KYC': 'DIXPS8173N',
    'DROP OFF PINCODE': 136118,
    'MAIL NATURE TYPE': 11,
    'MAIL TRANSPORT TYPE': 'AMS',
    'NON DELIVERY INSTRUCTIONS': 'N',
    'PBE TYPE': 'PBEIII',
    'PBE FILING': 'SELF',
    'POD FLAG': False,
    'PICKUP ADDRESS FLAG': False,
}

def clean_phone(val):
    if pd.isna(val):
        return ""
    digits = re.sub(r'\D', '', str(val))
    if digits.startswith('44') and len(digits) > 10:
        digits = digits[2:]
    if digits.startswith('0'):
        digits = digits[1:]
    return digits

def clean_price(val):
    if pd.isna(val):
        return 0.0
    val_clean = re.sub(r'[^\d.]', '', str(val))
    return float(val_clean) if val_clean else 0.0

# --- UI HEADER ---
st.title("📦 India Post Label Generator")
st.caption("Upload your eBay Order CSV to automatically generate India Post 'ArticleDetails' and 'SubPieces'.")

col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("1. Upload Files")
    ebay_file = st.file_uploader("Upload eBay Orders Report (CSV)", type=["csv"])
    
    # Template selection: auto-load local template.xlsx if available
    local_template_exists = os.path.exists("template.xlsx")
    if local_template_exists:
        st.success("✅ Loaded local 'template.xlsx' automatically")
        use_custom_template = st.checkbox("Upload a different template.xlsx")
        template_file = st.file_uploader("Upload Template (.xlsx)", type=["xlsx"]) if use_custom_template else "template.xlsx"
    else:
        template_file = st.file_uploader("Upload India Post Template (.xlsx)", type=["xlsx"])

with col_right:
    st.subheader("2. Settings")
    exchange_rate = st.number_input("Exchange Rate (GBP to INR)", value=128.0, step=1.0)
    
    # Catalog auto-load
    local_catalog_exists = os.path.exists("product_catalog.xlsx")
    if local_catalog_exists:
        st.success("✅ Loaded local 'product_catalog.xlsx' automatically")
        catalog_df = pd.read_excel("product_catalog.xlsx")
    else:
        st.warning("⚠️ No local 'product_catalog.xlsx' found.")
        uploaded_cat = st.file_uploader("Upload Product Catalog (.xlsx)", type=["xlsx"])
        catalog_df = pd.read_excel(uploaded_cat) if uploaded_cat else None

def lookup_catalog(item_id, title):
    if catalog_df is None or catalog_df.empty:
        return {}
    if 'item_id' in catalog_df.columns:
        match = catalog_df[catalog_df['item_id'].astype(str) == str(item_id)]
        if not match.empty:
            return match.iloc[0].to_dict()
    if 'title_keyword' in catalog_df.columns and pd.notna(title):
        for _, row in catalog_df.iterrows():
            kw = str(row.get('title_keyword', '')).strip().lower()
            if kw and kw in str(title).lower():
                return row.to_dict()
    return {}

# --- PROCESS BUTTON ---
st.markdown("---")
if ebay_file and template_file:
    if st.button("🚀 Process & Generate Labels", type="primary"):
        raw_ebay = pd.read_csv(ebay_file)
        valid_orders = raw_ebay.dropna(subset=['Order number', 'Item title']).copy()
        valid_orders = valid_orders[~valid_orders['Sales record number'].astype(str).str.contains('record', na=False)]

        if valid_orders.empty:
            st.error("No valid order rows found in the uploaded eBay CSV.")
        else:
            article_rows = []
            subpiece_rows = []

            serial_no = 1
            for order_no, group in valid_orders.groupby('Order number', sort=False):
                first = group.iloc[0]
                order_weight = 0
                order_inr_total = 0

                raw_date = first.get('Paid on date') or first.get('Sale date')
                try:
                    inv_date = datetime.strptime(str(raw_date).strip(), '%d-%b-%y').strftime('%Y-%m-%d')
                except Exception:
                    inv_date = datetime.now().strftime('%Y-%m-%d')

                # SubPieces rows
                for lsn, (_, item) in enumerate(group.iterrows(), start=1):
                    item_id = str(int(item['Item number'])) if pd.notna(item['Item number']) else ''
                    title = item['Item title']
                    qty = int(item['Quantity']) if pd.notna(item['Quantity']) else 1
                    fob_val = clean_price(item['Sold for'])
                    inr_val = round(fob_val * exchange_rate)

                    matched = lookup_catalog(item_id, title)
                    hs_code = matched.get('hs_code', '')
                    cth_code = matched.get('cth_code', hs_code)
                    hs_desc = matched.get('hs_description', '')
                    unit_weight = clean_price(matched.get('weight_grams', 0))
                    total_weight = round(unit_weight * qty) if unit_weight else ''

                    if total_weight != '':
                        order_weight += total_weight
                    order_inr_total += inr_val

                    subpiece_rows.append({
                        'SERIAL NUMBER REF ': serial_no,
                        'HS CODE': hs_code,
                        'CTH CODE': cth_code,
                        'HS DESCRIPTION': hs_desc,
                        'SP UNIT CD': 'PC',
                        'SP COUNT': qty,
                        'SP WEIGHT TOTAL': total_weight,
                        'SP NET WEIGHT': total_weight,
                        'SP ORIGIN COUNTRY CODE': 'IN',
                        'SP ORIGIN CURRENCY CODE': 'INR',
                        'SP COMM INVOICE NO': order_no,
                        'SP COM INVOICE DATE(DD-MM-YYYY)': inv_date,
                        'SP INVOICE LSN': lsn,
                        'SP INV CURRENCY CODE': 'GBP',
                        'SP INV EXCHANGE RATE': exchange_rate,
                        'SP ASBL FOB VALUE': fob_val,
                        'SP ASBL INR VAL': inr_val,
                        'SP TAX INVOICE NO': order_no,
                        'SP TAX INVOICE DATE': inv_date,
                        'SP INV VALUE PU': inr_val,
                        'SP INV VALUE TOTAL': inr_val,
                        'ecommerce_url': 'Ebay.com',
                        'ecommerce_sku': item_id,
                    })

                # ArticleDetails row
                article_rows.append({
                    'SERIAL NUMBER': serial_no,
                    'ARTICLE NUMBER': '',
                    'DESTINATION COUNTRY CODE': 'GB',
                    'DESTINATION COUNTRY NAME': first.get('Post to country', 'United Kingdom'),
                    'MAIL NATURE TYPE': STATIC_SENDER['MAIL NATURE TYPE'],
                    'MAIL TRANSPORT TYPE': STATIC_SENDER['MAIL TRANSPORT TYPE'],
                    'PHYSICAL WEIGHT': order_weight if order_weight > 0 else '',
                    'DECLARED VALUE': order_inr_total,
                    'NON DELIVERY INSTRUCTIONS': STATIC_SENDER['NON DELIVERY INSTRUCTIONS'],
                    'SENDER NAME': STATIC_SENDER['SENDER NAME'],
                    'SENDER COMPANY': STATIC_SENDER['SENDER COMPANY'],
                    'SENDER ADD LINE 1': STATIC_SENDER['SENDER ADD LINE 1'],
                    'SENDER ADD LINE 2': STATIC_SENDER['SENDER ADD LINE 2'],
                    'SENDER ADD LINE 3': STATIC_SENDER['SENDER ADD LINE 3'],
                    'SENDER CITY': STATIC_SENDER['SENDER CITY'],
                    'SENDER STATE': STATIC_SENDER['SENDER STATE'],
                    'SENDER COUNTRY NAME': STATIC_SENDER['SENDER COUNTRY NAME'],
                    'SENDER COUNTRY CODE': STATIC_SENDER['SENDER COUNTRY CODE'],
                    'SENDER PINCODE': STATIC_SENDER['SENDER PINCODE'],
                    'SENDER EMAILID': STATIC_SENDER['SENDER EMAILID'],
                    'SENDER MOBILE': STATIC_SENDER['SENDER MOBILE'],
                    'SENDER ALT CONTACT': STATIC_SENDER['SENDER ALT CONTACT'],
                    'SENDER KYC': STATIC_SENDER['SENDER KYC'],
                    'SENDER IOSS': '',
                    'RECEIVER NAME': first.get('Post to name', ''),
                    'RECEIVER COMPANY': '',
                    'RECEIVER ADD LINE 1': first.get('Post to address 1', ''),
                    'RECEIVER ADD LINE 2': first.get('Post to address 2', '') if pd.notna(first.get('Post to address 2')) else '',
                    'RECEIVER ADD LINE 3': '',
                    'RECEIVER CITY': first.get('Post to city', ''),
                    'RECEIVER STATE': first.get('Post to county', ''),
                    'RECEIVER ZIPCODE': first.get('Post to postcode', ''),
                    'RECEIVER EMAILID': STATIC_SENDER['SENDER EMAILID'],
                    'RECEIVER MOBILE NO': clean_phone(first.get('Post to phone', '')),
                    'RECEIVER ALT CONTACT': '',
                    'RECEIVER KYC': '',
                    'RECEIVER TAX': '',
                    'INSURANCE TYPE': '',
                    'INSURED VALUE': '',
                    'POD FLAG': STATIC_SENDER['POD FLAG'],
                    'BULK REF': '',
                    'PICKUP ADDRESS FLAG': STATIC_SENDER['PICKUP ADDRESS FLAG'],
                    'DROP OFF PINCODE': STATIC_SENDER['DROP OFF PINCODE'],
                    'PBE TYPE': STATIC_SENDER['PBE TYPE'],
                    'PBE FILING': STATIC_SENDER['PBE FILING']
                })
                serial_no += 1

            df_art = pd.DataFrame(article_rows)
            df_sub = pd.DataFrame(subpiece_rows)

            # Display Preview Tabs
            tab1, tab2 = st.tabs(["📄 ArticleDetails (Preview)", "📦 SubPieces (Preview)"])
            with tab1:
                st.dataframe(df_art)
            with tab2:
                st.dataframe(df_sub)

            # Write into template workbook
            wb = openpyxl.load_workbook(template_file)
            for sheet_name, data in [('ArticleDetails', article_rows), ('SubPieces', subpiece_rows)]:
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    header_cols = [cell.value for cell in ws[1]]
                    ws.delete_rows(2, ws.max_row)
                    for r in data:
                        ws.append([r.get(col, '') for col in header_cols])

            out_buf = BytesIO()
            wb.save(out_buf)
            out_buf.seek(0)

            st.success("✅ Conversion complete!")
            st.download_button(
                label="⬇️ Download Completed India Post Excel File",
                data=out_buf,
                file_name=f"IndiaPost_BulkUpload_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("👆 Please upload your eBay CSV file on the left to begin.")