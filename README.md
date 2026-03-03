# Apps Script Autofill Toolkit (Google Sheets)

A Google Sheets **Apps Script** toolkit for accelerating structured data entry via:
- header validation
- reference lookups (brands, categories, VAT, tags, age restriction, etc.)
- bulk autofill (fast batch reads/writes)
- conditional formatting for missing/invalid values
- optional “search helpers” and Google search link generation

This repository is published as a **generalized template**. It contains **no proprietary data**.  
You must provide your own reference tables (IDs, mappings, brand list).

---

## Features

- **Custom menus** added on spreadsheet open (`onOpen`)
- **Run All Processing** for fast batch processing
- Individual tools:
  - extract contents/units from product titles
  - resolve brand IDs from a reference list
  - fill category labels + extended metadata from a reference map
  - validate category IDs against a reference ID list
  - generate Google Search links from barcodes and titles
- **Visual feedback** (cell background highlights) for rows needing review
- **Caching** of reference ID list for performance (`CacheService`)

---

## Spreadsheet setup

### 1) Main working sheet (Sheet 1)
Row 1 must contain the headers you use in `HEADERS` inside `Code.gs`.

At minimum, you’ll typically want:
- `title_en`
- `title_local`
- `barcode`
- `category_id`

…and output columns such as:
- `brand_id`
- `category_gr`, `category_en`
- `vat_rate`
- `product_type`
- `tags`
- `age_restricted`, `age_minimum`
- `weight_value`, `weight_unit`
- `contents_value`, `contents_unit`
- etc.

> You can rename columns, as long as you keep `HEADERS` updated accordingly.

### 2) Reference sheets (tabs)
Create these tabs in the same spreadsheet:

- `Ref_Brands_list`  
  Two columns:
  - A: brand name (lowercased match key)
  - B: brand_id

- `Ref_Category_ID_list`  
  One column:
  - A: category_id values

- `Ref_Category_Map_A` and `Ref_Category_Map_B`  
  A mapping table where the first column is `category_id`, and additional columns contain metadata (category labels, VAT, tags, etc.).

> The script supports choosing between mapping tables via the dropdown in cell `A1`.

---

## Install

1. Open your Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Paste the contents of `Code.gs` into the editor.
4. Save.
5. Reload the spreadsheet.

You should now see a custom menu:
- **⚡ Autocomplete Tools**
- **🔗 Link Tools**

---

## Usage

1. Choose the dataset type from cell `A1` (e.g., “Catalog A” / “Catalog B”).
2. Paste or type your product rows in the main sheet.
3. Run:
   - **⚡ Autocomplete Tools → ▶️ Run All Processing**

Optional:
- Use “Convert Barcodes & Titles to Search Links” to turn cells into clickable Google searches.

---

## Notes on responsible use

This toolkit is intended for your own spreadsheets and reference data.
If you use it for web-related workflows (e.g., searching), do so responsibly and respect the policies of any external systems you interact with.

---

## Disclaimer

This project is a generalized template based on real-world spreadsheet automation patterns.  
It includes **no employer data, no proprietary IDs, and no internal reference tables**.

---

## License

MIT
