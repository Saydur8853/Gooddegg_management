# Inventory SKU Merger

Small Python web app to upload two or more inventory files, merge SKU quantities, and update the asset export file for `Furniture of America`.

## Supported upload files

- `.csv`
- `.xlsx`
- `.xls`

## Asset template

- Base export file: `assets/inventory_export_1.csv`
- Only rows where `Location` is `Furniture of America` are updated
- `On hand (new)` is filled from merged SKU totals
- If a Furniture of America SKU is not found in the merged upload data, `On hand (new)` is set to `0`

## Run

```powershell
python app.py
```

Then open `http://127.0.0.1:8000`

## Shopify orders

- Open `http://127.0.0.1:8000/orders`
- The app reads Shopify domain/client credentials from:
  - local `.env`
  - or sibling project `../Gooddegg_order_management/.env`
- Click `Connect Shopify` once to complete OAuth
- After success, the Shopify access token is saved in this project's local `.env`
- Click `Load latest orders` to see the latest 20 Shopify orders

## What it does

- Upload two or more inventory files
- Detect SKU and quantity columns automatically
- Sum matching SKU quantities across all uploaded files
- Update `On hand (new)` in `assets/inventory_export_1.csv` for `Furniture of America` rows only
- Download a new CSV file named like `inventory_export_updated_20260715_163000.csv`
# Gooddegg_management
