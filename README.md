# Inventory SKU Merger

Small Python web app to merge inventory reports for Furniture of America and Butler Specialty Company.

## Supported upload files

- `.csv`
- `.xlsx`
- `.xls`

## Asset template

- Base export file: `assets/inventory_export_1.csv`
- Only rows where `Location` is `Furniture of America` are updated
- `On hand (new)` is filled from merged SKU totals
- If a Furniture of America SKU is not found in the merged upload data, `On hand (new)` is set to `0`
- Butler template: `assets/merged_inventory_export.csv`
- Butler source mapping: `Product` → `SKU`, `Available` → `On hand (new)`
- Butler output filename: `Butler_inventory_{day}_{month}_{year}.csv` (for example, `Butler_inventory_1_jul_2026.csv`)

## Run

```powershell
python app.py
```

Then open `http://127.0.0.1:8000`

## Desktop EXE

Run `build_desktop.ps1` in PowerShell to create `dist\GooddeggInventory.exe`.
Double-clicking the EXE opens the app in its own desktop window (no separate browser is needed).
Keep the `assets` folder beside the EXE only if you want to replace the bundled asset template.

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
- Use the `Butler Inventory` menu to upload one Butler specialty product inventory CSV and download the mapped export
# Gooddegg_management
