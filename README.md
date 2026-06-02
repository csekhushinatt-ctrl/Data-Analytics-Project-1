# DecodeLabs Data Analytics Project 1

## Project: Data Cleaning & Preparation

This Week 1 project cleans a raw sales order dataset by handling missing values, duplicate IDs, and inconsistent formats across dates, numbers, and text fields.

## Files

- `data/raw_orders.csv` - original dirty dataset
- `src/data_cleaning.py` - cleaning and validation script
- `outputs/cleaned_orders.csv` - cleaned dataset
- `outputs/cleaning_audit.csv` - before/after quality checks
- `outputs/cleaning_log.csv` - row-level cleaning actions
- `outputs/data_cleaning_project_1.xlsx` - submission workbook with raw data, cleaned data, audit checks, and cleaning log

## Cleaning Rules

- Duplicate `Order ID` rows are removed, keeping the first complete record.
- Dates are standardized to `YYYY-MM-DD`.
- Missing dates are imputed with the median valid order date.
- Customer names, cities, categories, and status values are standardized to title case.
- Missing city values are filled as `Unknown`.
- Missing categories are filled as `Uncategorized`.
- Quantity values are converted to positive integers and missing quantities are imputed with the median valid quantity.
- Unit prices are converted to positive numeric values and missing prices are imputed with the median valid unit price.
- Total amount is recalculated as `Quantity * Unit Price`.

## Validation Result

Run:

```powershell
python src/data_cleaning.py
```

The script verifies that the cleaned dataset has:

- `0` duplicate order IDs
- `0` incorrectly formatted dates
- `0` missing critical values

