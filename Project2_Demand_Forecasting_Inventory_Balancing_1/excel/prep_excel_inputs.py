"""
Prepares the exact tables the Excel workbook will be built from, including
computing each SKU's contiguous row range within the weekly sales sheet
(needed so FORECAST.LINEAR formulas can reference a clean train/test split
per SKU without volatile array formulas).
"""
import pandas as pd
import numpy as np
import json

DATA = "/home/claude/project2/data"
SQL_OUT = "/home/claude/project2/sql"
EXCEL_DIR = "/home/claude/project2/excel"

sales = pd.read_csv(f"{DATA}/sales_history_clean.csv")
sku = pd.read_csv(f"{DATA}/sku_master_clean.csv")
diag = pd.read_csv(f"{SQL_OUT}/sku_diagnostics.csv")
inv = pd.read_csv(f"{DATA}/inventory_snapshot_clean.csv")

# --- Weekly aggregated sales, one row per SKU-week, sorted for contiguous ranges ---
weekly = (sales.groupby(["SKU_ID", "Week_Ending"], as_index=False)
                .agg(Qty_Sold=("Qty_Sold", "sum"), Revenue=("Revenue", "sum")))
weekly = weekly.sort_values(["SKU_ID", "Week_Ending"]).reset_index(drop=True)
weekly["Week_Index"] = weekly.groupby("SKU_ID").cumcount() + 1

weekly.to_csv(f"{EXCEL_DIR}/weekly_sales_sorted.csv", index=False)

# --- Row ranges (1-indexed data rows, header will be row 1, so data starts row 2) ---
HEADER_OFFSET = 2  # first data row in the eventual Excel sheet
ranges = {}
row_cursor = HEADER_OFFSET
HOLDOUT_WEEKS = 6
MIN_WEEKS_FOR_FORECAST = 15

for sid, grp in weekly.groupby("SKU_ID", sort=False):
    n = len(grp)
    start_row = row_cursor
    end_row = row_cursor + n - 1
    eligible = n >= MIN_WEEKS_FOR_FORECAST
    train_end_row = end_row - HOLDOUT_WEEKS if eligible else None
    test_start_row = train_end_row + 1 if eligible else None
    ranges[sid] = {
        "start_row": start_row, "end_row": end_row, "n_weeks": n,
        "eligible": bool(eligible),
        "train_start_row": start_row, "train_end_row": train_end_row,
        "test_start_row": test_start_row, "test_end_row": end_row if eligible else None,
    }
    row_cursor = end_row + 1

with open(f"{EXCEL_DIR}/sku_row_ranges.json", "w") as f:
    json.dump(ranges, f, indent=2)

n_eligible = sum(1 for v in ranges.values() if v["eligible"])
print(f"Weekly sales rows (sorted): {len(weekly)}")
print(f"SKUs eligible for statistical forecast (>= {MIN_WEEKS_FOR_FORECAST} weeks): {n_eligible} / {len(ranges)}")
print(f"SKUs too sparse for forecast: {[k for k,v in ranges.items() if not v['eligible']]}")

# --- Monthly inventory, sorted for contiguous ranges (used by AVERAGEIFS mostly, so order matters less,
#     but we sort for readability) ---
inv_sorted = inv.sort_values(["SKU_ID", "Month"]).reset_index(drop=True)
inv_sorted.to_csv(f"{EXCEL_DIR}/inventory_sorted.csv", index=False)

# --- SKU list sorted by revenue descending (drives ABC cumulative-% row order) ---
sku_by_rev = diag.sort_values("Total_Revenue", ascending=False).reset_index(drop=True)
sku_by_rev.to_csv(f"{EXCEL_DIR}/sku_by_revenue.csv", index=False)

print("Prep complete.")
