"""
Builds the Power BI-ready export layer: a clean, flat SKU-level fact/dimension
table (one row per SKU, matching Inventory_Planning_Model / Executive_Summary
exactly) plus a weekly sales fact table for the trend visuals. Column names
are Power-BI-friendly (no spaces, consistent casing) and every number here
was cross-checked against the Excel workbook in the same session.
"""
import pandas as pd

SQL_OUT = "/home/claude/project2/sql"
EXCEL_DIR = "/home/claude/project2/excel"
PBI = "/home/claude/project2/powerbi"

diag = pd.read_csv(f"{SQL_OUT}/sku_diagnostics.csv")
weekly = pd.read_csv(f"{EXCEL_DIR}/weekly_sales_sorted.csv")

pbi_sku = diag.rename(columns={
    "Supplier_Clean": "Supplier",
    "UOM_Clean": "UOM",
}).copy()

pbi_sku = pbi_sku[[
    "SKU_ID", "SKU_Description", "Category", "Supplier", "Warehouse", "UOM",
    "Unit_Cost", "Unit_Price", "Lead_Time_Days",
    "Total_Revenue", "Total_Qty_Sold", "ABC_Class", "XYZ_Class", "ABC_XYZ",
    "Mean_Weekly_Qty", "StDev_Weekly_Qty", "CV",
    "Avg_On_Hand_Qty", "Avg_On_Order_Qty", "Avg_Daily_Usage",
    "Days_of_Supply", "Days_of_Supply_Incl_Pipeline", "Inventory_Status",
    "InStock_Rate", "Inventory_Turnover", "Annual_COGS", "Avg_Inventory_Value",
    "Overstock_Flag", "Stockout_Risk_Flag", "Excess_Risk_SKU",
]]
pbi_sku.to_csv(f"{PBI}/PowerBI_SKU_Inventory_Model.csv", index=False)

pbi_weekly = weekly.merge(diag[["SKU_ID", "Category", "ABC_Class", "XYZ_Class"]], on="SKU_ID")
pbi_weekly.to_csv(f"{PBI}/PowerBI_Weekly_Sales_Fact.csv", index=False)

print(f"PowerBI_SKU_Inventory_Model.csv: {len(pbi_sku)} rows x {len(pbi_sku.columns)} cols")
print(f"PowerBI_Weekly_Sales_Fact.csv: {len(pbi_weekly)} rows x {len(pbi_weekly.columns)} cols")
