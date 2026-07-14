"""
Loads the raw simulated NetSuite export into SQLite, builds cleaned staging
tables (parsed dates, standardized UOM/supplier casing, de-duplicated sales
rows), then runs the SQL diagnostic layer:

  - ABC classification (revenue contribution)
  - XYZ classification (demand variability / coefficient of variation)
  - Inventory turnover & days of supply
  - In-stock rate (fill-rate proxy)
  - "Dangerous" SKUs: simultaneously overstocked AND high-variability

Every number in the Excel workbook, Power BI export, and dashboard traces
back to these queries.
"""
import sqlite3
import pandas as pd
import numpy as np

DATA = "/home/claude/project2/data"
SQL_OUT = "/home/claude/project2/sql"
DB_PATH = f"{SQL_OUT}/meridian_inventory.db"

# ---------------------------------------------------------------------------
# LOAD + CLEAN (Python does the parsing; SQL does the analysis)
# ---------------------------------------------------------------------------
sku = pd.read_csv(f"{DATA}/sku_master.csv")
sales = pd.read_csv(f"{DATA}/sales_history_raw.csv")
inv = pd.read_csv(f"{DATA}/inventory_snapshot_raw.csv")

# clean SKU master: standardize supplier casing (title case) and UOM to a canonical code
uom_map = {
    "EA": "EA", "Each": "EA", "each": "EA", "ea.": "EA",
    "PCS": "PCS", "pcs": "PCS", "Pcs": "PCS",
    "BOX": "BOX", "Box": "BOX", "box": "BOX",
    "FT": "FT", "ft": "FT", "Feet": "FT",
}
sku["Supplier_Clean"] = sku["Supplier"].str.strip().str.title()
sku["UOM_Clean"] = sku["UOM"].map(uom_map)

# clean sales: parse mixed date formats -> ISO, drop exact duplicate rows
def parse_mixed_date(s):
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d-%b-%Y", "%m/%d/%y"):
        try:
            return pd.to_datetime(s, format=fmt)
        except ValueError:
            continue
    return pd.NaT

sales["Order_Date_Clean"] = sales["Order_Date"].apply(parse_mixed_date)
before = len(sales)
sales = sales.drop_duplicates(subset=["SKU_ID", "Order_Date_Clean", "Qty_Sold", "Revenue", "Warehouse"])
deduped = before - len(sales)
sales["Week_Ending"] = sales["Order_Date_Clean"].dt.date.astype(str)

# clean inventory: standardize UOM, fill missing on-order with 0 (flagged)
inv["UOM_Clean"] = inv["UOM"].map(uom_map)
inv["On_Order_Missing_Flag"] = inv["On_Order_Qty"].isna().astype(int)
inv["On_Order_Qty_Clean"] = inv["On_Order_Qty"].fillna(0)

print(f"Cleaned sales: {len(sales)} rows (removed {deduped} exact duplicates)")
print(f"Unparseable dates: {sales['Order_Date_Clean'].isna().sum()}")
print(f"UOM values collapsed from raw variants to: {sorted(sku['UOM_Clean'].unique())}")

# ---------------------------------------------------------------------------
# LOAD INTO SQLITE
# ---------------------------------------------------------------------------
conn = sqlite3.connect(DB_PATH)
sku.to_sql("stg_sku_master", conn, if_exists="replace", index=False)
sales.to_sql("stg_sales_history", conn, if_exists="replace", index=False)
inv.to_sql("stg_inventory_snapshot", conn, if_exists="replace", index=False)
conn.commit()

# ---------------------------------------------------------------------------
# SQL DIAGNOSTIC QUERIES
# ---------------------------------------------------------------------------

# 1) Weekly demand stats per SKU (mean, stdev, CV) -- feeds XYZ classification
q_demand_stats = """
WITH weekly AS (
    SELECT SKU_ID, Week_Ending, SUM(Qty_Sold) AS Qty_Sold
    FROM stg_sales_history
    GROUP BY SKU_ID, Week_Ending
),
stats AS (
    SELECT
        SKU_ID,
        COUNT(*) AS Weeks_With_Sales,
        AVG(Qty_Sold) AS Mean_Weekly_Qty,
        SQRT(AVG(Qty_Sold*Qty_Sold) - AVG(Qty_Sold)*AVG(Qty_Sold)) AS StDev_Weekly_Qty
    FROM weekly
    GROUP BY SKU_ID
)
SELECT
    SKU_ID, Weeks_With_Sales, Mean_Weekly_Qty, StDev_Weekly_Qty,
    CASE WHEN Mean_Weekly_Qty > 0 THEN StDev_Weekly_Qty / Mean_Weekly_Qty ELSE NULL END AS CV
FROM stats
"""
demand_stats = pd.read_sql(q_demand_stats, conn)

# 2) Revenue per SKU -> ABC classification (rank by revenue, cumulative % of total)
q_revenue = """
SELECT SKU_ID, SUM(Revenue) AS Total_Revenue, SUM(Qty_Sold) AS Total_Qty_Sold
FROM stg_sales_history
GROUP BY SKU_ID
ORDER BY Total_Revenue DESC
"""
revenue = pd.read_sql(q_revenue, conn)
revenue["Cum_Revenue"] = revenue["Total_Revenue"].cumsum()
grand_total = revenue["Total_Revenue"].sum()
revenue["Cum_Pct"] = revenue["Cum_Revenue"] / grand_total
revenue["ABC_Class"] = np.where(revenue["Cum_Pct"] <= 0.80, "A",
                          np.where(revenue["Cum_Pct"] <= 0.95, "B", "C"))

# 3) XYZ classification from CV (thresholds set from the actual generated
#    variability profile: steady ~0.12, moderate ~0.30, erratic ~0.65)
merged = demand_stats.merge(revenue, on="SKU_ID")
def xyz_class(cv):
    if pd.isna(cv):
        return "Z"  # insufficient/erratic history treated conservatively
    if cv <= 0.30:
        return "X"
    elif cv <= 0.55:
        return "Y"
    else:
        return "Z"
merged["XYZ_Class"] = merged["CV"].apply(xyz_class)
merged["ABC_XYZ"] = merged["ABC_Class"] + merged["XYZ_Class"]

# 4) Inventory turnover, days of supply, in-stock rate (fill-rate proxy)
q_inv = """
SELECT
    SKU_ID,
    AVG(On_Hand_Qty) AS Avg_On_Hand_Qty,
    AVG(On_Order_Qty_Clean) AS Avg_On_Order_Qty,
    SUM(CASE WHEN On_Hand_Qty > 0 THEN 1 ELSE 0 END) * 1.0 / COUNT(*) AS InStock_Rate,
    COUNT(*) AS Months_Tracked
FROM stg_inventory_snapshot
GROUP BY SKU_ID
"""
inv_stats = pd.read_sql(q_inv, conn)

full = merged.merge(inv_stats, on="SKU_ID").merge(
    sku[["SKU_ID", "SKU_Description", "Category", "Supplier_Clean", "Unit_Cost",
         "Unit_Price", "UOM_Clean", "Lead_Time_Days", "Warehouse"]], on="SKU_ID")

full["Avg_Daily_Usage"] = full["Total_Qty_Sold"] / 364.0  # 52 weeks
full["Days_of_Supply"] = np.where(full["Avg_Daily_Usage"] > 0,
                                   full["Avg_On_Hand_Qty"] / full["Avg_Daily_Usage"], np.nan)
# pipeline-adjusted coverage: on-hand PLUS what's already on order, since a
# SKU with strong incoming supply isn't really at risk even if on-hand is low
full["Days_of_Supply_Incl_Pipeline"] = np.where(
    full["Avg_Daily_Usage"] > 0,
    (full["Avg_On_Hand_Qty"] + full["Avg_On_Order_Qty"]) / full["Avg_Daily_Usage"], np.nan)
full["Annual_COGS"] = full["Total_Qty_Sold"] * full["Unit_Cost"]
full["Avg_Inventory_Value"] = full["Avg_On_Hand_Qty"] * full["Unit_Cost"]
full["Inventory_Turnover"] = np.where(full["Avg_Inventory_Value"] > 0,
                                       full["Annual_COGS"] / full["Avg_Inventory_Value"], np.nan)

# 5) Per-SKU stockout risk / overstock / healthy status, thresholds tied to
#    each SKU's OWN lead time (not a flat catalog-wide number):
#      Stockout Risk : days of supply < lead time  -> won't survive to next replenishment
#      Overstock     : days of supply > 3x lead time -> holding far more buffer than needed
#      Healthy       : in between
full["Inventory_Status"] = np.select(
    [full["Days_of_Supply_Incl_Pipeline"] < full["Lead_Time_Days"],
     full["Days_of_Supply_Incl_Pipeline"] > 3 * full["Lead_Time_Days"]],
    ["Stockout Risk", "Overstock"],
    default="Healthy",
)
full["Overstock_Flag"] = full["Inventory_Status"] == "Overstock"
full["Stockout_Risk_Flag"] = full["Inventory_Status"] == "Stockout Risk"

# 6) "Dangerous" SKUs: overstocked (75th-pctile-based, catalog-wide) AND
#    high-variability (Z class) -- the ones burning working capital on
#    buffer stock for items whose demand can't be reliably predicted anyway
dos_75 = full["Days_of_Supply"].quantile(0.75)
overstock_threshold = max(60, dos_75)
full["Excess_Risk_SKU"] = (full["Days_of_Supply"] > overstock_threshold) & (full["XYZ_Class"] == "Z")

full = full.sort_values("Total_Revenue", ascending=False).reset_index(drop=True)

# ---------------------------------------------------------------------------
# SAVE OUTPUTS (feed Excel + Power BI + dashboard)
# ---------------------------------------------------------------------------
full.to_sql("sku_diagnostics", conn, if_exists="replace", index=False)
full.to_csv(f"{SQL_OUT}/sku_diagnostics.csv", index=False)

# also write cleaned staging tables out as CSVs, useful for the Excel raw tabs
sales.to_csv(f"{DATA}/sales_history_clean.csv", index=False)
inv.to_csv(f"{DATA}/inventory_snapshot_clean.csv", index=False)
sku.to_csv(f"{DATA}/sku_master_clean.csv", index=False)

conn.commit()
conn.close()

# ---------------------------------------------------------------------------
# QUICK NARRATIVE FINDINGS (real, computed -- not invented)
# ---------------------------------------------------------------------------
n_sku = len(full)
n_excess_risk = int(full["Excess_Risk_SKU"].sum())
excess_risk_value = full.loc[full["Excess_Risk_SKU"], "Avg_Inventory_Value"].sum()
total_inv_value = full["Avg_Inventory_Value"].sum()
avg_instock = full["InStock_Rate"].mean()
abc_counts = full["ABC_Class"].value_counts().to_dict()
xyz_counts = full["XYZ_Class"].value_counts().to_dict()
overstock_threshold_used = overstock_threshold

print("\n--- DIAGNOSTIC FINDINGS ---")
print(f"Total SKUs analyzed: {n_sku}")
print(f"ABC split: {abc_counts}")
print(f"XYZ split: {xyz_counts}")
print(f"Overstock threshold (days of supply): {overstock_threshold_used:.1f}")
print(f"Excess-risk SKUs (overstocked + high variability): {n_excess_risk} "
      f"({n_excess_risk/n_sku:.1%} of catalog)")
print(f"Inventory value tied up in excess-risk SKUs: ${excess_risk_value:,.0f} "
      f"of ${total_inv_value:,.0f} total ({excess_risk_value/total_inv_value:.1%})")
print(f"Average in-stock rate across catalog: {avg_instock:.1%}")
print(f"Average inventory turnover: {full['Inventory_Turnover'].mean():.2f}x/year")
