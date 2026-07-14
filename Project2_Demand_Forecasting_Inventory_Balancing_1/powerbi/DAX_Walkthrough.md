# Power BI Build Guide — Demand Forecasting & Inventory Balancing

This project exports two Power-BI-ready tables. Import both, relate them, and the
measures below reproduce every number on the Executive_Summary tab of the Excel
workbook and the HTML dashboard.

## 1. Import & relate

| File | Grain | Role |
|---|---|---|
| `PowerBI_SKU_Inventory_Model.csv` | 1 row per SKU (40 rows) | Dimension + SKU-level facts |
| `PowerBI_Weekly_Sales_Fact.csv` | 1 row per SKU per week (1,805 rows) | Transaction-level fact table |

In Power BI Desktop: **Get Data → Text/CSV** for both files, then in **Model view**
create a relationship: `PowerBI_Weekly_Sales_Fact[SKU_ID]` (many) → `PowerBI_SKU_Inventory_Model[SKU_ID]` (one).

## 2. Core DAX measures

```dax
Total SKUs = DISTINCTCOUNT(PowerBI_SKU_Inventory_Model[SKU_ID])

Stockout Risk SKUs =
CALCULATE(
    COUNTROWS(PowerBI_SKU_Inventory_Model),
    PowerBI_SKU_Inventory_Model[Inventory_Status] = "Stockout Risk"
)

Overstock SKUs =
CALCULATE(
    COUNTROWS(PowerBI_SKU_Inventory_Model),
    PowerBI_SKU_Inventory_Model[Inventory_Status] = "Overstock"
)

Excess-Risk SKUs =
CALCULATE(
    COUNTROWS(PowerBI_SKU_Inventory_Model),
    PowerBI_SKU_Inventory_Model[Excess_Risk_SKU] = TRUE
)

Total Inventory Value = SUM(PowerBI_SKU_Inventory_Model[Avg_Inventory_Value])

Excess-Risk Inventory Value =
CALCULATE(
    SUM(PowerBI_SKU_Inventory_Model[Avg_Inventory_Value]),
    PowerBI_SKU_Inventory_Model[Excess_Risk_SKU] = TRUE
)

Excess-Risk Value % = DIVIDE([Excess-Risk Inventory Value], [Total Inventory Value])

Catalog Inventory Turnover =
DIVIDE(
    SUMX(PowerBI_SKU_Inventory_Model, PowerBI_SKU_Inventory_Model[Total_Qty_Sold] * PowerBI_SKU_Inventory_Model[Unit_Cost]),
    [Total Inventory Value]
)

Avg In-Stock Rate = AVERAGE(PowerBI_SKU_Inventory_Model[InStock_Rate])
```

## 3. Recommended visuals

| Visual | Fields | Notes |
|---|---|---|
| **KPI cards** (top row) | `[Total SKUs]`, `[Stockout Risk SKUs]`, `[Overstock SKUs]`, `[Excess-Risk SKUs]` | Conditional color: red if Stockout Risk SKUs > 15, matching the Excel red/yellow/green convention |
| **ABC/XYZ matrix** | Scatter: X = `CV`, Y = `Total_Revenue`, size = `Avg_Inventory_Value`, color = `ABC_XYZ` | Reproduces the Executive_Summary matrix as an interactive bubble chart |
| **Reorder alert table** | Table: `SKU_ID`, `SKU_Description`, `Category`, `Inventory_Status`, `Days_of_Supply`, `Lead_Time_Days`, `Avg_Inventory_Value` | Filter to `Inventory_Status = "Stockout Risk"` for the planner's action list; slicer on `Category` and `Warehouse` |
| **12-week demand trend** | Line chart: `Week_Ending` (last 12 distinct values) vs `SUM(Qty_Sold)`, split by `Category` from `PowerBI_Weekly_Sales_Fact` | Use a **Top N filter** or a relative-date slicer (last 12 weeks) on `Week_Ending` |
| **Excess-risk callout** | Card + table filtered to `Excess_Risk_SKU = TRUE` | Mirrors the Excel Executive_Summary callout list |

## 4. Slicers

Add slicers on `Category`, `Warehouse`, `ABC_Class`, and `XYZ_Class` so a planner can
drill from "all 40 SKUs" down to "just the erratic-demand Class-A items in the
Nevada warehouse" in two clicks.

## 5. Data refresh

Both CSVs are static exports of a point-in-time simulated dataset. In a live
NetSuite environment, `PowerBI_Weekly_Sales_Fact` would be a scheduled query
against the sales-order saved search, and `PowerBI_SKU_Inventory_Model` would be
recomputed via the same SQL layer documented in `/sql/build_sql_layer.py` — the
DAX measures above don't change either way.
