"""
Generates a simulated NetSuite-style export for a fictional mid-size industrial
manufacturer, "Meridian Manufacturing Co."

Three raw tables, deliberately messy in the ways a real NetSuite saved-search
export tends to be messy:
  1. sku_master.csv           - 40 SKUs, categories, cost/price, supplier, lead time, UOM
  2. sales_history_raw.csv    - 52 weeks x 40 SKUs of sell-through (mixed date formats,
                                 sparse history on newer SKUs, a few duplicate rows)
  3. inventory_snapshot_raw.csv - month-end on-hand / on-order snapshot, 12 months x 40 SKUs
                                 (UOM inconsistencies, a few missing on-order values)

Everything is generated with a fixed seed so re-runs are reproducible, and every
number downstream (SQL, Excel, dashboard) traces back to this file.
"""
import numpy as np
import pandas as pd
import random
from datetime import datetime, timedelta

SEED = 42
rng = np.random.default_rng(SEED)
random.seed(SEED)

OUT = "/home/claude/project2/data"

# ---------------------------------------------------------------------------
# 1. SKU MASTER
# ---------------------------------------------------------------------------
categories = {
    "Fasteners & Hardware":     ("Midwest Fastener Co.",      2.5,  55, 1.75),
    "Bearings & Bushings":      ("Precision Bearing Supply",  4.0,  120, 1.55),
    "Electric Motors & Drives": ("Torque Dynamics LLC",       12.0, 45, 1.90),
    "Sensors & Controls":       ("SenTech Industrial",        6.0,  30, 2.10),
    "Seals & Gaskets":          ("Apex Rubber & Seal",        1.5,  35, 1.65),
    "Pneumatic Components":     ("AirFlow Pneumatics Inc.",   8.0,  40, 1.80),
    "Raw Material - Metal Stock": ("Great Lakes Steel Supply", 20.0, 60, 1.35),
    "Packaging Materials":      ("SecurePack Solutions",      0.8,  20, 2.20),
}
uom_variants = {
    "EA": ["EA", "Each", "each", "ea."],
    "PCS": ["PCS", "pcs", "Pcs"],
    "BOX": ["BOX", "Box", "box"],
    "FT": ["FT", "ft", "Feet"],
}

sku_rows = []
sku_id = 1000
for cat, (base_supplier, base_cost, base_lead, price_mult) in categories.items():
    n_skus = 5
    for i in range(n_skus):
        sid = f"MM-{sku_id}"
        sku_id += 1
        cost = round(base_cost * rng.uniform(0.6, 1.8), 2)
        price = round(cost * price_mult * rng.uniform(0.95, 1.15), 2)
        lead_time = int(max(5, rng.normal(base_lead, base_lead * 0.25)))
        uom_key = rng.choice(list(uom_variants.keys()), p=[0.55, 0.2, 0.15, 0.1])
        uom = rng.choice(uom_variants[uom_key])
        # messiness: supplier name casing varies across the file
        supplier_variant = rng.choice([base_supplier, base_supplier.upper(),
                                        base_supplier.lower()], p=[0.7, 0.15, 0.15])
        warehouse = rng.choice(["WH-OHIO", "WH-TEXAS", "WH-OHIO", "WH-TEXAS", "WH-NEVADA"])
        # ~15% of SKUs are recent launches with sparse history
        is_new = rng.random() < 0.15
        launch_week = rng.integers(30, 50) if is_new else 0
        sku_rows.append({
            "SKU_ID": sid,
            "SKU_Description": f"{cat.split(' - ')[0].split(' & ')[0]} Item {i+1:02d} - {cat}",
            "Category": cat,
            "Supplier": supplier_variant,
            "Unit_Cost": cost,
            "Unit_Price": price,
            "UOM": uom,
            "Lead_Time_Days": lead_time,
            "Warehouse": warehouse,
            "Launch_Week": launch_week,   # 0 = has full 52-week history
        })

sku_df = pd.DataFrame(sku_rows)

# ---------------------------------------------------------------------------
# 2. SALES HISTORY (weekly, 52 weeks trailing, week-ending dates)
# ---------------------------------------------------------------------------
last_week_ending = datetime(2026, 7, 5)
week_dates = [last_week_ending - timedelta(weeks=w) for w in range(51, -1, -1)]  # oldest -> newest

date_formats = ["%m/%d/%Y", "%Y-%m-%d", "%d-%b-%Y", "%m/%d/%y"]

sales_rows = []
sku_weekly_series = {}  # SKU_ID -> list of true weekly qty (pre-messiness), for realistic inventory sizing
for _, sku in sku_df.iterrows():
    # base weekly demand level per SKU, with category-driven scale
    base_demand = rng.uniform(15, 220)
    # demand variability profile: some SKUs are steady (X), some erratic (Z)
    cv_profile = rng.choice(["steady", "moderate", "erratic"], p=[0.35, 0.4, 0.25])
    cv = {"steady": 0.12, "moderate": 0.30, "erratic": 0.65}[cv_profile]
    # mild seasonal wave + slight trend for realism
    trend = rng.uniform(-0.15, 0.35)  # over the year
    seasonal_amp = rng.uniform(0.05, 0.25)

    weekly_true = []
    for wi, wdate in enumerate(week_dates):
        if wdate < (last_week_ending - timedelta(weeks=52 - sku["Launch_Week"])) and sku["Launch_Week"] > 0:
            continue  # no history before launch
        progress = wi / 51
        seasonal = 1 + seasonal_amp * np.sin(2 * np.pi * (wi / 52) * 2)
        level = base_demand * (1 + trend * progress) * seasonal
        qty = max(0, rng.normal(level, level * cv))
        qty = int(round(qty))
        weekly_true.append(qty)
        if qty == 0 and rng.random() < 0.5:
            continue  # stockout/no-sale week, sparse gap (messiness)

        revenue = round(qty * sku["Unit_Price"] * rng.uniform(0.97, 1.03), 2)

        # date format messiness: pick a random format per row
        fmt = random.choice(date_formats)
        date_str = wdate.strftime(fmt)

        sales_rows.append({
            "SKU_ID": sku["SKU_ID"],
            "Order_Date": date_str,
            "Qty_Sold": qty,
            "Revenue": revenue,
            "Warehouse": sku["Warehouse"],
        })

    sku_weekly_series[sku["SKU_ID"]] = weekly_true

sales_df = pd.DataFrame(sales_rows)

# inject ~1.5% duplicate rows (a known NetSuite export quirk when saved searches double-count)
dupe_sample = sales_df.sample(frac=0.015, random_state=SEED)
sales_df = pd.concat([sales_df, dupe_sample], ignore_index=True)
sales_df = sales_df.sample(frac=1.0, random_state=SEED).reset_index(drop=True)  # shuffle rows

# ---------------------------------------------------------------------------
# 3. INVENTORY SNAPSHOT (month-end, 12 months)
# ---------------------------------------------------------------------------
months = pd.date_range("2025-08-01", "2026-07-01", freq="MS")  # 12 month starts
inv_rows = []

# Planners tend to over-buffer volatile SKUs -- but inconsistently, which is
# exactly the kind of gap this project is built to surface. We derive each
# SKU's actual computed CV from its true underlying weekly series (captured
# during generation, before messiness/date-format noise was layered on).
for _, sku in sku_df.iterrows():
    weekly_true = sku_weekly_series.get(sku["SKU_ID"], [20])
    weekly_arr = np.array(weekly_true) if len(weekly_true) else np.array([20])
    avg_weekly = weekly_arr.mean() if weekly_arr.mean() > 0 else 20
    std_weekly = weekly_arr.std()
    cv_actual = (std_weekly / avg_weekly) if avg_weekly else 0

    if cv_actual < 0.30:       # steady
        target_weeks_cover = rng.uniform(2.0, 5.0)
    elif cv_actual < 0.55:     # moderate
        target_weeks_cover = rng.uniform(3.0, 8.0)
    else:                      # erratic -- wide, noisy cover: some overstocked, some still exposed
        target_weeks_cover = rng.uniform(2.0, 15.0)

    # lean SKUs occasionally dip to a genuine stockout; this makes in-stock
    # rate a real, varying metric instead of a flat 100% across the catalog
    stockout_base_prob = 0.35 if target_weeks_cover < 3.0 else (0.12 if target_weeks_cover < 5.0 else 0.02)

    for m in months:
        month_label = m.strftime("%Y-%m")
        if rng.random() < stockout_base_prob:
            on_hand = 0
        else:
            on_hand = max(0, int(round(avg_weekly * target_weeks_cover * rng.uniform(0.7, 1.3))))
        on_order = int(round(avg_weekly * rng.uniform(0, 2))) if rng.random() > 0.1 else np.nan
        uom_variant = rng.choice(uom_variants.get(sku["UOM"] if sku["UOM"] in uom_variants else "EA", ["EA"]))
        inv_rows.append({
            "SKU_ID": sku["SKU_ID"],
            "Month": month_label,
            "On_Hand_Qty": on_hand,
            "On_Order_Qty": on_order,
            "UOM": uom_variant,
            "Warehouse": sku["Warehouse"],
        })

inv_df = pd.DataFrame(inv_rows)

# ---------------------------------------------------------------------------
# SAVE
# ---------------------------------------------------------------------------
sku_export = sku_df.drop(columns=["Launch_Week"])
sku_export.to_csv(f"{OUT}/sku_master.csv", index=False)
sales_df.to_csv(f"{OUT}/sales_history_raw.csv", index=False)
inv_df.to_csv(f"{OUT}/inventory_snapshot_raw.csv", index=False)

print(f"SKU master: {len(sku_export)} rows")
print(f"Sales history: {len(sales_df)} rows  (incl. {len(dupe_sample)} injected duplicates)")
print(f"Inventory snapshot: {len(inv_df)} rows  (missing on-order: {inv_df['On_Order_Qty'].isna().sum()})")
print(f"Date formats present: {sorted(set(date_formats))}")
