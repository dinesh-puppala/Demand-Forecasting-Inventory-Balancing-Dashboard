# Demand Forecasting & Inventory Balancing Dashboard
### Supply Chain Analytics Portfolio — Project 2 of 3

A multi-SKU inventory health and demand-planning system built for a fictional
mid-size industrial manufacturer, **Meridian Manufacturing Co.**, from a simulated
NetSuite-style export. SQL for classification, Excel for the live planning model,
Power BI for the executive dashboard, Python for a forecasting stretch, and a
documented Zapier automation for the alerting layer.

![Dashboard preview](https://github.com/dinesh-puppala/Demand-Forecasting-Inventory-Balancing-Dashboard/blob/main/Project2_Demand_Forecasting_Inventory_Balancing_1/screenshots/dashboard_preview.png)

---

## The problem this solves

Rapid demand swings are one of the top threats supply chain leaders flag today,
and naive forecasting methods routinely miss by a wide margin — which means
companies end up simultaneously overstocked on the wrong SKUs and stocked out on
the right ones. This project builds the diagnostic and planning layer that catches
that pattern before it shows up as a write-off or a missed order: which SKUs are
actually at risk, how much working capital is tied up in the wrong places, and how
accurate a simple forecasting method really is once you measure it instead of
assuming it.

## What was built

| Layer | Tool | Deliverable |
|---|---|---|
| Data | Python | Simulated 40-SKU, 52-week NetSuite export — deliberately messy (mixed date formats, UOM inconsistencies, duplicate rows, sparse new-SKU history) |
| Diagnostics | SQL (SQLite) | ABC/XYZ classification, inventory turnover, days of supply, in-stock rate |
| Planning model | Excel | Safety stock, reorder point, EOQ, next-period forecast, stockout/overstock flags — **3,176 live formulas, zero errors** |
| Executive dashboard | Power BI-ready export + DAX guide | KPI cards, ABC/XYZ matrix, reorder alerts, 12-week trend |
| Interactive dashboard | HTML/JS | Recruiter-facing visual centerpiece — see below |
| Automation | Zapier (documented design) | Low-stock alert workflow: Sheets → Filter → Slack/Gmail |
| Stretch | Python (Jupyter) | Moving-average forecast benchmarked against the Excel baseline |

## Key findings

- **20 of 40 SKUs (50%)** carry less inventory than their own lead time requires —
  they would stock out before a reorder could arrive under current buffer levels.
- **4 SKUs (10% of the catalog) hold $111,069 — 47.5% of total inventory value —**
  while being both overstocked *and* demand-erratic (XYZ class Z). This is where
  working capital is actually stuck.
- Catalog-wide **inventory turnover is 7.7x/year**; average in-stock rate is 88.1%.
- The Excel linear-trend forecast (`FORECAST()`) produces a **naive MAPE of 60.5%**,
  but that figure is inflated by a handful of holdout weeks with single-digit actual
  demand — a well-known MAPE failure mode. The volume-weighted **WAPE of 28.1%** is
  the more representative accuracy figure, and it's what the dashboard leads with.
- The Python stretch notebook's **4-week moving average actually beats the linear
  trend** on this catalog (25.6% WAPE vs. 28.1%), but only for steady (X-class)
  SKUs — erratic (Z-class) SKUs punish both methods, just differently. See
  `/python` for the full breakdown and the practical routing recommendation.
- 4 SKUs have under 15 weeks of sales history and are explicitly excluded from
  statistical forecasting rather than force-fit — flagged for judgmental forecasting
  instead.

## The dashboard

`dashboard/Meridian_Inventory_Dashboard.html` — open directly in any browser, no
install required. Built around an industrial instrument-panel concept rather than
a generic SaaS template:

- **Radial gauge KPIs** for stockout risk, overstock, in-stock rate, and turnover
- **Andon board** — every SKU as a signal-light tile, grouped by category, colored
  red/amber/green by inventory status (the same red/yellow/green language used
  throughout the Excel model) — a real factory-floor status-board concept repurposed
  as the dashboard's signature visual
- **ABC × XYZ bubble matrix** — revenue vs. demand variability, bubble size = inventory
  value at stake
- Fully interactive: category/warehouse/status/XYZ filters, sortable table, hover
  tooltips, click-to-inspect tiles — all cross-checked against the Excel workbook
  before shipping (see `dashboard/dashboard_data.json` for the exact payload)

## Automation — Zapier low-stock alert (documented design)

![Zapier workflow diagram](https://github.com/dinesh-puppala/Demand-Forecasting-Inventory-Balancing-Dashboard/blob/main/Project2_Demand_Forecasting_Inventory_Balancing_1/diagrams/zapier_workflow_diagram_preview.png)

`docs/Zapier_Automation_Design.md` documents the full trigger → filter → action
chain (Google Sheets → Filter → Slack + Gmail) and field mapping. Designed and
documented rather than deployed live, since this is a portfolio project without
production NetSuite/Slack credentials — the same approach used for Project 3's
Airtable design.

## Repository structure

```
data/          Simulated NetSuite export (raw + cleaned) and the generator script
sql/           SQLite diagnostic layer: ABC/XYZ, turnover, days of supply
excel/         Project2_Demand_Forecasting_Inventory_Model.xlsx + build scripts
powerbi/       Power BI-ready CSVs + DAX_Walkthrough.md
dashboard/     Meridian_Inventory_Dashboard.html (open this one) + data payload
python/        Moving-average forecast notebook (executed, with charts)
docs/          Zapier automation design
diagrams/      Zapier workflow diagram (SVG + PNG preview)
screenshots/   Dashboard preview image used in this README
```

## Methodology

**Data.** 40 SKUs across 8 categories (fasteners, bearings, motors, sensors,
gaskets, pneumatics, raw metal stock, packaging), 52 weeks of sell-through, 12
months of month-end inventory snapshots. Messiness was injected deliberately —
4 date formats, UOM variants, ~1.5% duplicate transaction rows, ~10% missing
on-order quantities, 4 sparse/new SKUs — and resolved in the SQL/Python layer,
documented row-by-row in `excel/Data_Quality_Log` (workbook tab).

**Classification.** ABC by cumulative revenue contribution (80/95/100 cutoffs);
XYZ by coefficient of variation of weekly demand (population stdev via a
SUMPRODUCT formula, not a CSE array — see workbook for why).

**Planning model.** Safety stock = `Z × σ_weekly × √(lead time in weeks)` at a 95%
service level; reorder point = `avg daily usage × lead time + safety stock`; EOQ
from a $75 order cost and 22% annual holding-cost assumption; MOQ-rounded via
`ROUNDUP`. Every assumption is documented, with rationale, in the workbook's
`Assumptions` tab — nothing is a hardcoded magic number without a cell explaining it.

**Inventory status.** Thresholds are SKU-specific, not catalog-wide: a SKU is
"Stockout Risk" if pipeline-adjusted days of supply is under its *own* lead time,
"Overstock" if over 3x its lead time. The separate "Excess-Risk" designation (used
for the $111K finding) uses a catalog-relative 75th-percentile threshold combined
with XYZ = Z — a deliberately different, clearly-named lens on the same data, kept
consistent across Excel, SQL, and the dashboard so the numbers never diverge.

**Forecasting.** `FORECAST()` (Excel's legacy alias for `FORECAST.LINEAR`, since
the newer function isn't supported by this environment's recalculation engine)
trained on all but the most recent 6 weeks per SKU, tested against those 6 —
true walk-forward holdout, no leakage. 36 of 40 SKUs have enough history to be
eligible; the other 4 are flagged, not force-fit.

## Assumptions & limitations

- All data is simulated; Meridian Manufacturing Co. is fictional.
- Fill rate is approximated as "in-stock rate" (share of months with on-hand > 0)
  rather than a true order-level fill rate, since a NetSuite export at this grain
  doesn't carry demand-vs-fulfillment detail — documented explicitly rather than
  quietly assumed.
- Service level (Z = 1.65), order cost ($75), and holding-cost rate (22%) are
  stated assumptions, not fitted values — change them in one place (`Assumptions`
  tab) and every downstream formula recalculates.
- The Zapier and Power BI layers are designed and documented rather than deployed
  live, since this is a portfolio project without production NetSuite/Slack
  credentials — the same approach used for Project 3's Airtable design.
