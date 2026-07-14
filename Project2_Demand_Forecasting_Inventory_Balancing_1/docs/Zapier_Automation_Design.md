# Zapier Automation Design — Low-Stock Alert Workflow

**Status:** Designed and documented, not connected live (no production NetSuite/Slack
credentials in this environment) — same approach used for the Airtable design in
Project 3. This documents exactly how the automation would be built and wired up.

## Why this automation

The Inventory_Planning_Model tab recomputes `Inventory_Status` every time the
underlying sales/inventory data refreshes, but nobody re-opens the workbook daily.
The automation closes that gap: the moment a SKU's *days of supply (incl. pipeline)*
crosses below its own lead time, the assigned planner gets pinged — no manual
spreadsheet-checking required.

## Trigger → Filter → Action

1. **Trigger — Google Sheets: New or Updated Row**
   Watches a live Google Sheet mirror of the `Inventory_Planning_Model` tab
   (in production: a scheduled Power Query / Power Automate refresh from
   NetSuite would keep this sheet current every morning).

2. **Filter — Zapier Filter step**
   Only continue if `Inventory_Status` = `Stockout Risk`. This keeps the workflow
   from firing on every refresh — only on rows that actually need attention.

3. **Action 1 — Slack: Send Channel Message**
   Posts to `#inventory-planning` with the SKU, description, category, current
   days of supply, and lead time, so the planner can triage without opening
   the sheet.

4. **Action 2 — Gmail: Send Email** *(parallel action, same trigger)*
   Sends a daily-digest-style email to the SKU's assigned category planner
   (mapped via a lookup table keyed on `Category`) for anyone not on Slack.

5. **Action 3 — Google Sheets: Update Row** *(closes the loop)*
   Writes a timestamp into an `Alert_Sent_At` column so the same SKU doesn't
   re-trigger every refresh cycle until it's acknowledged or its status changes.

## Field mapping

| Zap field | Source column | Notes |
|---|---|---|
| SKU | `Inventory_Planning_Model!A` | |
| Description | `Inventory_Planning_Model!B` | |
| Category | `Inventory_Planning_Model!C` | drives planner routing |
| Days of Supply (incl. pipeline) | `Inventory_Planning_Model!Q` | rounded to 1 decimal in the message |
| Lead Time (days) | `Inventory_Planning_Model!D` | shown alongside days of supply for context |
| Status | `Inventory_Planning_Model!R` | filter condition |

## Why documented instead of live

Standing up the live version needs a NetSuite saved-search connector or a
scheduled export, a Slack workspace with channel-posting permissions, and a
maintained planner-routing table — none of which exist for a fictional
company. Documenting the exact trigger/filter/action chain and field mapping
demonstrates the same automation-design thinking a hiring manager is screening
for, without requiring production credentials this project was never going to
have. See `zapier_workflow_diagram.svg` for the visual flow.
