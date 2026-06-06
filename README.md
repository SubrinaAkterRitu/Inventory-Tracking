# Inventory Management System — Excel

A fully connected, formula-driven inventory management workbook built in Microsoft Excel. All data flows automatically between sheets — update a product cost, add a transaction, or change a reorder level and every dependent sheet recalculates instantly with no manual input required.

---

## Overview

This workbook manages stock across **3 store locations** (Store A, Store B, Store C) for **20 products** across **6 categories**, supplied by **9 vendors**. It tracks opening stock, sales, and receipts through a transaction log and surfaces live inventory levels, stock values, and reorder signals — all without any macros or VBA.

---

## Sheet architecture

```
Products (master data)
    │
    ├──► Inventory (VLOOKUP ← Products for product details)
    │         │
    │         └── Transactions (SUMIFS → live qty on hand)
    │                   │
    ├──► Reports (SUMIFS ← Inventory for summaries)
    │
    └──► Order (SUMIFS ← Inventory for reorder quantities)
```

### Products
Master data table. The single source of truth for all product attributes.

| Column | Description |
|---|---|
| ProductID | Unique identifier (P001 – P020) |
| Product Name | Display name |
| Category | One of 6 categories |
| Cost/Unit ($) | Unit cost — change here to update all stock values automatically |
| Reorder Level | Minimum qty threshold — change here to update reorder signals automatically |
| Supplier | Supplying vendor name |

**Categories:** Electronics, Lifestyle, Health & Beauty, Apparel, Stationery, Food & Drink

**Suppliers:** NovaBrew Supplies, TechLink Distributors, PureGlow Imports, UrbanCraft Co., GreenEdge Nurseries, PageMaster Supplies, NutriPeak Foods, EcoRoot Imports, ActiveCore Wholesale

---

### Transactions
The movement log. Every stock event is recorded here as a single row.

| Column | Description |
|---|---|
| TransID | Unique transaction ID (T001 – T081) |
| Date | Transaction date (DD/MM/YYYY) |
| ProductID | Links to Products sheet |
| Site | Store A, Store B, or Store C |
| Quantity | Positive = stock in (receipt/opening), negative = stock out (sale) |
| Type | Opening Stock / Receipt / Sale |

**81 transactions** covering opening stock loads, sales, and receipts across all 3 stores.

> **To add a transaction:** append a new row to this sheet. Inventory quantities update automatically via SUMIFS.

---

### Inventory
Live stock view. No values are hardcoded — everything is formula-driven.

| Column | Formula source | Description |
|---|---|---|
| Site | Manual | Store A / B / C |
| ProductID | Manual | Links to Products |
| Product Name | `VLOOKUP` → Products | Auto-populated |
| Supplier | `VLOOKUP` → Products | Auto-populated |
| Cost/Unit ($) | `VLOOKUP` → Products | Updates when Products changes |
| Reorder Level | `VLOOKUP` → Products | Updates when Products changes |
| Qty on Hand | `SUMIFS` → Transactions | Live sum of all movements |
| Stock Value ($) | `Qty × Cost` | Recalculates automatically |
| Reorder? | `IF(Qty ≤ Reorder Level)` | Yes / No flag |
| Order Date | Links to Order sheet | Populated when reorder triggered |

**60 rows** — 20 products × 3 stores.

---

### Reports
Live summary tables — all values pull from Inventory via SUMIFS. No manual data entry.

- **Quantity on hand by store** — product-level qty breakdown across Store A, B, C with grand totals
- **Stock value by store ($)** — same structure for monetary value
- **Store health scorecard** — per-store summary of total qty, total value, items below reorder level, reorder risk %, and a Good / Needs Attention / Critical health status

---

### Order
Reorder planning sheet. Shows how many units each store needs to order to bring stock back to the reorder level.

- **Reorder qty** = `MAX(0, Reorder Level − Qty on Hand)` where `Reorder? = Yes`
- **Product names** pull from Products via VLOOKUP
- **Order date** is set in cell B3 — change it once and all triggered order dates update across Inventory
- Totals row sums units to order by site and overall

---

## How the formulas connect

### Qty on Hand (Inventory → Transactions)
```excel
=SUMIFS(Transactions!$E$3:$E$83,
        Transactions!$C$3:$C$83, [ProductID],
        Transactions!$D$3:$D$83, [Site])
```
Sums all quantity movements for a given product at a given store.

### Product details (Inventory → Products)
```excel
=VLOOKUP(ProductID, Products!$A$3:$F$22, [col], 0)
```
Pulls name, supplier, cost, and reorder level from the Products master.

### Report summaries (Reports → Inventory)
```excel
=SUMIFS(Inventory!$G$3:$G$62,
        Inventory!$B$3:$B$62, [ProductID],
        Inventory!$A$3:$A$62, [Site])
```
Aggregates qty or value by product and store.

### Reorder quantities (Order → Inventory)
```excel
=IF(Reorder?="Yes", ReorderLevel − QtyOnHand, 0)
```
Calculates how many units to order per site.

---

## Making changes

| What you want to do | Where to do it | What auto-updates |
|---|---|---|
| Add a new product | Products sheet, new row | Inventory (add rows manually for each site) |
| Change a product cost | Products → Cost/Unit column | Inventory Stock Value, Reports Stock Value |
| Change a reorder threshold | Products → Reorder Level column | Inventory Reorder?, Reports Scorecard, Order qty |
| Record a sale or receipt | Transactions sheet, new row | Inventory Qty on Hand, Stock Value, Reorder?, Reports, Order |
| Change the order date | Order sheet, cell B3 | All order dates in Inventory |
| Add a new store | Transactions + Inventory (new rows) | Reports and Order (if formulas extended) |

---

## File details

| Property | Value |
|---|---|
| Format | `.xlsx` (Microsoft Excel) |
| Sheets | 5 (Products, Transactions, Inventory, Reports, Order) |
| Products | 20 |
| Stores | 3 (Store A, Store B, Store C) |
| Transactions | 81 |
| Inventory rows | 60 |
| Macros / VBA | None |
| External dependencies | None |
| Theme | Midnight Teal (`#085041`, `#0F6E56`, `#1D9E75`, `#5DCAA5`) |

---

## Color legend

| Color | Hex | Meaning |
|---|---|---|
| Deep teal | `#085041` | Sheet headers and titles |
| Store A | `#0F6E56` | Store A accent |
| Store B | `#1D9E75` | Store B accent |
| Store C | `#5DCAA5` | Store C accent |
| Tint | `#E1F5EE` | Alternating row fill |
| Good | `#0F6E56` | Health status: Good |
| Warning | `#BA7517` | Health status: Needs Attention |
| Critical | `#993C1D` | Health status: Critical |

---

## Getting started

1. Download `Inventory_Management.xlsx`
2. Open in Microsoft Excel (Excel 2016 or later recommended for full formula support)
3. Enable editing if prompted
4. Start with the **Products** sheet to review or update product data
5. Use the **Transactions** sheet to log stock movements
6. View live levels in **Inventory**, summaries in **Reports**, and reorder planning in **Order**

> Google Sheets is partially compatible but SUMIFS across sheets may need adjustment.

---

## Project background

This workbook was designed as a dummy dataset demonstration of a fully formula-connected inventory system — no VBA, no Power Query, no external data sources. It demonstrates how a well-structured Excel workbook can replace simple inventory software for small multi-location retail operations.

---

## License

This project is released for educational and demonstration purposes. The data is entirely fictional — all product names, supplier names, and transaction figures are made up.
