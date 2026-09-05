# Pharmaceutical Supply Chain & Inventory Management (Excel)

A self-contained Excel workbook that models inventory management for a pharmaceutical
distributor/pharmacy — from raw SKU and demand data through to Economic Order Quantity (EOQ),
Safety Stock, Reorder Points, ABC classification, batch expiry tracking, and a live KPI dashboard.

No macros, no add-ins — everything is native Excel formulas, so it opens and works in any
version of Excel (2010+) or LibreOffice Calc.

## Why this project

Pharmaceutical inventory has two pressures that most inventory problems don't: products
**expire**, and stockouts can mean a patient doesn't get a critical medicine. This model treats
both issues explicitly — it doesn't just track stock levels, it calculates *how much* to hold,
*when* to reorder, and flags stock that's about to become unsellable.

## What's inside

| Sheet | Type | Purpose |
|---|---|---|
| `Instructions` | Docs | Explains the workbook, every formula, and how to extend it |
| `Dashboard` | Output | KPI cards + 5 live charts summarizing the whole model |
| `Product Master` | Input | 30 SKUs — cost, price, supplier, stock, lead time, ordering/holding cost |
| `Monthly Demand` | Input | 12 months of demand per SKU, with realistic seasonality |
| `Inventory Analysis` | Calculation engine | EOQ, Safety Stock, Reorder Point, ABC class, stock status |
| `Expiry Tracker` | Calculation | Batch-level expiry dates, flags expired / expiring-soon stock |

Only `Product Master`, `Monthly Demand`, and `Expiry Tracker` contain raw input data (blue
text). Every other sheet is 100% formula-driven off those three — change an input and the whole
workbook, including the dashboard charts, recalculates.

## Methods & formulas used

- **Economic Order Quantity (EOQ):** `EOQ = √(2 × Annual Demand × Ordering Cost ÷ Holding Cost per Unit)`
  — the order size that minimizes combined ordering + holding cost.
- **Safety Stock:** `Z-score × StdDev(monthly demand) × √(Lead Time ÷ 30)`
  — a demand-variability buffer sized to a chosen service level (Z is an adjustable input cell).
- **Reorder Point:** `(Average Daily Demand × Lead Time) + Safety Stock`
  — the stock level that should trigger a new purchase order.
- **ABC Analysis:** SKUs ranked by annual consumption value; cumulative value share buckets them
  into A (top ~70% of value), B (next ~20%), C (remaining ~10%) — the Pareto principle applied
  to inventory control.
- **Expiry flagging:** date-driven (`TODAY()`) status — Expired / Expiring Soon (≤90 days) / OK.

### Excel techniques demonstrated

- `INDEX` + `MATCH` — robust cross-sheet lookups (used instead of `VLOOKUP` so the model
  doesn't break if columns or rows are inserted)
- `RANK` + `SUMPRODUCT` — building a cumulative-percentage ranking without a helper sort
- `SUMIF` / `COUNTIF` — dashboard aggregation from raw transactional data
- Conditional formatting rules — automatic red/yellow/green status highlighting
- Native Excel charts (bar, pie, line) driven entirely by formula-fed helper tables
- Data validation, freeze panes, autofilters, and a clear input/formula color convention (blue
  = input, black = formula) for spreadsheet auditability

## How to use it

1. Download `Pharma_Inventory_Management.xlsx` and open it in Excel or LibreOffice Calc.
2. Read the `Instructions` tab first — it documents every formula and column.
3. Replace the sample data in `Product Master`, `Monthly Demand`, and `Expiry Tracker` (blue
   cells) with your own SKU/demand/batch data.
4. Everything downstream — reorder flags, ABC classes, KPIs, charts — recalculates
   automatically.
5. Adjust the **Service Level Z-score** cell on `Inventory Analysis` to make the model more or
   less conservative (1.28 ≈ 90% service level, 1.65 ≈ 95%, 2.33 ≈ 99%).

## Repository structure

```
├── Pharma_Inventory_Management.xlsx   # the workbook
├── README.md                          # this file
└── screenshots/                       # dashboard screenshots for this README
```

## Possible extensions

- Replace the sample data with a real ERP/WMS export
- Add supplier lead-time reliability tracking (actual vs. quoted lead time)
- Add a purchase-order log sheet feeding back into Current Stock
- Port the calculation logic to Python/pandas for automation at scale

## License

This project uses sample/synthetic data for demonstration. Feel free to fork and adapt —
MIT License (add a `LICENSE` file if you want this formalized on GitHub).
