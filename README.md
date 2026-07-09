# US Regional Sales — SUMIFS / AVERAGEIFS Excel Analysis

A worked solution to an Excel formula exercise: populate a 47-category weekly sales summary from a 7,991-row raw order ledger, using **SUMIFS**, **AVERAGEIFS**, and **VLOOKUP** — no manual data entry, no pivot tables, no macros.

## Files

| File | Description |
|---|---|
| `US_Regional_Sales_Data_IFs_Original.xlsx` | The dataset exactly as given — raw orders plus a blank summary template. |
| `US_Regional_Sales_Data_IFs_Completed.xlsx` | The solved workbook: every summary cell is a live formula, verified against an independent recalculation. |

## The dataset

The workbook has 7 sheets. Three matter for this exercise:

- **Sales Orders Sheet** — 7,991 individual order line items (`OrderNumber`, `OrderDate`, `_ProductID`, `Order Quantity`, `Unit Price`, `Unit Cost`, `Discount Applied`, plus a few pre-built formula columns such as `Order day` = `=TEXT(E2,"ddd")` and `Order value` = `=O2*M2`).
- **Products Sheet** — a 47-row lookup mapping `_ProductID` (numeric key used in the orders table) to `Product Name` (the human-readable category, e.g. "Cookware", "Photo Frames").
- **Product Sales Summary** — the task: a 47-category × (avg price, avg cost, 5 weekdays × 3 metrics) grid, left completely blank in the original file.

The catch: the orders table identifies products only by a numeric `_ProductID`, while the summary is organised by category **name**. SUMIFS/AVERAGEIFS can't filter on a name that isn't present in the orders table, so the first job is bridging that gap.

## Approach

**1. Bridge the ID → category gap with VLOOKUP.**
A helper column (`W`) was added to the Sales Orders Sheet:

```
W2 = VLOOKUP(L2, 'Products Sheet'!$A:$B, 2, FALSE)
```

filled down all 7,991 rows. This turns every order row's numeric `_ProductID` into its category name, so the raw data can now be filtered by the same labels used in the summary.

**2. Populate the summary with SUMIFS / AVERAGEIFS, filtered on category *and* weekday.**
For each of the 47 category rows (5–51) and each weekday column:

```
Average unit price   = AVERAGEIFS('Sales Orders Sheet'!$O:$O, 'Sales Orders Sheet'!$W:$W, C5)
Average unit cost    = AVERAGEIFS('Sales Orders Sheet'!$P:$P, 'Sales Orders Sheet'!$W:$W, C5)
Sales Quantity (Mon) = SUMIFS('Sales Orders Sheet'!$M:$M, 'Sales Orders Sheet'!$W:$W, C5, 'Sales Orders Sheet'!$Q:$Q, F$4)
Sales Value (Mon)    = SUMIFS('Sales Orders Sheet'!$T:$T, 'Sales Orders Sheet'!$W:$W, C5, 'Sales Orders Sheet'!$Q:$Q, K$4)
Avg Discount (Mon)   = AVERAGEIFS('Sales Orders Sheet'!$N:$N, 'Sales Orders Sheet'!$W:$W, C5, 'Sales Orders Sheet'!$Q:$Q, P$4)
```

Every weekday column repeats the same pattern against `Tue`–`Fri`, giving 47 × 17 = 799 live formula cells.

**3. Totals row.** Row 52 sums the quantity/value columns and averages the price/cost/discount columns across all 47 categories, so the sheet self-checks against a grand total.

## Verification

Formulas alone aren't proof — a formula can be logically wrong and still return a plausible-looking number. Every cell was checked two independent ways:

1. **Recalculation.** The completed workbook was opened and recalculated headlessly in LibreOffice Calc (not just re-saved — an actual formula recalculation pass), and the resulting cached values were read back out.
2. **Independent computation.** The same 7,991-row raw order table was loaded into pandas and pivoted independently (`groupby`/`pivot_table` on category × weekday), completely bypassing the Excel formulas.

The two results were compared cell-for-cell:

- Cookware, Monday: quantity = **67**, value = **£85,880.60**, average discount = **10.33%** — exact match.
- Grand total order quantity across all 47 categories, all 5 weekdays: **25,566** units — exact match.
- Grand total sales value: **£58,662,962.20** — exact match.

No discrepancies were found across any category/weekday combination.

## Why SUMIFS/AVERAGEIFS instead of a PivotTable

A PivotTable would answer this faster to build, but it wouldn't fit the given template — the summary layout (fixed category rows, fixed weekday columns, mixed metrics per block) is not something a pivot table renders directly, and the exercise's point is demonstrating conditional-aggregation formula fluency (multi-criteria SUMIFS/AVERAGEIFS, cross-sheet VLOOKUP) rather than reporting-tool usage. The formulas also stay live: if the raw order data changes, every summary cell recalculates automatically.
