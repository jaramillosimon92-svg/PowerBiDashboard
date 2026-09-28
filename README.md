# Retail Sales & Inventory Performance Dashboard

This Power BI portfolio project analyzes sales and inventory performance for a fictional retail business. It is designed to demonstrate data cleaning, modeling, DAX measures, dashboard design, and practical reporting.

## Business questions

- How are sales, profit, and units trending over time?
- Which regions and product categories perform best?
- Which product-region combinations require inventory attention?

## Data source

Import `data/retail_sales_inventory.xlsx` into Power BI Desktop. The workbook contains:

- `SalesData`: January to June 2026 transaction-level sales data.
- `InventoryData`: inventory snapshot as of June 30, 2026.
- `DataDictionary`: field definitions and data notes.

The data is synthetic and was created specifically for this portfolio project.

## Power BI build plan

1. Import `SalesData` and `InventoryData` from the Excel workbook.
2. In Power Query, confirm data types: dates for date fields; whole numbers for units; decimal/currency for sales metrics.
3. Create a relationship from `SalesData[Product ID]` to `InventoryData[Product ID]`. Use a Product table if Power BI warns that the direct relationship is many-to-many.
4. Create these DAX measures:

```DAX
Total Sales = SUM(SalesData[Revenue])
Total Profit = SUM(SalesData[Profit])
Units Sold = SUM(SalesData[Units Sold])
Profit Margin = DIVIDE([Total Profit], [Total Sales])
Inventory Units = SUM(InventoryData[Units On Hand])
Reorder Items = CALCULATE(COUNTROWS(InventoryData), InventoryData[Stock Status] = "Reorder")
```

5. Build one report page with KPI cards for Total Sales, Total Profit, Profit Margin, and Reorder Items.
6. Add a line chart for monthly sales, a bar chart for sales by region, a column chart for profit by category, and a table for inventory items requiring reorder.
7. Add slicers for Region and Category. Use a clean navy, white, and light-gray theme.
