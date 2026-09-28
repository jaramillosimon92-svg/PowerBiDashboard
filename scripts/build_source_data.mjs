import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const projectRoot = "C:\\Users\\jaram\\OneDrive\\Desktop\\REPO\\PowerBiDashboard";
const outputPath = `${projectRoot}\\data\\retail_sales_inventory.xlsx`;

const products = [
  ["P-1001", "Electronics", "Wireless Headphones", 89, 52],
  ["P-1002", "Electronics", "Bluetooth Speaker", 64, 38],
  ["P-1003", "Electronics", "Webcam", 79, 46],
  ["P-2001", "Home Office", "Ergonomic Chair", 249, 156],
  ["P-2002", "Home Office", "Standing Desk", 399, 252],
  ["P-2003", "Home Office", "Desk Lamp", 45, 23],
  ["P-3001", "Accessories", "Laptop Sleeve", 32, 12],
  ["P-3002", "Accessories", "USB-C Hub", 58, 29],
];

const regions = ["East", "Central", "Prairies", "West"];
const months = [0, 1, 2, 3, 4, 5];
const salesRows = [];
let orderNo = 10001;
for (const month of months) {
  for (let p = 0; p < products.length; p += 1) {
    for (let r = 0; r < regions.length; r += 1) {
      const [productId, category, product, unitPrice, unitCost] = products[p];
      const seasonality = [0, 2, 5, 7, 4, 8][month];
      const regionDemand = [7, 10, 5, 8][r];
      const units = 12 + ((p * 5 + r * 3 + month * 4) % 18) + seasonality + regionDemand;
      const revenue = units * unitPrice;
      const cost = units * unitCost;
      salesRows.push([
        new Date(2026, month, 5 + ((p * 3 + r * 2) % 20)),
        `ORD-${orderNo++}`,
        regions[r],
        productId,
        category,
        product,
        units,
        unitPrice,
        unitCost,
        revenue,
        cost,
        revenue - cost,
      ]);
    }
  }
}

const inventoryRows = [];
for (let p = 0; p < products.length; p += 1) {
  for (let r = 0; r < regions.length; r += 1) {
    const [productId, category, product] = products[p];
    const reorderPoint = 50 + p * 8;
    const unitsOnHand = 32 + ((p * 17 + r * 11) % 125);
    inventoryRows.push([
      new Date(2026, 5, 30),
      productId,
      category,
      product,
      regions[r],
      unitsOnHand,
      reorderPoint,
      unitsOnHand <= reorderPoint ? "Reorder" : "In Stock",
    ]);
  }
}

const workbook = Workbook.create();
const sales = workbook.worksheets.add("SalesData");
const inventory = workbook.worksheets.add("InventoryData");
const dictionary = workbook.worksheets.add("DataDictionary");

sales.getRange("A1:L1").values = [["Order Date", "Order ID", "Region", "Product ID", "Category", "Product", "Units Sold", "Unit Price", "Unit Cost", "Revenue", "Cost", "Profit"]];
sales.getRange(`A2:L${salesRows.length + 1}`).values = salesRows;
sales.tables.add(`A1:L${salesRows.length + 1}`, true, "SalesTable");
sales.getRange(`A2:A${salesRows.length + 1}`).format.numberFormat = "yyyy-mm-dd";
sales.getRange(`G2:G${salesRows.length + 1}`).format.numberFormat = "#,##0";
sales.getRange(`H2:L${salesRows.length + 1}`).format.numberFormat = "$#,##0.00";

inventory.getRange("A1:H1").values = [["Snapshot Date", "Product ID", "Category", "Product", "Region", "Units On Hand", "Reorder Point", "Stock Status"]];
inventory.getRange(`A2:H${inventoryRows.length + 1}`).values = inventoryRows;
inventory.tables.add(`A1:H${inventoryRows.length + 1}`, true, "InventoryTable");
inventory.getRange(`A2:A${inventoryRows.length + 1}`).format.numberFormat = "yyyy-mm-dd";
inventory.getRange(`F2:G${inventoryRows.length + 1}`).format.numberFormat = "#,##0";

dictionary.getRange("A1:C1").values = [["Table", "Field", "Description"]];
dictionary.getRange("A2:C11").values = [
  ["SalesData", "Order Date", "Date of the sales transaction"],
  ["SalesData", "Revenue", "Units Sold multiplied by Unit Price"],
  ["SalesData", "Cost", "Units Sold multiplied by Unit Cost"],
  ["SalesData", "Profit", "Revenue less Cost"],
  ["InventoryData", "Snapshot Date", "Date inventory was measured"],
  ["InventoryData", "Units On Hand", "Current stock quantity by product and region"],
  ["InventoryData", "Reorder Point", "Minimum desired stock quantity"],
  ["InventoryData", "Stock Status", "Reorder when Units On Hand is at or below Reorder Point"],
  ["Project", "Data source", "Synthetic portfolio data created for this Power BI project"],
  ["Project", "Period", "January to June 2026 sales; June 30, 2026 inventory snapshot"],
];
dictionary.tables.add("A1:C11", true, "DictionaryTable");

for (const sheet of [sales, inventory, dictionary]) {
  sheet.showGridLines = false;
  const used = sheet.getUsedRange();
  used.format.font = { name: "Arial", size: 10 };
  used.format.verticalAlignment = "center";
  used.format.autofitColumns();
  used.format.autofitRows();
  sheet.getRange("A1:Z1").format = { fill: "#16324F", font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" } };
  sheet.freezePanes.freezeRows(1);
}
sales.tabColor = "#16324F";
inventory.tabColor = "#2E75B6";
dictionary.tabColor = "#7F8C8D";

workbook.recalculate();
const check = await workbook.inspect({ kind: "table", range: "SalesData!A1:L6", include: "values", tableMaxRows: 6, tableMaxCols: 12 });
console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 30 } });
console.log(errors.ndjson);

const preview = await workbook.render({ sheetName: "SalesData", range: "A1:L16", scale: 1.4, format: "png" });
await fs.mkdir(`${projectRoot}\\data`, { recursive: true });
await fs.writeFile(`${projectRoot}\\data\\sales_data_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
