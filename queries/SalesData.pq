let
    Source = Excel.Workbook(
        File.Contents("C:\Users\jaram\OneDrive\Desktop\REPO\PowerBiDashboard\data\retail_sales_inventory.xlsx"),
        null,
        true
    ),
    SalesTable = Source{[Item="SalesTable", Kind="Table"]}[Data],
    ChangedTypes = Table.TransformColumnTypes(
        SalesTable,
        {
            {"Order Date", type datetime},
            {"Order ID", type text},
            {"Region", type text},
            {"Product ID", type text},
            {"Category", type text},
            {"Product", type text},
            {"Units Sold", Int64.Type},
            {"Unit Price", type number},
            {"Unit Cost", type number},
            {"Revenue", type number},
            {"Cost", type number},
            {"Profit", type number}
        }
    ),
    DateOnly = Table.TransformColumns(ChangedTypes, {{"Order Date", Date.From, type date}})
in
    DateOnly
