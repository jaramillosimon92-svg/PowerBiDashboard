let
    Source = Excel.Workbook(
        File.Contents("C:\Users\jaram\OneDrive\Desktop\REPO\PowerBiDashboard\data\retail_sales_inventory.xlsx"),
        null,
        true
    ),
    InventoryTable = Source{[Item="InventoryTable", Kind="Table"]}[Data],
    ChangedTypes = Table.TransformColumnTypes(
        InventoryTable,
        {
            {"Snapshot Date", type datetime},
            {"Product ID", type text},
            {"Category", type text},
            {"Product", type text},
            {"Region", type text},
            {"Units On Hand", Int64.Type},
            {"Reorder Point", Int64.Type},
            {"Stock Status", type text}
        }
    ),
    DateOnly = Table.TransformColumns(ChangedTypes, {{"Snapshot Date", Date.From, type date}})
in
    DateOnly
