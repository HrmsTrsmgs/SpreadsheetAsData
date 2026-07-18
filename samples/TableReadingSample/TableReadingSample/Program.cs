using Marimo.SpreadSheetAsData;

var workbookPath = Path.Combine(
    AppContext.BaseDirectory,
    "SampleData",
    "orders.xlsx");

using var book = Workbook.Open(workbookPath);

Console.WriteLine("Typed table");
Console.WriteLine("-----------");

foreach (var order in book.ReadTable<OrderRow>("注文一覧"))
{
    Console.WriteLine(
        $"{order.ProductName}: {order.Quantity} x {order.UnitPrice:#,0} = {order.TotalPrice:#,0}");
}

Console.WriteLine();
Console.WriteLine("Raw table");
Console.WriteLine("---------");

var table = book.Tables["注文一覧"];

Console.WriteLine($"{table.Name} ({table.Worksheet.Name}!{table.Range})");

foreach (var row in table.Rows)
{
    Console.WriteLine(
        $"{row.WorksheetRowIndex}: {row["商品名"].Value} / {row["数量"].Value}");
}

public sealed class OrderRow
{
    [SpreadsheetColumn("商品名")]
    public string ProductName { get; set; } = "";

    [SpreadsheetColumn("数量")]
    public int Quantity { get; set; }

    [SpreadsheetColumn("単価")]
    public double UnitPrice { get; set; }

    public double TotalPrice => Quantity * UnitPrice;
}
