# 型付き読み書き

## 目的

SpreadsheetAsDataでは、Excelテーブルを単なるセル範囲ではなく、列名と行を持つデータ集合として読み書きします。

型付き読み書きは、次の四段階で扱います。

1. Excelテーブルを列名で扱う
2. Excelテーブルを利用者定義型へ対応付けて読み書きする
3. ブック全体をデータオブジェクトへ対応付ける
4. `.xlsx` から型付き読み書きコードを生成する

この文書は、現行C#版で実装済みの型付きAPIと、現在の制約を記録します。

## 1. Excelテーブルを列名で扱う

`Workbook.Tables` からExcelテーブルを取得し、`Table.Rows` でデータ行を列挙します。

```csharp
using Marimo.SpreadSheetAsData;

using var book = Workbook.Open("orders.xlsx");

var table = book.Tables["Orders"];

foreach (var row in table.Rows)
{
    var productName = row["ProductName"].Value;
    var quantity = row["Quantity"].Value;

    Console.WriteLine($"{productName}: {quantity}");
}
```

`TableRow` では、列名、0始まりの列位置、`TableColumn` からセルを取得できます。

```csharp
var firstRow = table.Rows.First();

var byName = firstRow["ProductName"];
var byOrdinal = firstRow[0];
var byColumn = firstRow[table.Columns["ProductName"]];
```

列や行の構造を確認したい場合は、`Table.Columns`、`Table.Range`、`Table.Worksheet` を使用します。
取得したセルの `Value` へ値を設定し、`Workbook.Save()` または `SaveAs()` で保存できます。

## 2. 利用者定義型へ対応付けて読み書きする

`Workbook.ReadTable<T>(string name)` または `Table.Enumerate<T>()` で、Excelテーブルの各データ行を利用者定義型へ対応付けて列挙できます。
取得した `Table<T>` の `Replace()` へ型付き行を渡すと、既存のExcelテーブル行をワークシート上の順序で置き換えます。

```csharp
using Marimo.SpreadSheetAsData;

public sealed class Order
{
    public string ProductName { get; set; } = "";

    public int Quantity { get; set; }

    public double UnitPrice { get; set; }
}

using var book = Workbook.Open("orders.xlsx");

foreach (var order in book.ReadTable<Order>("Orders"))
{
    Console.WriteLine(order.ProductName);
}
```

```csharp
var table = book.ReadTable<Order>("Orders");
var orders = table.ToArray();

orders[0].Quantity = 3;

table.Replace(orders);
book.Save();
```

プロパティ名とExcel列名が異なる場合は、`SpreadSheetNameAttribute` で列名を指定します。

```csharp
public sealed class Order
{
    [SpreadSheetName("商品名")]
    public string ProductName { get; set; } = "";
}
```

読み取り時の基本規則は次のとおりです。

* publicな引数なしコンストラクターが必要
* public setterを持つプロパティを列へ対応付ける
* `SpreadSheetNameAttribute` がある場合は、属性の列名を使用する
* 属性がない場合は、プロパティ名を列名として使用する
* 属性がない読み取り専用プロパティは無視する
* 属性がある読み取り専用プロパティはエラーにする
* 対応するプロパティがないExcel列は無視する
* 同じ列へ複数のプロパティを対応付けた場合はエラーにする
* 指定した列が存在しない場合はエラーにする

書き込み時は、public getterを持つプロパティを同じ列名規則で対応付けます。
属性がない書き込み専用プロパティは無視し、属性があるのにpublic getterがない場合はエラーにします。
`Table<T>.Replace()` は既存行を置き換えるAPIであり、行の追加、挿入、削除は行いません。

現在対応している変換は、`int`、`double`、`string` です。

* 整数値の `double` は `int` へ変換できる
* 小数値を `int` へ変換しようとした場合はエラーにする
* `double` は `double` へ変換できる
* `string` は `string` へ変換できる

マッピングに失敗した場合は `TableMappingException` を投げます。
例外には、可能な範囲でテーブル名、列名、プロパティ名、ワークシート上の行番号、元セル値が設定されます。

## 3. ブック全体をデータオブジェクトへ対応付ける

`Workbook.Read<T>()` は、ブック内の定義名とExcelテーブルをデータオブジェクトのプロパティへ読み込みます。
プロパティを変更して `Workbook.Replace<T>()` へ渡すと、同じ対応規則で書き戻せます。

```csharp
public sealed class OrderBookData
{
    [SpreadSheetName("Orders")]
    public IEnumerable<Order> OrderLines { get; set; } = [];

    [SpreadSheetName("ReportTitle")]
    public string Title { get; set; } = "";
}

using var book = Workbook.Open("orders.xlsx");
var data = book.Read<OrderBookData>();

data.Title = "Updated orders";

book.Replace(data);
book.Save();
```

`SpreadSheetNameAttribute` は、Excelテーブル名、列名、定義名の明示的な対応付けに共通して使用します。
シートローカル定義名では、`WorksheetName` も指定します。
複数セル定義名に対応するプロパティの型は `IEnumerable<IEnumerable<object?>>` です。

## 4. `.xlsx` から型付き読み書きコードを生成する

`Marimo.SpreadSheetAsData.CodeGeneration` では、Excelブックから型付きラッパーのC#ソースコードを生成します。

```csharp
using Marimo.SpreadSheetAsData.CodeGeneration;

var sources = WorkbookWrapperGenerator.GenerateSources(
    "orders.xlsx",
    options =>
    {
        options.Namespace = "MyApp.SpreadSheets";
        options.NameMappings = new()
        {
            ["注文一覧"] = "Orders",
            ["注文一覧.商品名"] = "ProductName"
        };
    });
```

`GenerateSources` はC#ソース文字列の配列を返します。
現在は1つのソースファイルを生成します。

生成コードには、主に次の型とメンバーが含まれます。

* ブック全体を表す `Book` 派生型
* ワークシートごとの `Worksheet` 派生型
* Excelテーブルごとの `Table<T>` 派生型
* Excelテーブル行を表すPOCO
* ブック、シート、テーブル、定義名へアクセスする型付きプロパティ

生成されたTable型は既存の `Table<T>` を継承します。
そのため、生成行型の値変換は手書きPOCOと同じ `Table<T>` のマッピング規則を使用します。

現行実装では、生成されたBook型の引数なしコンストラクターに、生成元Excelファイルのパスを埋め込みます。

```csharp
using Generated;

using var book = new OrdersBook();

foreach (var order in book.Orders)
{
    Console.WriteLine(order.ProductName);
}

var data = book.Read();
data.Orders = data.Orders.ToArray();
data.Orders.First().Quantity = 3;

book.Replace(data);
book.Save();
```

## 名前変換

Excel上のブック名、シート名、テーブル名、列名、定義名は、C#識別子として使える名前へ変換します。

現在の変換では、ASCIIのcamelCaseや区切り記号をPascalCaseへ寄せ、日本語などの非ASCII文字は内部の単語境界を推測せずに扱います。
生成できない名前や、生成後に同じ名前へ衝突する名前は診断対象です。

自動変換だけでは意図した名前にならない場合は、`CodeGenerationOptions.NameMappings` で明示的に対応を指定します。

```csharp
options.NameMappings = new()
{
    ["sales_detail"] = "SalesDetail",
    ["sales_detail.customer_id"] = "CustomerId",
    ["book.main_cell"] = "MainCell",
    ["sales_data.local_cell"] = "LocalCell"
};
```

文脈付きキーは、単純キーより優先されます。
同じExcel名を、テーブル、ブック定義名、シートローカル定義名などの文脈ごとに異なるC#名へ変換できます。

## 診断

`WorkbookWrapperGenerator.GenerateDiagnostics` は、コード生成前に検出できる問題を返します。

現在は、主に次の問題を診断します。

* 有効なC#識別子を生成できないシート名
* Book型内で生成プロパティ名が衝突するシート名
* 同じ行データ型内で生成プロパティ名が衝突する列名

診断で扱っていない不一致は、生成コードの利用時に既存の `Workbook`、`Table`、`Table<T>` のAPIで検出されます。

## 現在の制約

現行のコード生成は、基本的な読み取りと書き戻しに使用するラッパーを生成します。

次はまだ対象外です。

* 複数ソースファイルへの分割生成
* bool、日付、nullable型、decimalなどの型推論
* 数式、書式、構造化参照からの型生成
* 複数テーブル間の関連推測
* 任意のセル範囲からの型生成
* 行の追加、挿入、削除とテーブル範囲の拡張

現行版では、Excelテーブルを独立した型付きデータ集合として読み書きすることに集中します。
