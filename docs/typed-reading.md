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
取得したセルの `Value` へ値を設定し、ファイルパスから開いた場合は `Workbook.Save()`、別ファイルへ出力する場合は `SaveAs(path)` で保存できます。
`Save()` は正常終了時点で元ファイルへの保存を完了します。Streamから開いた場合の `Save()` は `NotSupportedException` になります。
Close/Disposeは保存を行いません。Streamだけで編集結果の出力まで完結するAPIは、現時点では提供していません。

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

### テーブル行の値変換

現在対応しているプロパティ型は、`int`、`double`、`bool`、`string` と、`int?`、`double?`、`bool?` です。

| プロパティ型 | 読み込めるセル値 | 空白セルを読み込んだ値 |
| --- | --- | --- |
| `int` | 整数値の `double` | `0` |
| `double` | `double` | `0d` |
| `bool` | `bool` | `false` |
| `string` | `string` | `""` |
| `int?` | 整数値の `double` | `null` |
| `double?` | `double` | `null` |
| `bool?` | `bool` | `null` |

小数値を `int` または `int?` へ変換しようとした場合はエラーになります。
`string?` の注釈によって変換規則を切り替えることはなく、文字列プロパティへ空白を読み込むと空文字列になります。

`Table<T>.Replace()` では、`int` はExcelの数値へ、`double`、`bool`、文字列はそれぞれ対応するセル値へ書き込みます。
nullable値型に値がある場合も、対応する非nullable型と同じ規則です。
プロパティ値が `null` または空文字列 `""` の場合は空白セルにし、スペースを含む文字列 `" "` はそのまま書き込みます。

空白を非nullable型へ読み込むと、元が空白だったという区別は失われます。
例えば空白を `int` の `0` として読み、そのまま書き戻すと数値の `0` になります。数値や真偽値で空白を保持したい場合は、nullable型を使用してください。

これはテーブル行のマッピング規則です。低水準の `Cell.Value` は空白を `BlankValue` として返し、`null` の代入で空白にします。
`Cell.Value` へ直接 `""` を代入した場合は空文字列を持つ文字列セルとなり、型付き `Replace()` の空白への変換とは異なります。

マッピングに失敗した場合は `TableMappingException` を投げます。
例外には、可能な範囲でテーブル名、列名、プロパティ名、ワークシート上の行番号、元セル値が設定されます。

## 3. ブック全体をデータオブジェクトへ対応付ける

`Workbook.Read<T>()` は、ブック内の定義名とExcelテーブルをデータオブジェクトのプロパティへ読み込みます。
プロパティを変更して `Workbook.Replace<T>()` へ渡すと、同じ対応規則で書き戻せます。

この対応付けは、生成されたData型だけでなく、手書きのクラスでも利用できます。
属性がない場合は、元のExcel名を `ToCSharpIdentifier()` でC#識別子へ変換して、プロパティ名と照合します。
例えば、Excelテーブル `orders` と単一セル定義名 `report_title` は、次のプロパティへ属性なしで対応します。

```csharp
public sealed class OrderBookData
{
    public IEnumerable<Order> Orders { get; set; } = [];

    public string ReportTitle { get; set; } = "";
}

using var book = Workbook.Open("orders.xlsx");
var data = book.Read<OrderBookData>();

data.ReportTitle = "Updated orders";

book.Replace(data);
book.Save();
```

別のプロパティ名を使いたい場合は、元のExcel名を属性で明示します。属性の指定は自動対応より優先されます。

```csharp
public sealed class OrderBookData
{
    [SpreadSheetName("orders")]
    public IEnumerable<Order> OrderLines { get; set; } = [];

    [SpreadSheetName("report_title")]
    public string Title { get; set; } = "";
}
```

定義名の自動対応では、ブックスコープとシートローカルの両方を検索します。
例えば、シートローカルの `cell_name` も、変換後の名前に一致する定義名がブック全体で一つなら `CellName` へ対応します。
同じ名前へ変換される定義名が複数ある場合は一意に決まらないため失敗します。シートローカル定義名を明示する場合は、`[SpreadSheetName("cell_name", WorksheetName = "Sheet2")]` のように指定します。

`SpreadSheetNameAttribute` は、Excelテーブル名、列名、定義名の明示的な対応付けに共通して使用します。
ここで説明した自動変換は、ブックのデータオブジェクトと定義名・テーブル名の対応規則です。行データ型 `Order` の列プロパティは、前節の列名規則に従います。
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
生成Bookの `Read()` と `Replace(Data)` でData内のテーブルを扱う場合も、同じ規則を使用します。

### テーブル列の型推論

行データ型のプロパティ型は、生成時のExcelテーブルの列値から判定します。
空白以外の値が一つ以上ある列について、対応する型は次のとおりです。

| 列の値 | 空白を含まない場合 | 空白を含む場合 |
| --- | --- | --- |
| 整数値だけの数値 | `int` | `int?` |
| 小数値を含む数値 | `double` | `double?` |
| 真偽値 | `bool` | `bool?` |
| 文字列 | `string` | `string` |

空白を含む文字列列も `string` として生成し、読み取り時は空白を空文字列へ変換します。
データ行があり、空白だけを持つ列は `string` のプロパティとして生成します。
データ行がないテーブルは、列定義があっても列プロパティを持たない空の行データクラスとして生成します。
日付や `decimal` の型推論は未対応です。

### 生成Bookの利用

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
* 日付や `decimal` の型推論
* 数式、書式、構造化参照からの型生成
* 複数テーブル間の関連推測
* 任意のセル範囲からの型生成
* 行の追加、挿入、削除とテーブル範囲の拡張

現行版では、Excelテーブルを独立した型付きデータ集合として読み書きすることに集中します。
