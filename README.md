# SpreadsheetAsData

SpreadsheetAsDataは、Excelをインストールしていない環境でもExcelファイルを扱えるようにするライブラリです。

現在はC#版を再整備中です。
Open XML SDKを内部実装として使いながら、利用側コードからはワークブック、ワークシート、セルをコレクション操作に近い感覚で扱えるAPIを目指しています。

## 現在のC#版でできること

* `.xlsx` ファイルを開く
* Excelテーブルを名前で取得する
* Excelテーブルの列、データ行、セルを取得する
* Excelテーブルの各データ行を、利用者定義型へ対応付けて列挙する
* 定義名、A1形式、左上セルと右下セルの指定でセル範囲を取得する
* ワークシートを名前または位置で取得する
* セルをA1形式、または列番号と行番号で取得する
* セル参照、行番号、列番号を取得する
* 空白、数値、真偽値、共有文字列セルの値を取得する

現在のC#版は、読み取り機能を中心に再整備している段階です。
書き込みは、このライブラリの主要な拡張対象です。
書式、日付、数式、広範なExcel機能への対応は、基本的なデータ読み書きを整えた後の補助機能として扱います。

## 使用例

SpreadsheetAsDataは、Excelファイルを低水準のシート、行、セルの集合として扱うのではなく、まず業務上の表データとして読み取れるAPIを優先します。

### 型付きテーブルとして読む

Excelテーブルの列名とC#のプロパティを対応付けると、各データ行を利用者定義型として列挙できます。
列名とプロパティ名が異なる場合は、`SpreadsheetColumn` 属性でExcelテーブル列名を指定します。

```csharp
using Marimo.SpreadSheetAsData;

public sealed class OrderRow
{
    [SpreadsheetColumn("商品名")]
    public string ProductName { get; set; } = "";

    [SpreadsheetColumn("数量")]
    public int Quantity { get; set; }

    [SpreadsheetColumn("単価")]
    public double UnitPrice { get; set; }
}

using var book = Workbook.Open("orders.xlsx");

foreach (var order in book.ReadTable<OrderRow>("注文一覧"))
{
    Console.WriteLine(
        $"{order.ProductName}: {order.Quantity * order.UnitPrice}");
}
```

列名とプロパティ名が同じ場合は、属性を書かずに読み取れます。

```csharp
using Marimo.SpreadSheetAsData;

public sealed class CustomerRow
{
    public string Name { get; set; } = "";

    public int Age { get; set; }
}

using var book = Workbook.Open("customers.xlsx");

var customers = book.ReadTable<CustomerRow>("Customers");
```

型付きテーブルでは、現在 `int`、`double`、`string` への基本的な変換を扱います。
対応する列がない場合、変換できない値がある場合、同じ列へ複数のプロパティを対応付けた場合は `TableMappingException` で失敗します。

### Excelテーブルを直接読む

型を用意せず、Excelテーブルの構造をそのまま扱うこともできます。
列、データ行、セルの位置情報が必要な場合はこちらを使います。

```csharp
using Marimo.SpreadSheetAsData;

using var book = Workbook.Open("orders.xlsx");

var table = book.Tables["注文一覧"];

Console.WriteLine(table.Name);
Console.WriteLine(table.Worksheet.Name);
Console.WriteLine(table.Range);

foreach (var column in table.Columns)
{
    Console.WriteLine($"{column.Ordinal}: {column.Name}");
}

foreach (var row in table.Rows)
{
    var productName = row["商品名"].Value;
    var quantity = row["数量"].Value;

    Console.WriteLine(
        $"{row.WorksheetRowIndex}: {productName} x {quantity}");
}
```

`TableRow` では、列名、列位置、`TableColumn` からセルを取得できます。

```csharp
var firstRow = table.Rows.First();

var byName = firstRow["商品名"];
var byOrdinal = firstRow[0];
var byColumn = firstRow[table.Columns["商品名"]];
```

### 定義名とセル範囲を読む

ブックまたはワークシートの `Range` から、定義名やA1形式の範囲参照を解決できます。
名前付き範囲として取得した場合、`CellRange.Name` と `ToString()` はその名前を返します。

```csharp
using Marimo.SpreadSheetAsData;

using var book = Workbook.Open("report.xlsx");

var namedRange = book.Range["集計範囲"];

Console.WriteLine(namedRange.Name);
Console.WriteLine(namedRange.TopLeftCell.Value);
Console.WriteLine(namedRange.BottomRightCell.Value);

var sheet = book.Sheets["売上"];
var range = sheet.Range["A1", "C10"];

Console.WriteLine(range.TopLeftCell.Reference);
Console.WriteLine(range.BottomRightCell.Reference);
```

定義名が単一セルを表す場合は、`Cell` から直接取得できます。

```csharp
var total = book.Cell["総合計"].Value;
```

### ワークシートとセルを読む

より低水準の操作として、ワークシートやセルを直接取得できます。

```csharp
using Marimo.SpreadSheetAsData;

using var book = Workbook.Open("Book1.xlsx");

var sheet = book.Sheets["いろいろなデータ"];

var number = sheet.Cells["A1"].Value;
var text = sheet.Cells["A3"].Value;
var cell = sheet.Cells[1, 1];

Console.WriteLine(cell.Reference);
Console.WriteLine(cell.RowIndex);
Console.WriteLine(cell.ColumnIndex);
```

`Cell.Value` は現在、セルの内容に応じて次の値を返します。

* 空白セル: `BlankValue`
* 数値セル: `double`
* 真偽値セル: `bool`
* 共有文字列セル: `string`

## 設計方針

* Open XML SDKの型や要素構造を、公開APIへできるだけ露出させない
* Excelの全機能対応を先回りして目指さない
* 利用側コードの意図が読み取れるAPIを優先する
* v0.1の再整備中は、API互換性より設計の一貫性を優先する
* C#/.NETの新しい安定版機能を積極的に採用する
* 古い.NET環境への対応は、実利用や公開上の必要が明確になった時点で検討する
* 実装、テスト、READMEの内容を矛盾させない
* 変更しやすい小さな単位で機能を追加する

## ドキュメント

* [プロジェクト概要](docs/project-overview.md)
* [設計方針](docs/design.md)
* [初回公開版の範囲](docs/public-release-scope.md)
* [ロードマップ](docs/roadmap.md)
* [型付き読み取り](docs/typed-reading.md)

## ビルドとテスト

C#版は `CSharp/SpreadSheetAsData.sln` に含まれています。
ライブラリ本体とテストプロジェクトは `net10.0` を対象にしています。
現在のコードはC# 14の構文を使用します。

.NET 10 SDKが入っている環境では、次のコマンドでビルドとテストを実行できます。

```powershell
dotnet build .\CSharp\SpreadSheetAsData.sln
dotnet test .\CSharp\SpreadSheetAsData.sln
```

整形と基本的なスタイルチェックは `.editorconfig` に定義しています。

```powershell
dotnet format .\CSharp\SpreadSheetAsData.sln --verify-no-changes --no-restore --severity warn
```

## NuGetパッケージ

C#版は、NuGetパッケージとして公開できるように準備しています。
パッケージIDは `Marimo.SpreadSheetAsData` です。

ローカルでパッケージを生成する場合は、次のコマンドを実行します。

```powershell
dotnet pack .\CSharp\SpreadSheetAsData\SpreadSheetAsData.csproj -c Release
```

生成されたパッケージは、既定では `CSharp\SpreadSheetAsData\bin\Release\` に出力されます。
ローカルで別プロジェクトから確認する場合は、生成先をNuGetソースとして指定します。

```powershell
dotnet add package Marimo.SpreadSheetAsData --version 0.1.0 --source .\CSharp\SpreadSheetAsData\bin\Release
```

NuGet.orgへ公開した後は、通常のNuGetソースから次のように追加できます。

```powershell
dotnet add package Marimo.SpreadSheetAsData
```

## APIドキュメント

C#版のAPIドキュメントは、XMLドキュメントコメントからDocFXで生成します。
生成には、リポジトリに含めているローカル.NETツール設定を使用します。

ローカルでブラウザ表示まで行う場合は、次のスクリプトを使用します。

```powershell
.\scripts\serve-csharp-api-docs.ps1
```

スクリプトは必要な.NETローカルツールを復元し、DocFXでHTMLを生成してからローカルWebサーバーを起動します。
コンソールに表示されたURLをブラウザで開くと、生成されたAPIドキュメントを確認できます。
確認を終えるときは、コマンドを実行しているターミナルで `Ctrl+C` を押してサーバーを停止します。

HTML生成だけを確認する場合は、次のように実行します。

```powershell
.\scripts\serve-csharp-api-docs.ps1 -BuildOnly
```

生成されたHTMLは `docs/api/csharp/_site/` に出力されます。
DocFXが生成する `docs/api/csharp/metadata/` と `docs/api/csharp/_site/` は、再生成できる成果物としてGit管理に含めません。

## 現在の確認状況

直近の再整備では、次を確認しています。

* ビルド: 成功
* テスト: 成功、172件成功
* XMLドキュメント生成: 成功、警告なし
* `dotnet format --verify-no-changes`: 成功

## 制約

NuGetパッケージは公開準備中です。
現時点では、ローカルで生成したパッケージと、リポジトリを取得してC#プロジェクトを直接参照する形で確認しています。

現行C#版には、まだセル値を書き込む公開APIはありません。
過去のRuby版には書き込み機能がありましたが、C#版では再設計しながら追加する予定です。

Ruby版は過去実装です。
現在はC#版を優先して再整備していますが、余裕ができたらRuby版もC#版と同等の機能へ整備する予定です。
