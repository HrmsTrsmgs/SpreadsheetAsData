# SpreadsheetAsData

SpreadsheetAsDataは、Excelをインストールしていない環境でもExcelファイルを扱えるようにするライブラリです。

現在はC#版を再整備中です。
Open XML SDKを内部実装として使いながら、利用側コードからはワークブック、ワークシート、セルをコレクション操作に近い感覚で扱えるAPIを目指しています。

## まずできること

SpreadsheetAsDataは、Excelファイルを低水準のOpen XML要素としてではなく、業務で使う表データとして読み取るためのAPIを優先しています。

* Excelテーブルを、型付きのC#オブジェクトとして列挙する
* Excelファイルから、ブック、シート、テーブル、行データを表すC#コードを生成する
* Visual Studioでは、Excelファイルのビルドアクションを `SpreadsheetAsData` にするだけで生成コードを利用する
* 型を用意せず、Excelテーブル、列、行、セルを直接読む
* 定義名、A1形式、シート名、セル位置からセルや範囲を取得する
* Excelをインストールしていない環境で `.xlsx` を読む

現在のC#版は、読み取り機能を中心に再整備している段階です。
書き込みは、このライブラリの主要な拡張対象です。
書式、日付、数式、広範なExcel機能への対応は、基本的なデータ読み書きを整えた後の補助機能として扱います。

## 最初のチュートリアル

このチュートリアルでは、リポジトリに含めている `SampleData/sales_report.xlsx` から型付き読み取りコードを生成し、`Program.cs` からExcelテーブルを読み取ります。
サンプルデータには、`sales_summary` シートと `sales_detail` Excelテーブルが含まれています。

### Visual Studioから使う

1. Visual Studioでコンソールアプリの新規プロジェクトを作成する
2. ［NuGet パッケージの管理］から `Marimo.SpreadSheetAsData` をインストールする
3. ソリューション エクスプローラーで、プロジェクトへ `SampleData/sales_report.xlsx` を追加する
4. 追加した `sales_report.xlsx` を選択してプロパティを開く
5. ［ビルド アクション］を `SpreadsheetAsData` へ変更する
6. プロジェクトをビルドする
7. `Program.cs` に次のコードを書く

```csharp
using SpreadsheetTutorial;

using var book = new SalesReportBook();

foreach (var sale in book.SalesDetail)
{
    Console.WriteLine(
        $"{sale.ProductName}: {sale.Quantity} x {sale.UnitPrice}");
}
```

`SpreadsheetTutorial` は、新規作成したプロジェクトの既定名前空間です。
別のプロジェクト名で作成した場合は、そのプロジェクトの名前空間を `using` に指定してください。
`SalesReportBook`、`SalesDetail`、`ProductName` などは、`sales_report.xlsx` のファイル名、Excelテーブル名、列名から生成されます。
生成された `.g.cs` は `sales_report.xlsx` の隣へ出力され、同じビルドの `Compile` に自動で追加されます。
このファイルは生成物なので編集せず、通常はGit管理にも含めません。

### CLIから同じことを行う

Visual Studioを使わない場合は、同じ設定をプロジェクトファイルへ直接書けます。

```powershell
dotnet new console -n SpreadsheetTutorial
cd .\SpreadsheetTutorial
dotnet add package Marimo.SpreadSheetAsData
mkdir SampleData
$repository = "C:\path\to\SpreadsheetAsData"
Copy-Item "$repository\SampleData\sales_report.xlsx" .\SampleData\sales_report.xlsx
```

`SpreadsheetTutorial.csproj` に、対象Excelファイルを `SpreadsheetAsData` 項目として追加します。

```xml
<ItemGroup>
  <SpreadsheetAsData Include="SampleData\sales_report.xlsx" />
</ItemGroup>
```

`Program.cs` には、Visual Studioの場合と同じコードを書きます。

```csharp
using SpreadsheetTutorial;

using var book = new SalesReportBook();

foreach (var sale in book.SalesDetail)
{
    Console.WriteLine(
        $"{sale.ProductName}: {sale.Quantity} x {sale.UnitPrice}");
}
```

ビルドまたは実行すると、Excelファイルから型付き読み取りコードが生成されます。

```powershell
dotnet run
```

## よく使う読み取り例

SpreadsheetAsDataは、Excelファイルを低水準のシート、行、セルの集合として扱うのではなく、まず業務上の表データとして読み取れるAPIを優先します。

### 生成された型付きコードで読む

NuGetパッケージを参照しているVisual Studioプロジェクトでは、Excelファイルのビルドアクションから型付き読み取りコードを生成できます。
たとえば `orders.xlsx` を `SpreadsheetAsData` ビルドアクションにすると、ブック、ワークシート、定義名、Excelテーブル、行データを表す型を利用できます。

```csharp
using MyApp;

using var book = new OrdersBook();

foreach (var order in book.Orders)
{
    Console.WriteLine(order.ProductName);
}
```

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

### コード生成の詳しい設定

NuGetパッケージを参照しているVisual Studioプロジェクトでは、Excelファイルのビルドアクションから型付き読み取りコードを生成できます。
通常の `None` や `Content` として追加したExcelファイルは、コード生成対象になりません。
対象にするファイルだけ、ビルドアクションを `SpreadsheetAsData` へ変更してください。

生成ファイルはExcelファイルの隣に、`*.SpreadsheetAsData.g.cs` という名前で出力されます。
Visual Studioでは元Excelファイルに紐づく生成コードとして確認できます。
生成ファイルは手で編集せず、通常はGit管理に含めません。
このリポジトリでは `*.SpreadsheetAsData.g.cs` を `.gitignore` に登録しています。

Visual Studio以外で明示的に設定する場合は、プロジェクトファイルへ次の項目を追加します。

```xml
<ItemGroup>
  <SpreadsheetAsData Include="Schemas\注文.xlsx" />
</ItemGroup>
```

識別子名を調整したい場合は、Excelファイルごとに任意のJSON辞書を置けます。
辞書ファイル名は、Excelファイルの拡張子を除いた名前に `.spreadsheetasdata.json` を付けます。

```text
SampleProject/
├─ SampleProject.csproj
├─ 注文.spreadsheetasdata.json
└─ Schemas/
   └─ 注文.xlsx
```

Excelファイルごとに個別の辞書を使う場合は、Excelファイルと同じディレクトリへ置きます。

```text
SampleProject/
└─ Schemas/
   ├─ 注文.xlsx
   └─ 注文.spreadsheetasdata.json
```

同じ名前の辞書が両方にある場合は、Excelファイルと同じディレクトリの辞書を優先します。
辞書ファイルはプロジェクトへ追加しなくても、ビルド時にディスク上から検出します。
辞書には、既定変換から変更したい名称だけを書きます。
辞書にない名称は既定規則で変換されます。

```json
{
  "注文一覧": "Orders",
  "注文一覧.商品名": "ProductName"
}
```

文脈付きキーを使うと、同じ元名でもテーブルやシートごとに別の生成名を指定できます。
異なるディレクトリに同名Excelファイルがあり、別々の辞書が必要な場合は、それぞれのExcelファイルの隣へ辞書を置いてください。

コードから直接生成する場合は、`Marimo.SpreadSheetAsData.CodeGeneration` の `WorkbookWrapperGenerator.GenerateSources` を使用します。
このAPIはC#ソース文字列の配列を返します。
MSBuild連携も同じ生成処理と診断処理を使用します。

現在の生成コードは、生成元ExcelファイルのパスをBook型の引数なしコンストラクターに埋め込みます。
生成前に検出できる名前衝突や無効名は `WorkbookWrapperGenerator.GenerateDiagnostics` で確認できます。
詳しい規則は [型付き読み取り](docs/typed-reading.md) を参照してください。

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
* v0.xの再整備中は、API互換性より設計の一貫性を優先する
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

C#版は `CSharp/SpreadSheetAsData.slnx` に含まれています。
ライブラリ本体とテストプロジェクトは `net10.0` を対象にしています。
現在のコードはC# 14の構文を使用します。

このリポジトリのPowerShellスクリプトは、PowerShell 7以降の `pwsh` で実行します。
Windows PowerShell 5.1の `powershell.exe` とPowerShell 7の `pwsh.exe` は同じマシンに共存できますが、このリポジトリでは `pwsh` を明示して使います。
開発環境の前提は次のコマンドで確認できます。

```powershell
pwsh -NoProfile -File .\scripts\Test-DevelopmentEnvironment.ps1
```

初回セットアップでは、次のスクリプトを実行します。

```powershell
pwsh -NoProfile -File .\scripts\Initialize-DevelopmentEnvironment.ps1
```

このスクリプトは、開発環境の前提を確認したうえで、.NETローカルツールとC#ソリューションのNuGetパッケージを復元します。
PowerShell 7、.NET 10 SDK、Gitなどの導入自体は行いません。

.NET 10 SDKが入っている環境では、次のコマンドでビルドとテストを実行できます。

```powershell
dotnet build .\CSharp\SpreadSheetAsData.slnx
dotnet test .\CSharp\SpreadSheetAsData.slnx
```

整形と基本的なスタイルチェックは `.editorconfig` に定義しています。

```powershell
dotnet format .\CSharp\SpreadSheetAsData.slnx --verify-no-changes --no-restore --severity warn
```

## NuGetパッケージ

C#版は、NuGetパッケージとして公開できるように準備しています。
推奨パッケージIDは `Marimo.SpreadSheetAsData` です。
この短い名前のパッケージは、実行時ライブラリ、コード生成API、Visual Studio/MSBuild連携をまとめる全部入りパッケージです。

内部の責務は、次のパッケージに分けています。

* `Marimo.SpreadSheetAsData.Core`: `Workbook`、`Worksheet`、`Table` などの実行時ライブラリ
* `Marimo.SpreadSheetAsData.CodeGeneration`: `.xlsx` から型付き読み取りコードを生成するAPI
* `Marimo.SpreadSheetAsData.Build`: Visual StudioとMSBuildからコード生成を起動するビルドタスク

通常の利用者は `Marimo.SpreadSheetAsData` を参照してください。
依存を絞りたい場合だけ、用途に応じて個別パッケージを参照します。
コード生成タスクは .NET 10 / MSBuild 18 以降の .NET TaskHost を前提にしています。

ローカルでパッケージを生成する場合は、次のコマンドを実行します。

```powershell
dotnet pack .\CSharp\SpreadSheetAsData.slnx -c Release -o .\artifacts\nupkg
```

生成されたパッケージは `artifacts\nupkg\` に出力されます。
NuGet.orgへ公開する前にローカルで別プロジェクトから確認する場合は、検証先プロジェクトに `PackageReference` を追加し、復元時にローカルパッケージ出力先とNuGet.orgをNuGetソースとして指定します。

```xml
<PackageReference Include="Marimo.SpreadSheetAsData" Version="0.2.5" />
```

```powershell
dotnet restore .\YourProject.csproj --source .\artifacts\nupkg --source "https://api.nuget.org/v3/index.json"
```

NuGet.orgへ公開した後は、通常のNuGetソースから次のように追加できます。

```powershell
dotnet add package Marimo.SpreadSheetAsData
```

## サンプル

NuGetパッケージとして参照する利用者向けサンプルは、`samples/TableReadingSample/` にあります。

このサンプルは、リポジトリ内のプロダクトコードを `ProjectReference` では参照せず、外部利用者と同じように `PackageReference` で `Marimo.SpreadSheetAsData` を参照します。

NuGet.orgへ公開する前に動かす場合は、先にローカルパッケージを生成し、サンプルの復元時にその生成先をNuGetソースとして指定します。
詳しい手順は [samples/TableReadingSample/README.md](samples/TableReadingSample/README.md) を参照してください。

## APIドキュメント

C#版のAPIドキュメントは、XMLドキュメントコメントからDocFXで生成します。
生成には、リポジトリに含めているローカル.NETツール設定を使用します。

ローカルでブラウザ表示まで行う場合は、次のスクリプトを使用します。

```powershell
pwsh -NoProfile -File .\scripts\serve-csharp-api-docs.ps1
```

スクリプトは必要な.NETローカルツールを復元し、DocFXでHTMLを生成してからローカルWebサーバーを起動します。
コンソールに表示されたURLをブラウザで開くと、生成されたAPIドキュメントを確認できます。
確認を終えるときは、コマンドを実行しているターミナルで `Ctrl+C` を押してサーバーを停止します。

HTML生成だけを確認する場合は、次のように実行します。

```powershell
pwsh -NoProfile -File .\scripts\serve-csharp-api-docs.ps1 -BuildOnly
```

生成されたHTMLは `docs/api/csharp/_site/` に出力されます。
DocFXが生成する `docs/api/csharp/metadata/` と `docs/api/csharp/_site/` は、再生成できる成果物としてGit管理に含めません。

公開後のC# APIドキュメントは `https://hrmstrsmgs.github.io/SpreadsheetAsData/api/csharp/` で参照できます。

## 現在の確認状況

直近の再整備では、次を確認しています。

* ビルド: 成功
* テスト: 成功、本体176件、コード生成197件
* XMLドキュメント生成: 成功、警告なし
* `dotnet format --verify-no-changes`: 成功

## 制約

NuGetパッケージは公開準備中です。
現時点では、ローカルで生成したパッケージと、リポジトリを取得してC#プロジェクトを直接参照する形で確認しています。

現行C#版には、まだセル値を書き込む公開APIはありません。
過去のRuby版には書き込み機能がありましたが、C#版では再設計しながら追加する予定です。

Ruby版は過去実装です。
現在はC#版を優先して再整備していますが、余裕ができたらRuby版もC#版と同等の機能へ整備する予定です。
