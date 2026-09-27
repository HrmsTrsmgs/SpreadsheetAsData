# TableReadingSample

`Marimo.SpreadsheetAsData` をNuGetパッケージとして参照する、利用者向けの最小サンプルです。

このサンプルは0.4.0向けです。0.3.0とは名前空間・属性名が異なります。変更点は[移行案内](../../README.md#040への更新)を参照してください。

このサンプルは、リポジトリ内のプロダクトコードを `ProjectReference` では参照しません。
外部利用者と同じように、`PackageReference` で `Marimo.SpreadsheetAsData` を参照します。

## 実行内容

`TableReadingSample/SampleData/orders.xlsx` に含まれるExcelテーブル `注文一覧` を読み取ります。

サンプルでは次の2通りの読み取りを確認できます。

* `ReadTable<T>` でExcelテーブルの各データ行を `OrderRow` へ対応付ける
* `book.Tables["注文一覧"]` でExcelテーブル、列、行、セルを直接読む

## Streamから開いて編集する場合

`Program.cs`の`using var book = Workbook.Open(workbookPath);`を次の処理に置き換えると、Stream入力から編集し、別ファイルへ保存できます。

```csharp
using var stream = File.OpenRead(workbookPath);
using var book = Workbook.Open(stream);

book.Tables["注文一覧"].Rows.First()["数量"].Value = 3;
book.SaveAs("updated-orders.xlsx");
```

Streamから開いた場合、`Save()`は元Streamの内容を変更する前に`NotSupportedException`を投げます。
`Close()` / `Dispose()`も保存せず、渡したStream自体は閉じません。
0.4.0では`SaveAs(Stream)`で空の`MemoryStream`へも出力できます。上の例は従来どおり`SaveAs(path)`でファイルへ保存します。Stream出力の例と制約は[README](../../README.md)を参照してください。
ファイルパスから開いた場合の`Save()`は、正常終了時点で元ファイルへの保存を完了します。

## NuGet公開前に実行する

NuGet.orgへ公開する前は、リポジトリ内で生成したローカルパッケージをNuGetソースへ追加します。通常のNuGet設定でNuGet.orgが有効になっていることを前提とします。

```powershell
dotnet pack .\CSharp\SpreadsheetAsData.slnx -c Release -o .\artifacts\nupkg
$localFeed = (Resolve-Path .\artifacts\nupkg).Path
dotnet restore .\samples\TableReadingSample\TableReadingSample\TableReadingSample.csproj --packages .\artifacts\sample-packages "-p:RestoreAdditionalProjectSources=$localFeed"
dotnet run --no-restore --project .\samples\TableReadingSample\TableReadingSample\TableReadingSample.csproj
```

## NuGet公開後に実行する

NuGet.orgへ公開した後は、通常のNuGetソースから復元できます。

```powershell
dotnet run --project .\samples\TableReadingSample\TableReadingSample\TableReadingSample.csproj
```
