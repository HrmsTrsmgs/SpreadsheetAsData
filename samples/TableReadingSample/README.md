# TableReadingSample

`Marimo.SpreadSheetAsData` をNuGetパッケージとして参照する、利用者向けの最小サンプルです。

このサンプルは、リポジトリ内のプロダクトコードを `ProjectReference` では参照しません。
外部利用者と同じように、`PackageReference` で `Marimo.SpreadSheetAsData` を参照します。

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
**Streamだけで編集結果の出力まで完結するAPIは、現時点では提供していません。** 出力には`SaveAs(path)`でファイルパスを指定します。
ファイルパスから開いた場合の`Save()`は、正常終了時点で元ファイルへの保存を完了します。

## NuGet公開前に実行する

NuGet.orgへ公開する前は、リポジトリ内で生成したローカルパッケージをNuGetソースとして指定します。

```powershell
dotnet pack .\CSharp\SpreadSheetAsData.slnx -c Release -o .\artifacts\nupkg
$localFeed = (Resolve-Path .\artifacts\nupkg).Path
dotnet restore .\samples\TableReadingSample\TableReadingSample\TableReadingSample.csproj --source $localFeed --source "https://api.nuget.org/v3/index.json"
dotnet run --no-restore --project .\samples\TableReadingSample\TableReadingSample\TableReadingSample.csproj
```

## NuGet公開後に実行する

NuGet.orgへ公開した後は、通常のNuGetソースから復元できます。

```powershell
dotnet run --project .\samples\TableReadingSample\TableReadingSample\TableReadingSample.csproj
```
