# TableReadingSample

`Marimo.SpreadSheetAsData` をNuGetパッケージとして参照する、利用者向けの最小サンプルです。

このサンプルは、リポジトリ内のプロダクトコードを `ProjectReference` では参照しません。
外部利用者と同じように、`PackageReference` で `Marimo.SpreadSheetAsData` を参照します。

## 実行内容

`TableReadingSample/SampleData/orders.xlsx` に含まれるExcelテーブル `注文一覧` を読み取ります。

サンプルでは次の2通りの読み取りを確認できます。

* `ReadTable<T>` でExcelテーブルの各データ行を `OrderRow` へ対応付ける
* `book.Tables["注文一覧"]` でExcelテーブル、列、行、セルを直接読む

## NuGet公開前に実行する

NuGet.orgへ公開する前は、リポジトリ内で生成したローカルパッケージをNuGetソースとして指定します。

```powershell
dotnet pack .\CSharp\SpreadSheetAsData\SpreadSheetAsData.csproj -c Release
dotnet restore .\samples\TableReadingSample\TableReadingSample.slnx --source .\CSharp\SpreadSheetAsData\bin\Release --source https://api.nuget.org/v3/index.json
dotnet run --no-restore --project .\samples\TableReadingSample\TableReadingSample\TableReadingSample.csproj
```

## NuGet公開後に実行する

NuGet.orgへ公開した後は、通常のNuGetソースから復元できます。

```powershell
dotnet run --project .\samples\TableReadingSample\TableReadingSample\TableReadingSample.csproj
```
