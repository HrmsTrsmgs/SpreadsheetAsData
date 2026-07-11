# SpreadsheetAsData

SpreadsheetAsDataは、Excelをインストールしていない環境でもExcelファイルを扱えるようにするライブラリです。

現在はC#版を再整備中です。
Open XML SDKを内部実装として使いながら、利用側コードからはワークブック、ワークシート、セルをコレクション操作に近い感覚で扱えるAPIを目指しています。

## 現在のC#版でできること

* `.xlsx` ファイルを開く
* ワークシートを名前または位置で取得する
* セルをA1形式、または列番号と行番号で取得する
* セル参照、行番号、列番号を取得する
* 空白、数値、真偽値、共有文字列セルの値を取得する

現在のC#版は、読み取り機能を中心に再整備している段階です。
書き込みは、このライブラリの主要な拡張対象です。
書式、日付、テーブル、広範なExcel機能への対応は、書き込みの基礎機能を整えた後の補助機能として扱います。

## 使用例

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
* 実装、テスト、READMEの内容を矛盾させない
* 変更しやすい小さな単位で機能を追加する

詳細な設計方針は [docs/design.md](docs/design.md) を参照してください。

## ビルドとテスト

C#版は `CSharp/SpreadSheetAsData.sln` に含まれています。
ライブラリ本体とテストプロジェクトは `net8.0` を対象にしています。

.NET 8 SDKが入っている環境では、次のコマンドでビルドとテストを実行できます。

```powershell
dotnet build .\CSharp\SpreadSheetAsData.sln
dotnet test .\CSharp\SpreadSheetAsData.sln
```

整形と基本的なスタイルチェックは `.editorconfig` に定義しています。

```powershell
dotnet format .\CSharp\SpreadSheetAsData.sln --verify-no-changes --no-restore --severity warn
```

## 現在の確認状況

直近の再整備では、次を確認しています。

* ビルド: 成功
* テスト: 成功、71件成功
* XMLドキュメント生成: 成功、警告なし
* `dotnet format --verify-no-changes`: 成功

## 制約

現行C#版には、まだセル値を書き込む公開APIはありません。
過去のRuby版には書き込み機能がありましたが、C#版では再設計しながら追加する予定です。

Ruby版は過去実装です。
現在はC#版を優先して再整備していますが、余裕ができたらRuby版もC#版と同等の機能へ整備する予定です。
