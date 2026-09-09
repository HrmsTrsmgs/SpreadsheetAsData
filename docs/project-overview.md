# SpreadsheetAsData プロジェクト概要

## 目的

SpreadsheetAsDataは、ExcelファイルをExcelアプリケーションの操作対象ではなく、構造化されたデータソースとして扱うためのライブラリです。

Open XML SDKを内部実装として使用しますが、利用側コードからはOOXMLの要素構造を意識せず、ワークブック、ワークシート、定義名、テーブル、行、列、セルをコレクションに近い形で扱えるAPIを目指します。

ExcelのオブジェクトモデルやVBAのAPIをC#上へ再現することは目的としません。

## 中心となる利用方法

SpreadsheetAsDataでは、次の操作を中心に据えます。

1. `.xlsx`ファイルを開く
2. 定義名からセルまたは範囲を取得する
3. Excelテーブルを名前で取得する
4. テーブル行を列名で参照する
5. テーブル行を利用者定義型へ変換する
6. セル、セル範囲、既存のテーブル行を書き換えて保存する
7. `.xlsx`から型付き読み書きコードを生成する

例えば、列の位置を直接指定するのではなく、次のように列名を使って読み取れることを目指します。

```csharp
using var book = Workbook.Open("orders.xlsx");

var orders = book.Tables["Orders"];

foreach (var row in orders.Rows)
{
    Console.WriteLine(row["ProductName"].Value);
    Console.WriteLine(row["Quantity"].Value);
}
```

さらに、テーブル行を利用者定義型へ変換して列挙できるようにします。

```csharp
var orders = book.ReadTable<Order>("Orders");

foreach (var order in orders)
{
    Console.WriteLine(order.ProductName);
}
```

型付きコードを生成すると、ブック、シート、テーブル、行データを表す型を使用して、文字列指定を減らした読み書きもできます。

```csharp
using var book = new OrdersBook();

foreach (var order in book.Orders)
{
    Console.WriteLine(order.ProductName);
}

var updatedOrders = book.Orders.ToArray();
updatedOrders[0].Quantity = 3;

book.Orders.Replace(updatedOrders);
book.Save();
```

現行のコード生成は、読み取りと書き戻しに使用するC#ソース文字列を生成するAPIとして実装しています。
Visual StudioとMSBuildからの生成には `Marimo.SpreadSheetAsData`、生成APIを直接使う場合は `Marimo.SpreadSheetAsData.CodeGeneration` を利用します。

## プロジェクトの経緯

このプロジェクトは、およそ10年前にExcelをデータソースとして扱うためのライブラリとして開始しました。

その後、何度か開発を再開しましたが、再開のたびに.NET、テスト環境、依存ライブラリ、コード記述方式などの更新が必要になり、本来の機能開発へ戻る前に中断する状態が続いていました。

2026年の再整備では、AI支援を利用して既存コードの調査、環境更新、テスト基盤の移行、ドキュメント整備を進め、環境更新だけで終了する循環を初めて抜けました。

現在は、定義名、Excelテーブル、列名による読み取りと書き込み、オブジェクトマッピング、型付きコード生成の基本機能を実装しています。
今後は、基本APIの整理を続けながら、未対応のデータ型やテーブル構造の変更を必要性に応じて追加します。

## 開発の優先順位

基本的な読み取りと書き込みでは、定義名とExcelテーブルを構造化されたデータとして扱えることを優先します。

セル値とセル範囲、既存のテーブル行は書き換えて保存できます。
行の追加や削除、テーブル範囲の拡張、数式、書式などは、実際の利用上の必要性を確認してから対応します。
