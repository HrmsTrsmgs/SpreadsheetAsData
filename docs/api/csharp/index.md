# SpreadsheetAsData C# API

SpreadsheetAsDataは、Excelファイルを低水準のOpen XML要素ではなく、ワークブック、ワークシート、定義名、Excelテーブル、行、列、セルとして読み取るためのC#ライブラリです。
このAPIリファレンスでは、NuGetパッケージから利用する公開APIを確認できます。

## 主要な入口

通常の読み取りは、`Workbook.Open` から始めます。
開いたワークブックから、ワークシート、定義名、Excelテーブル、セル、セル範囲を取得できます。

ExcelテーブルをC#の利用者定義型へ変換する場合は、`Workbook.ReadTable<T>` または `Table.Enumerate<T>` を使用します。
列名とプロパティ名が異なる場合は、`SpreadsheetColumnAttribute` で対応するExcel列名を指定します。

## 型付きコード生成

NuGetパッケージのMSBuild連携を使うと、`.xlsx` ファイルからブック、シート、テーブル、行データを表すC#コードを生成できます。
生成されたBook型からは、テーブルや行データを文字列指定を減らして読み取れます。

コード生成機能の設定やVisual Studioでの使い方は、リポジトリのREADMEを参照してください。

## 代表的な型

* `Workbook`: `.xlsx` ワークブックを表します。
* `Worksheet`: ワークシートを表します。
* `DefinedName`: Excelの定義名を表します。
* `Cell`: セルを表します。
* `CellRange`: セル範囲を表します。
* `Table`: Excelテーブルを表します。
* `TableColumn`: Excelテーブル列を表します。
* `TableRow`: Excelテーブルのデータ行を表します。
* `Table<T>`: Excelテーブル行を利用者定義型として列挙します。
* `TableMappingException`: 型付きテーブル読み取りで、列対応や値変換に失敗したときの例外です。

## 関連情報

プロジェクト全体の目的、使用例、現在の制約は、リポジトリ直下のREADMEと `docs/` 配下の文書を参照してください。
