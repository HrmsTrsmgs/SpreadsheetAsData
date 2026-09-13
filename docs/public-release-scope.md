# 現行公開版の範囲

## 目的

現行公開版では、SpreadsheetAsDataの中心概念である「Excelを構造化されたデータソースとして扱う」ことを、実際に利用できる形で示します。

Excelの全機能への対応ではなく、通常の業務データを格納した`.xlsx`ファイルを読み書きするための基本機能に範囲を限定します。

## 対象機能

### 基本的なセルの読み書きと保存

* ファイルパスまたは `Stream` から `.xlsx` を開く
* ワークシートを名前または位置で取得する
* セルをA1形式または行列番号で取得する
* 空白、文字列、数値、真偽値を読み取る
* セルへ数値、文字列、真偽値、空白を書き込む
* 矩形範囲へ行ごとの値を書き込む
* ファイルパスから開いた場合、`Save()` の正常終了時点で元ファイルへの保存を完了する
* 変更内容を別ファイルへ保存する
* Open XML SDKの型を公開APIへ露出させない

Streamから開いた場合の `Save()` は、元Streamの内容を変更する前に `NotSupportedException` を投げます。
`Close()` / `Dispose()` は保存を行わず、呼び出し側から渡されたStreamも閉じません。
Stream版でも `SaveAs(path)` は利用できますが、Streamだけで編集結果の出力まで完結するAPIは提供していません。

### 定義名

* ワークブックに定義された名前を列挙する
* 定義名を名前で取得する
* 定義名から単一セルを取得する
* 定義名から矩形範囲を取得する

単一セルに付けられた名前と複数セル範囲に付けられた名前は、同じ定義名機能として扱います。

現行版では、単一セルまたは矩形範囲を参照する定義名を対象とします。

定数、数式、外部ブック参照などを表す定義名は対象外とします。

### Excelテーブル

* ワークブック内のテーブルを列挙する
* テーブルを名前で取得する
* テーブル名、対象範囲、列定義を取得する
* ヘッダー名を取得する
* データ行を列挙する
* テーブル列を名前で取得する
* 型付き行から既存のデータ行を置き換える

現行版では、ヘッダー行とデータ行を持つ標準的なExcelテーブルを対象とします。

### 列名による行データの読み書き

テーブル行を、列番号ではなく列名で参照できるようにします。

```csharp
var table = book.Tables["Orders"];

foreach (var row in table.Rows)
{
    var productName = row["ProductName"].Value;
    var quantity = row["Quantity"].Value;
}
```

列順序が変更されても、利用側コードの変更を最小限にできることを目標とします。
取得したセルの `Value` へ値を設定することで、列名を使って行データを書き換えられます。

### 利用者定義型へのマッピング

テーブルの各行を、利用者が定義した型へ変換して列挙できるようにします。

```csharp
public sealed class Order
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
    public double UnitPrice { get; set; }
}
```

```csharp
IEnumerable<Order> orders = book.ReadTable<Order>("Orders");
```

既存のテーブル行は、同じ型の値から置き換えられます。

```csharp
var table = book.ReadTable<Order>("Orders");
var orders = table.ToArray();

orders[0].Quantity = 3;

table.Replace(orders);
book.Save();
```

現行版では、プロパティ名と列名の一致を基本規則とします。

列名とプロパティ名が異なる場合は、`SpreadSheetNameAttribute` で列名を指定します。

### 型付き読み書きコードの生成

`.xlsx`に含まれる定義名とExcelテーブルを解析し、ブック、シート、テーブル、行データの型付き読み書きAPIを生成します。

生成されたコードからは、定義名、テーブル名、列名を文字列で指定せずにデータを読み書きできます。

```csharp
using var book = new OrdersBook();

foreach (var order in book.Orders)
{
    Console.WriteLine(order.ProductName);
    Console.WriteLine(order.Quantity);
}

var orders = book.Orders.ToArray();
orders[0].Quantity = 3;

book.Orders.Replace(orders);
book.Save();
```

現行実装では、生成されたBook型の引数なしコンストラクターに生成元Excelファイルのパスを埋め込みます。
通常の利用では、実行時API、コード生成API、Visual StudioとMSBuildの連携をまとめた `Marimo.SpreadSheetAsData` パッケージを使用します。

## 公開版の確認条件

公開時には、次の条件を確認します。

* 対象機能が公開APIとして利用できる
* 各公開機能に自動テストがある
* サンプル用の`.xlsx`と利用コードがある
* READMEから基本的な利用方法を確認できる
* APIドキュメントが生成できる
* 対応範囲と未対応範囲が明記されている
* Open XML SDKの型が通常の利用側コードへ露出していない
* NuGetパッケージとコード生成機能の提供単位が決まっている
* ビルド、テスト、コード整形確認が成功する

## 現行版の対象外

次の機能は現行版には含めません。

* 行の追加、挿入、削除
* テーブル範囲の拡張と縮小
* 数式の計算
* Excelと同等の数式エンジン
* 書式の詳細操作
* テーブルスタイルの操作
* フィルターや並べ替え状態の操作
* グラフ、画像、図形
* ピボットテーブル
* マクロ
* 外部ブック参照
* Excelの全構造化参照への対応
* 大規模ファイル向けストリーミング更新
* Excelの全機能との互換性

これらは、基本的なデータの読み書きを整えた後の拡張対象とします。
