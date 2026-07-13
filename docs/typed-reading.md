# 型付き読み取り

## 目的

SpreadsheetAsDataでは、Excelテーブルを単なるセル範囲ではなく、列定義を持ったデータ集合として読み取ります。

型付き読み取りは、次の三段階で提供する方針です。

1. 列名を使用した行読み取り
2. 利用者定義型へのマッピング
3. `.xlsx`からの型付き読み取りコード生成

書き込み機能は、この文書の対象外です。

## 1. 列名による行読み取り

最も基本的な読み取り方法です。

テーブル行を列挙し、各値を列名で取得します。

```csharp
using var book = Workbook.Open("orders.xlsx");

var table = book.Tables["Orders"];

foreach (var row in table.Rows)
{
    var productName = row["ProductName"];
    var quantity = row["Quantity"];
    var unitPrice = row["UnitPrice"];
}
```

この層では、利用者がC#型を事前に定義する必要はありません。

列名の誤りや値の型変換は、実行時に検出されます。

## 2. 利用者定義型へのマッピング

利用者が定義した型へ、テーブルの各行を変換します。

```csharp
public sealed class Order
{
    public string ProductName { get; init; } = "";
    public int Quantity { get; init; }
    public decimal UnitPrice { get; init; }
}
```

```csharp
using var book = Workbook.Open("orders.xlsx");

IEnumerable<Order> orders = book.Tables["Orders"].As<Order>();

foreach (var order in orders)
{
    Console.WriteLine(order.ProductName);
}
```

API名は設計中の例です。

### 基本規則

初期実装では、次の規則を基本とします。

* プロパティ名と列名が一致する場合に対応付ける
* 対象は公開された書き込み可能なプロパティとする
* 空白値は参照型またはnullable値型へ変換できる
* 空白値を非nullable値型へ変換する場合はエラーとする
* 変換できない値を暗黙に既定値へ置き換えない
* 不足している必須列は読み取り開始時に検出する
* 使用されない余剰列は許可する

属性による列名指定は、必要性を確認してから追加します。

```csharp
public sealed class Order
{
    [SpreadsheetColumn("商品名")]
    public string ProductName { get; init; } = "";
}
```

属性名とAPIは設計中の例です。

## 3. `.xlsx`からの型付き読み取りコード生成

型付きDataSetに近い利用方法として、`.xlsx`自体をスキーマの入力としてC#コードを生成します。

生成対象は、初回公開版ではExcelテーブルに限定します。

例えば、`Orders`というテーブルから次のような行型を生成します。

```csharp
public sealed class OrderRow
{
    public string ProductName { get; init; } = "";
    public int Quantity { get; init; }
    public decimal UnitPrice { get; init; }
}
```

ワークブック全体を表す型も生成し、次のように利用できることを目標とします。

```csharp
using var book = OrdersWorkbook.Open("orders.xlsx");

foreach (var order in book.Orders)
{
    Console.WriteLine(order.ProductName);
    Console.WriteLine(order.Quantity);
}
```

生成される型名、メソッド名、プロパティ名は設計中の例です。

## 型の判定

Excelテーブルには、データベースの列型に相当する固定的な型定義がないため、生成時に値を調査してC#型を決定します。

初期対応候補は次のとおりです。

| Excel上の値 | C#型候補 |
|---|---|
| 文字列 | `string` |
| 整数として扱える数値 | `int`または`long` |
| 小数を含む数値 | `decimal`または`double` |
| 真偽値 | `bool` |
| 日付 | `DateTime` |
| 空白を含む値型列 | 対応するnullable型 |

実際の型決定規則は、実装時にテスト可能な形で固定します。

曖昧な列を無理に狭い型へ変換せず、安全に読み取れる型を選ぶことを優先します。

## 列名からプロパティ名への変換

Excelの列名は、そのままではC#の識別子として使用できない場合があります。

コード生成では、少なくとも次のケースを扱います。

* 空白を含む列名
* 記号を含む列名
* 数字から始まる列名
* C#の予約語と一致する列名
* 大文字小文字だけが異なる列名
* 変換後に同じ名前になる複数の列
* 空の列名
* 日本語を含む列名

元の列名と生成後のプロパティ名の対応は、生成コードまたは生成時の診断情報から確認できるようにします。

## スキーマの不一致

コード生成後に`.xlsx`の構造が変更される可能性があります。

読み取り時には、少なくとも次の不一致を検出できるようにします。

* 必要なテーブルが存在しない
* 必要な列が存在しない
* 同名の列が複数存在する
* 値を生成型へ変換できない
* テーブル名または列名が変更されている

不一致を暗黙に無視して誤ったデータを生成するより、原因が分かる例外または診断結果を返すことを優先します。

## 生成コードの依存関係

生成された公開型には、Open XML SDKの型を露出させません。

生成コードは、SpreadsheetAsDataの公開APIだけを利用してデータを読み取る構造とします。

これにより、Open XML SDKのバージョンやOOXMLの内部構造が、利用側コードへ直接影響することを避けます。

## 初回公開版では扱わないこと

* 生成型を使用した書き込み
* 双方向データバインディング
* Excel数式からC#プロパティを生成すること
* 複数テーブル間の関連を自動推測すること
* 主キーや外部キーの自動推測
* 任意のセル範囲からの型生成
* Excelブック全体を完全なデータベーススキーマとして解釈すること

初回公開版では、各Excelテーブルを独立した型付きデータ集合として読み取ることに集中します。