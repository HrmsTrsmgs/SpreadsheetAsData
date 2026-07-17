# コード生成のテスト仕様

## 目的

`.xlsx` ファイルから型付き読み取り用の C# ソースを生成する機能について、内部実装を固定せず、外部から観測できる仕様を定義する。

この文書は実装前のテスト設計の基準とする。`Core`、Open XML SDK の保持方法、生成時のファクトリー構造などは対象に含めない。

## 入出力

入力は次の二つとする。

- `.xlsx` ファイル
- コード生成設定

出力は次の二つとする。

- C# ソースコード一式
- 生成診断

## 生成型の構造

生成型は次の構造とする。

```csharp
OrderDataBook : Workbook
SalesDataSheet : Worksheet
SalesDetailTable : Table<SalesDetailRow>

SalesDetailRow // 通常の POCO
```

Book、Sheet、Table は既存の非型付き API を継承して利用できるようにする。Row は `TableRow` を継承しない POCO とする。

## テスト一覧

### 1. 生成型の構造

1. 生成された Book 型は `Workbook` を継承する。
2. 生成された Sheet 型は `Worksheet` を継承する。
3. 生成された Table 型は生成 Row を型引数とする `Table<T>` を継承する。
4. 生成された Row 型は `TableRow` を継承しない POCO である。
5. 生成された Row の列プロパティは public な getter と setter を持つ。
6. 生成された型は別ファイルの partial 定義と共にコンパイルできる。

### 2. Book からのアクセス

7. Book は各ワークシートを型付きプロパティとして公開する。
8. Book は各 Excel テーブルを型付きプロパティとして公開する。
9. 生成された Book 型から `Workbook` の非型付き API を使用できる。

確認対象には `Sheets`、`Tables`、`Cell`、`Range`、シート名インデクサーを含める。

### 3. Sheet からのアクセス

10. Sheet はそのシートに属する Excel テーブルを型付きプロパティとして公開する。
11. Sheet は別シートに属する Excel テーブルを公開しない。
12. 生成された Sheet 型から `Worksheet` の非型付き API を使用できる。

確認対象には `Name`、`Book`、`Cell`、`Range`、`Cells` を含める。

### 4. 定義名

13. ブックスコープの単一セル定義名を Book の `Cell` プロパティとして生成する。
14. ブックスコープの複数セル定義名を Book の `CellRange` プロパティとして生成する。
15. シートローカルの単一セル定義名を Sheet の `Cell` プロパティとして生成する。
16. シートローカルの複数セル定義名を Sheet の `CellRange` プロパティとして生成する。

ブックスコープとシートローカルで同名の定義名が存在しても、それぞれの型に別プロパティとして生成できるものとする。

### 5. 型付き Table

17. 生成された Table は POCO の Row を Excel 上の順序で列挙する。
18. 生成された Table を `Table` として扱うと非型付き Row を利用できる。
19. 生成された Table 型から `Table` の構造情報を使用できる。
20. 生成 Row は利用者定義 POCO と同じ変換規則で読み込まれる。

構造情報には `Name`、`Worksheet`、`Range`、`Columns` を含める。

### 6. 列型の生成

初期段階では `int`、`double`、`string` のみを対象とする。

21. 整数値だけを持つ列を `int` プロパティとして生成する。
22. 小数値を持つ列を `double` プロパティとして生成する。
23. 文字列値を持つ列を `string` プロパティとして生成する。
24. 生成された列プロパティに元の Excel 列名を設定する。

例：

```csharp
[SpreadsheetColumn("customer_id")]
public int CustomerId { get; set; }
```

### 7. ASCII 名の自動変換

25. ASCII だけの名前を PascalCase へ変換する。

代表例：

```text
sales_detail  -> SalesDetail
sales-detail  -> SalesDetail
sales detail  -> SalesDetail
SALES_DETAIL  -> SalesDetail
customer_id   -> CustomerId
```

この規則は次の名前へ適用する。

- ファイル名から生成する Book 型名
- Sheet 型名とプロパティ名
- Table 型名とプロパティ名
- Row 型名
- 列プロパティ名
- 定義名プロパティ名

Table と Row にはそれぞれ接尾辞を付ける。

```text
sales_detail -> SalesDetailTable
sales_detail -> SalesDetailRow
sales_detail -> SalesDetail
```

### 8. 日本語・ローマ字混在名

26. 非 ASCII 文字を含む名前は先頭文字以外の有効な表記を変更しない。
27. 非 ASCII 文字を含む名前の使用できない識別子文字をアンダースコアへ置換する。

代表例：

```text
商品_明細          -> 商品_明細
商品_id            -> 商品_id
sales商品_detail   -> Sales商品_detail
商品sales_detail   -> 商品sales_detail
商品 明細          -> 商品_明細
商品-明細          -> 商品_明細
sales商品-detail   -> Sales商品_detail
```

先頭が ASCII 小文字の場合は、その文字だけを大文字化する。内部の大小文字やアンダースコアから単語境界を推測しない。

### 9. 簡易 NameMappings

28. `NameMappings` は自動名前変換より優先される。
29. `NameMappings` は対象種類を指定せず、同じ元名へ全体的に適用される。

例：

```csharp
options.NameMappings["cust_id"] = "CustomerID";
```

生成名を再度 PascalCase 変換しない。

### 10. 文脈付き名前指定

30. 文脈付き名前設定は同じ元列名を Table ごとに異なる名前へ変更できる。
31. 文脈付き名前設定はブック定義名とシートローカル定義名を区別できる。
32. 文脈付き名前設定は簡易 `NameMappings` より優先される。

具体的な Fluent API のメソッド名は、この段階では固定しない。

名前決定の優先順位は次とする。

1. 文脈付き名前設定
2. `NameMappings`
3. 自動変換

### 11. 名前衝突診断

33. 自動変換後に同じ列プロパティ名となる場合にエラーを診断する。
34. 名前衝突を自動的な連番追加では解消しない。
35. `NameMappings` で生成名を変更すると名前衝突を解消できる。
36. 同じ生成型内の異なる種類のメンバー名が衝突した場合にも診断する。
37. 有効な C# 識別子を生成できない名前を診断する。

例：

```text
customer_id -> CustomerId
customer-id -> CustomerId
```

この場合、`CustomerId2` などを自動生成せずエラーとする。

### 12. 生成結果全体

38. 指定した名前空間へすべての型を生成する。
39. 有効な Excel ファイルから生成したすべてのソースはコンパイルできる。
40. 同じ Excel ファイルと設定から同じ生成結果を返す。
41. 正常な Excel ファイルではエラー診断を返さない。

生成ソースのコンパイルテストでは、型名、相互参照、属性、継承関係、構文の整合性を確認する。

## テストデータ構成

コード生成用の Excel ファイルは専用フォルダへまとめる。

```text
TestData/
└─ コード生成/
   ├─ 基本構造.xlsx
   ├─ 定義名.xlsx
   ├─ ASCII名前変換.xlsx
   ├─ 日本語混在名前.xlsx
   ├─ 簡易名前置換.xlsx
   ├─ 文脈付き名前置換.xlsx
   ├─ 列名衝突.xlsx
   ├─ Bookメンバー名衝突.xlsx
   ├─ 無効名.xlsx
   └─ 統合.xlsx
```

テストプロジェクトでは個別ファイルを列挙せず、フォルダ単位で出力先へコピーする。

```xml
<ItemGroup>
  <Content Include="TestData\コード生成\**\*.xlsx">
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </Content>
</ItemGroup>
```

## 今回固定しない事項

次はこのテスト仕様に含めない。

- 内部 Core の有無
- 内部で Open XML SDK をどう保持するか
- 公開 API からの Open XML SDK 非露出を動的に検査するテスト
- Book、Sheet、Table のオブジェクト同一性
- 型付きオブジェクトのキャッシュ方法
- `Table<T>` の内部マッピング実装
- 生成ソースの改行、インデント、ファイル分割方法
- 生成される `.cs` ファイル名
- CLI
- MSBuild
- Source Generator
- Visual Studio 統合
- ファイルへの物理出力
- nullable
- `DateTime`
- `DateOnly`
- `decimal`
- `long`
- `bool`
- enum
- 数式セル
- 混在型列
- 空列

## 件数

この文書で定義するテストは合計 41 本である。
