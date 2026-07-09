# SpreadsheetAsData

いろいろな言語で、Excelのインストールされてない環境でもExcelの読み書きができるライブラリ。
APIがVBAライクではなく、コレクション操作に近い感覚で扱えるようになっている。

C#版は `CSharp/SpreadSheetAsData.sln` に含まれている。
C#版のライブラリ本体とテストプロジェクトは `net8.0` を対象にしている。

.NET 8 SDKが入っている環境では、次のコマンドでビルドとテストを実行できる。

```powershell
dotnet test .\CSharp\SpreadSheetAsData.sln
```

この手順で確認した結果は次のとおり。

* ビルド: 成功
* テスト: 成功、71件成功
