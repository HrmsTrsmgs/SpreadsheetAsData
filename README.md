# SpreadsheetAsData

いろいろな言語で、Excelのインストールされてない環境でもExcelの読み書きができるライブラリ。
APIがVBAライクではなく、コレクション操作に近い感覚で扱えるようになっている。

## C#版の現行テスト実行手順

C#版は `CSharp/SpreadSheetAsData.sln` に含まれている。

現行のテストプロジェクトは `netcoreapp3.1` を対象にしている。
.NET Core 3.1はサポート終了済みのため、現在の開発環境には入っていないことがある。
その場合、通常の `dotnet test CSharp\SpreadSheetAsData.sln --no-restore` はテストホスト起動時に失敗する。

バージョンアップ前の現行版を確認するための暫定手順として、ユーザー領域に.NET Core 3.1 SDKを入れてから、3.1 SDKに含まれるMSBuildとVSTestを直接実行する。
管理者権限がない環境でも、`$env:USERPROFILE\.dotnet` に入れれば確認できる。

```powershell
Invoke-WebRequest -Uri https://dot.net/v1/dotnet-install.ps1 -OutFile .dotnet-install.ps1
powershell -ExecutionPolicy Bypass -File .\.dotnet-install.ps1 -Version 3.1.426 -InstallDir "$env:USERPROFILE\.dotnet" -Architecture x64

$env:DOTNET_ROOT = "$env:USERPROFILE\.dotnet"
$env:PATH = "$env:USERPROFILE\.dotnet;$env:PATH"

& "$env:USERPROFILE\.dotnet\dotnet.exe" "$env:USERPROFILE\.dotnet\sdk\3.1.426\MSBuild.dll" .\CSharp\SpreadSheetAsData.sln /restore:false
& "$env:USERPROFILE\.dotnet\dotnet.exe" "$env:USERPROFILE\.dotnet\sdk\3.1.426\vstest.console.dll" ".\CSharp\SpreadSheetAdDataのテスト\bin\Debug\netcoreapp3.1\SpreadSheetAdDataのテスト.dll"
```

この手順で確認した結果は次のとおり。

* ビルド: 成功、警告0、エラー0
* テスト: 成功、71件成功

この手順は、.NETバージョンアップ前に現行版の挙動を確認するための一時的なもの。
バージョンアップ後は、通常の `dotnet test` で実行できる手順へ置き換える。

確認後、インストールスクリプトが不要であれば削除する。

```powershell
Remove-Item .\.dotnet-install.ps1
```
