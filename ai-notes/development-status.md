# 開発再開メモ

更新: 2026-11-24。現在の作業だけを上書き更新し、過去の経過はGit履歴を参照する。

## 現在の作業

* 目的: .NET 11 / C# 15へ移行し、テスト補助のPowerShell起動・出力取得を簡素化する。
* 基準: `79b6be4`。.NET 11への移行とPowerShell起動補助の簡素化をコミット済み。
* 現在地: 全プロジェクト、NuGetのタスク配置、開発環境チェック、文書CIを .NET 11 に揃えた。テスト補助は `Process.RunAndCaptureTextAsync` を使う。セル値のUnion化、生成ソースの構文、パッケージの公開バージョンは変更していない。

## 検証と次回

* 検証: SDK 11.0.100でビルドの警告・エラー0。全917テスト（コード生成・MSBuild連携551件、本体366件）が成功。全部入りとBuild単体のパッケージで生成型のコンパイルとExcel読み書きを確認。環境チェック、同梱のテーブル読み取りサンプル、変更したC#ファイルの書式検査も成功。端末にないMAUIワークロードを誤って解決しないよう、検証プロセスだけ `MSBuildEnableWorkloadResolver=false` を指定した。
* Visual Studio連携の検証: Git管理外の `artifacts/vs-net11-check/Consumer/SpreadsheetVsCheck.csproj` でローカルNuGet `0.4.1-vscheck` を参照。Visual Studio 18.12同梱のMSBuildでビルドしてサンプル4行を読み取れた。デザイン時評価で生成ソースの `DependentUpon` は `sales_report.xlsx`。IDEの画面取得では対象外の画面が返り、ソリューションエクスプローラーの表示は未確認。
* 次の着手点: 検証プロジェクトをVisual Studioで開き、Excel配下の `.g.cs` 表示を確認する。製品の精査に戻る場合は、F8の識別子診断の対応範囲を現行コードと照合してからF9へ進む。文書CIは未確認。pushとNuGet公開は未承認。
