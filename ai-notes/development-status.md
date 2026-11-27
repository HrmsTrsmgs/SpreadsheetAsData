# 開発再開メモ

更新: 2026-11-27。現在の作業だけを上書き更新し、過去の経過はGit履歴を参照する。

## 現在の作業

* 目的: C#の2テストプロジェクトをMicrosoft.Testing.Platformへ移行する。実行手順はREADMEを参照。
* 基準: `codex/microsoft-testing-platform` の起点は `5d4a155`。GitHub Projectの移行項目は進行中。依存と構成の移行案は了承済みで、コミット前の差分レビュー待ち。
* 変更: 両テストを `xunit.v3.mtp-v2` 4.0.2へ切り替え、`global.json` でMTPを選択した。復元されたMTPは2.5.1。テスト用のMSBuild参照を.NET SDK同梱版に揃え、テストから起動する子ビルドのノード再利用を無効にした。

## 検証と次回

* .NET SDK 11.0.100でのアナライザー付き非増分ビルドは警告・エラー0件。通常の `dotnet test` とVisual Studio 2026 18.12のテスト エクスプローラーは、それぞれ成功934件・失敗0件・既存のSkip 2件。Visual StudioではMTPホストでの実行を確認した。
* 修正前はVisual Studioの.NET Framework用MSBuild DLLを読み込んで `System.Xaml` が見つからず、ノード再利用が有効な条件で一時DLLの削除にも失敗した。修正後は両方の実行経路で解消した。
* `dotnet format --verify-no-changes --severity info` に残るのは変更対象外の `CellRange.cs` の既存 `CA2208` 1件。変更箇所にアナライザー指摘はなし。
* 次の着手点: 差分の了承後、`master` 宛てのPRを作成し、GitHub Projectの移行項目に関連づける。取り込み後に完了とする。
