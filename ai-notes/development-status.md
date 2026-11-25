# 開発再開メモ

更新: 2026-11-25。現在の作業だけを上書き更新し、過去の経過はGit履歴を参照する。

## 現在の作業

* 目的: C#・テスト・ビルド連携・Ruby全体を、公開APIと既存の振る舞いを変えずに整理する。変更理由と別件は[全体の精査記録](refactoring-review.md#2026-11-25の全体リファクタリング)を参照。
* 基準: `master` と `origin/master` は `4038667`。F8の修正 `15eb057` は `codex/common-generated-name-validation` にコミット済み。今回の作業ブランチ `codex/repository-refactor` はそのコミットから分けた。
* 現在地: F8と全体リファクタリングを別コミットに分け、`codex/repository-refactor` から `master` 宛てのPRで取り込む。

## 検証と次回

* 検証: .NET SDK 11.0.100でアナライザー付きビルドは警告・エラー0。C#全934件（本体371、コード生成・MSBuild連携563）が成功。保留2例を加えた後は該当Theoryを再実行し、5件成功・2件Skip。Rubyは隔離したデータでRSpec 262例・失敗0・pending 3、サンプルも成功。PowerShellの構文検査は成功。Visual Studioの画面確認、公開操作、性能計測は対象外。
* 静的検査: 全体の `dotnet format --verify-no-changes --severity info` に残るのは `CellRange.Values` の `CA2208` 1件。寸法不一致時の `ArgumentException` のメッセージ・`ParamName` は変更していない。
* 次の着手点: PRの取り込み後、必要に応じて保留例を別のRedから扱う。Rubyの列順とC#の生成名の先頭文字の修正は、この差分に含めない。
