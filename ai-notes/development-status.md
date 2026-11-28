# 開発再開メモ

更新: 2026-11-28。現在の作業だけを上書き更新し、過去の経過はGit履歴を参照する。

## 現在の作業

* 目的: `NameMappings` の生成名の先頭にある書式文字を診断する。[非公開GitHub Project](https://github.com/users/HrmsTrsmgs/projects/1)の「SpreadsheetAsData: C# 生成名の先頭で書式文字を拒否する」。
* 基準: PR #19取り込み後の `origin/master`（`9f570a6`）。元の作業ツリーの `codex/reject-leading-format` で検証完了。
* 変更: 先頭の `U+200C` の既存Skipを解除。書式文字を既存の後続専用グループに移し、途中の文字の許容は維持した。

## 検証と次回

* .NET SDK 11.0.100、MTP。コード生成診断クラスのRedは201件成功、対象の1件が診断なしで失敗。Greenは202件成功。全C#テストは858件成功、失敗・Skipなし。アナライザー付き非増分ビルドは警告・エラーなし。`dotnet format` は既存の `CellRange.cs` の `CA2208` のみ。`git diff --check` は成功。Rubyは未変更のため未実行。
* 次の着手点: `master` 宛てPRを確認し、マージの指示後に取り込む。
