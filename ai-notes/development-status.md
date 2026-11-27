# 開発再開メモ

更新: 2026-11-28。現在の作業だけを上書き更新し、過去の経過はGit履歴を参照する。

## 現在の作業

* 目的: Rubyを除くDRY精査の候補を一件ずつ扱う。[非公開GitHub Project](https://github.com/users/HrmsTrsmgs/projects/1)の対応項目は進行中。
* 基準: `codex/dry-cleanup` の `17ad781`。Int32判定と生成型名の共通化はコミット済み。単一セル判定は変更なし。
* 現在地: `.targets` で生成ファイルの接尾辞を定義し、Visual Studioの関連ファイル表示、コンパイル登録、Cleanで共有する差分がレビュー待ち。`.props` の早期評価とC#タスクの直接利用のため、この2箇所の定義は維持。
* 衝突診断と生成名先頭文字の不具合修正は別ブランチ `codex/fix-identifier-start` の担当チャットに引き継ぎ済み。衝突診断のコミットは `codex/dry-follow-up` に残る。DRY作業のレビュー場所は `dry-follow-up` 作業ツリーに戻し、同じ場所で `codex/dry-cleanup` を使用する。

## 検証と次回

* 生成型名の変更は関連テスト518件成功・既存Skip 2件、非増分ビルドと変更箇所の整形検査も成功。
* 今回は単独インポートのMSBuild評価でLastGenOutputとCompile項目を確認し、Cleanで実行時に追加した項目の生成ファイルが削除され、元Excelが残ることを確認。既存MSBuild連携テストは結果が出る前に中断し、パッケージ経由の回帰確認は未完了。
* 次の着手点: 未コミット差分の人間レビューとコミット了承を待つ。
