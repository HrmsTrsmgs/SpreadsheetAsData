# TDDの入口

通常は共通Skill `tdd-workflow` を使う。共通Skillが利用できない環境では、リポジトリ内の [進行手順](tdd/run.md) を参照版として使う。プロジェクト固有の指示と記録先は [AGENTS.md](../../AGENTS.md) に従う。

共通標準はRed→Green→人間レビュー・Gコミット→Refactor→変更時の人間レビュー・Fコミット。段階指定ではそこで停止し、Refactorは検討するが不要な変更や空コミットを作らない。

進行役は進行手順だけ、工程担当は [工程共通](tdd/common.md) と自分の工程だけを読む。

- [Red](tdd/red.md)
- [Green](tdd/green.md)
- [Refactor](tdd/refactor.md)
- [モデル方針（進行役が必要な場合だけ）](tdd/models.md)

共通Skillと参照版は同じ工程内で混用しない。参照版の基準は `codex-workflows e66db1d`。この索引は工程別の検討を一括で読み込ませるためのものではない。
