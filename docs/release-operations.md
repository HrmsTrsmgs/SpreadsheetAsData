# 公開運用

この文書は、SpreadsheetAsDataの公開作業で使用する手順を記録します。
利用者向けの説明ではなく、リリース作業やGitHub Pages更新を行う開発者向けのメモです。

## GitHub Pages

C# APIドキュメントは、DocFXで生成した静的サイトをGitHub Pagesへ公開します。

ローカルでGitHub Pages用の成果物を生成する場合は、次のスクリプトを使用します。

```powershell
pwsh -NoProfile -File .\scripts\build-github-pages.ps1
```

生成先は `artifacts/github-pages/` です。
C# APIドキュメントは `artifacts/github-pages/api/csharp/` に配置します。
Pagesルートには、C# APIへ移動する `index.html` と、GitHub PagesのJekyll処理を無効にする `.nojekyll` を生成します。

`master` へpushすると、`.github/workflows/publish-api-docs.yml` が同じ生成処理を実行し、GitHub Pagesへデプロイします。
リポジトリのPages設定では、公開元にGitHub Actionsを選択します。

公開後のC# APIドキュメントは次のURLで参照します。

```text
https://hrmstrsmgs.github.io/SpreadsheetAsData/api/csharp/
```
