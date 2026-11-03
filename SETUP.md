# GitHub Pagesで配布する手順

公開画面は github-pages/。既存のdist/は確認用で送信不可です。
回答データ・認証情報はリポジトリに保存しません。

1. Googleスプレッドシートを作成し、拡張機能→Apps Scriptを開く。
2. integration/Code.gs と integration/Index.html を設置する。
3. ウェブアプリとしてデプロイ（実行ユーザー：自分、アクセス：対象の外部回答者が利用できる範囲）。必要なGoogle認証は所有者が行う。
4. github-pages/config.js にウェブアプリの /exec URLを記入する。
5. GitHubリポジトリを作り、このソースをpushする。
6. Settings → Pages → Source を GitHub Actions にする。
7. Actionsの成功後、発行されたGitHub Pages URLを開く。
8. 架空データで実際に送信し、4つのシートへの保存、二重送信防止、スマホ表示を確認する。
9. テスト登録を削除してURLを配布する。

GitHub PagesはGoogle Apps Scriptの入力画面をフレーム内で表示するため、ブラウザーのアドレスはGitHub Pagesのままです。Google側のアクセス設定やブラウザーのCookie制限によって動作が異なる場合があるため、未ログイン状態でもテストしてください。

未完了：Google接続・実送信テスト・締切自動停止・車両1台の制限。
GitHub Pagesでの公開可否はアカウントプランとリポジトリ公開範囲に依存します。
