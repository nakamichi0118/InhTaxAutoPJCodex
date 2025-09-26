# デプロイ手順書

このプロジェクトをGitHub、Railway、Cloudflare Pagesにデプロイするための手順書です。

## ステップ 1: GitHubリポジトリの作成とプッシュ

### 1-1. GitHub CLIを使用する方法（推奨）

```bash
# GitHub CLIをインストール（まだの場合）
# Windows: winget install GitHub.cli
# WSL/Linux: curl -fsSL https://cli.github.com/packages/githubcli-archive-keyring.gpg | sudo dd of=/usr/share/keyrings/githubcli-archive-keyring.gpg
# Mac: brew install gh

# ログイン
gh auth login

# リポジトリを作成してプッシュ
gh repo create InhTaxAutoPJ --public --source=. --remote=origin --push
```

### 1-2. 手動でGitHubリポジトリを作成する方法

1. [GitHub.com](https://github.com/new)にアクセス
2. Repository name: `InhTaxAutoPJ`
3. Public を選択
4. "Create repository" をクリック
5. 以下のコマンドを実行：

```bash
git remote add origin https://github.com/YOUR_USERNAME/InhTaxAutoPJ.git
git branch -M main
git push -u origin main
```

## ステップ 2: Gemini API キーの取得

1. [Google AI Studio](https://makersuite.google.com/app/apikey)にアクセス
2. "Create API Key" をクリック
3. APIキーをコピー（後で使用）

## ステップ 3: Railway でバックエンドをデプロイ

### 3-1. Railwayプロジェクトの作成

1. [Railway.app](https://railway.app/)にアクセスしてログイン
2. "New Project" をクリック
3. "Deploy from GitHub repo" を選択
4. GitHubアカウントを連携（初回のみ）
5. "InhTaxAutoPJ" リポジトリを選択

### 3-2. 環境変数の設定

Railwayダッシュボードで:

1. プロジェクトを開く
2. "Variables" タブをクリック
3. "RAW Editor" をクリック
4. 以下を貼り付け（YOUR_GEMINI_API_KEY を実際のAPIキーに置き換え）：

```env
GEMINI_API_KEY=YOUR_GEMINI_API_KEY
PORT=8000
PYTHON_VERSION=3.11
```

5. "Save Variables" をクリック

### 3-3. デプロイの確認

1. "Deployments" タブで進行状況を確認
2. デプロイ完了後、"Settings" タブで公開URLを確認
3. 例: `https://inhtaxautopj-production.up.railway.app`

## ステップ 4: Cloudflare Pages でフロントエンドをデプロイ

### 4-1. Cloudflare Pagesプロジェクトの作成

1. [Cloudflare Dashboard](https://dash.cloudflare.com/)にログイン
2. "Workers & Pages" → "Create application" → "Pages" タブ
3. "Connect to Git" を選択
4. GitHubアカウントを認証
5. "InhTaxAutoPJ" リポジトリを選択

### 4-2. ビルド設定

以下の設定を入力:

- **Project name**: `inhtaxautopj`
- **Production branch**: `main`
- **Build command**: （空欄のまま）
- **Build output directory**: `frontend`
- **Root directory (Advanced)**: `/frontend`

"Save and Deploy" をクリック

### 4-3. バックエンドURLの更新

デプロイ完了後:

1. RailwayのURLをコピー（例: `https://inhtaxautopj-production.up.railway.app`）
2. GitHubで `frontend/config.js` を編集:

```javascript
const config = {
    API_BASE_URL: window.location.hostname === 'localhost'
        ? 'http://localhost:8000/api'
        : 'https://YOUR-RAILWAY-URL.up.railway.app/api'  // ここを更新
};
```

3. 変更をコミット・プッシュ
4. Cloudflare Pagesが自動的に再デプロイ

## ステップ 5: GitHub Actions の設定

### 5-1. Cloudflare API トークンの作成

1. [Cloudflare Dashboard](https://dash.cloudflare.com/profile/api-tokens)
2. "Create Token" → "Custom token" を選択
3. 権限を設定:
   - Cloudflare Pages:Edit
   - Zone:Read
4. トークンをコピー

### 5-2. GitHub Secrets の設定

GitHubリポジトリで:

1. "Settings" → "Secrets and variables" → "Actions"
2. "New repository secret" で以下を追加:
   - `CLOUDFLARE_API_TOKEN`: 上記で作成したトークン
   - `CLOUDFLARE_ACCOUNT_ID`: CloudflareダッシュボードのアカウントID

## ステップ 6: 動作確認

### 6-1. バックエンド確認

```bash
curl https://YOUR-RAILWAY-URL.up.railway.app/api/health
```

期待される応答:
```json
{"status": "healthy", "version": "1.0.0"}
```

### 6-2. フロントエンド確認

ブラウザで `https://inhtaxautopj.pages.dev` にアクセス

### 6-3. 統合テスト

1. フロントエンドでファイルをアップロード
2. 処理が正常に完了することを確認
3. CSVダウンロードをテスト

## トラブルシューティング

### Railway でビルドが失敗する場合

1. Pythonバージョンを確認（3.11を使用）
2. requirements.txtの依存関係を確認
3. 環境変数が正しく設定されているか確認

### Cloudflare Pages でページが表示されない場合

1. ビルド出力ディレクトリが`frontend`になっているか確認
2. index.htmlが存在するか確認
3. デプロイログでエラーを確認

### CORSエラーが発生する場合

1. backend/core/config.pyのCORS_ORIGINSを更新:
```python
CORS_ORIGINS: List[str] = [
    "https://inhtaxautopj.pages.dev",
    "https://YOUR-ACTUAL-DOMAIN.pages.dev"
]
```

2. 変更をコミット・プッシュ
3. Railwayが自動的に再デプロイ

## 完了チェックリスト

- [ ] GitHubリポジトリ作成完了
- [ ] Gemini APIキー取得完了
- [ ] Railwayバックエンドデプロイ完了
- [ ] Cloudflare Pagesフロントエンドデプロイ完了
- [ ] frontend/config.js のURL更新完了
- [ ] GitHub Actions設定完了
- [ ] ヘルスチェック確認完了
- [ ] 統合テスト完了

すべて完了したら、システムは本番環境で稼働開始です！