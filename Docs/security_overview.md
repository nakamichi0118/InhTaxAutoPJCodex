## セキュリティ構成サマリー（SOROBOCR）

### 1. 全体アーキテクチャ
- **クライアント**: Cloudflare Pages にホストした SPA（`webapp/`）。PDF をブラウザで読み込み、Railway の API に直接 POST。
- **API 層**: Railway 上の FastAPI (`backend/app`)。PDF を chunking して Gemini / Azure Form Recognizer に連携し、抽出結果を CSV へ整形。
- **外部 AI サービス**:
  - *Google Gemini*: OCR テキスト補強とバックアップ抽出。
  - *Azure Form Recognizer*: テーブル解析の一次ソース。
- **データ保管**: 一時的なメモリのみ。永続 DB やファイルストレージは使用せず、返信後に PDF/抽出結果は破棄。

### 2. データフロー
1. 利用者が Cloudflare Pages の UI から PDF を選択。
2. `fetch` で Railway API (`/api/documents/analyze`) に送信（TLS）。CORS は現在 `*` だが、運用上は CF Pages ドメインに絞ることを推奨。
3. FastAPI が PDF をページ/バイト数で分割 → Gemini/Azure へ HTTPS で送信。
4. 抽出結果を統合・後処理し、CSV 文字列を生成してフロントへ返却（Base64）。
5. ブラウザでプレビュー表示 / CSV ダウンロード。サーバ側にはファイルを保存しない。

### 3. 機密情報・キー管理
- `.env` と Railway Secrets で `AZURE_FORM_RECOGNIZER_*`, `GEMINI_API_KEYS`, `GEMINI_MODEL` 等を設定。Git にはコミットしない。
- Gemini は複数キーのフェイルオーバー実装済み：403（leaked 等）時は自動で次キーに切り替え。
- ログ (`.venv/logs/*.log`) には API 応答コード・例外のみを記録し、PDF 内容や個人情報は出力しない設計。

### 4. 実装済みセキュリティ対策
- **入力制限**: PDF 以外をアップロードした場合は 400。ページサイズ上限（6 MB など）を超えると 413 で拒否。
- **外部通信**: すべて HTTPS。Gemini/Azure 以外へ外部コールしないため SSRF リスク低。
- **CORS**: FastAPI の `CORSMiddleware` で全オリジン許可。PoC では問題ないが、本番は限定予定。
- **XSS 対策**: フロントの表示値は `escapeHtml()` でサニタイズ。
- **秘密情報**: サーバコードでのみ参照。フロントへの露出なし。
- **Abort/キャンセル**: ユーザーキャンセル時は AbortController で即座に API 呼び出しを終了し、不要な処理を行わない。

### 5. 既知リスクと推奨事項
| 項目 | 現状 | 推奨措置 |
| --- | --- | --- |
| 認証/認可 | API はオープン。社内 PoC 想定 | Cloudflare Pages → Railway 間の IP 制限 / API Key / OAuth を導入 |
| CORS | `allow_origins=["*"]` | 本番ドメインのみ許可。 |
| アップロードファイルのマルウェア対策 | 未実装（OCR専用前提） | AV スキャン or サンドボックスを追加。 |
| ログ監査 | FastAPI / Railway ログのみ | 監査用に「誰がどのファイルを何時処理したか」を別途記録。 |
| データ保持 | メモリのみ | 運用手順で「API応答後はファイルを保存しない」ことを明文化。 |

### 6. インシデント対応指針
1. 異常ログ（403/500 連発など）を検知したら即座に該当 API キーを失効。
2. Railway の最新デプロイを停止し、Cloudflare Pages からのアップロードを一時遮断。
3. ログの確認時も個人情報を含む payload を出力しない方針を徹底。

### 7. 今後の拡張／ToDo
- API トークン認証の導入、CF Workers などでのプリサイン。
- ファイル暗号化・署名付きダウンロード（利用履歴の監査）。
- Secrets Rotation の自動化（例：Google Secret Manager / Azure Key Vault 連携）。
- SOC2/ISMS を視野に入れた運用ドキュメント（権限管理、変更管理、ログ保管ポリシー）。

本ドキュメントを基に、セキュリティ顧問には構成図とあわせて上記ポイントを説明してください。
