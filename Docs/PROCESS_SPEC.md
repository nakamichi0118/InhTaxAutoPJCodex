# SOROBOCR 処理仕様書 (v0.8)

本ドキュメントは、SOROBOCR（InhTaxAutoPJ Codex）の現行処理フローをフロントエンドからCSV出力まで詳細にまとめた仕様書です。関係者・新規メンバーがこれを読むだけで全体像と各処理の役割を把握できるよう、構成・通信・演算ロジックを段階的に記載しています。

---

## 1. 全体概要

| 項目 | 内容 |
| --- | --- |
| 主な入力 | PDF形式の通帳/取引履歴（1口座=1ファイル） |
| 主な出力 | `assets.csv`（資産情報）、`bank_transactions.csv`（取引明細） |
| 解析エンジン | Gemini 2.5 Pro（標準） / Gemini 2.5 Flash（高速） |
| 対応ページ・サイズ | 目安: 1ファイル 20MB / 100ページまで |
| 平均処理速度 | 1取引行あたり約0.6〜1.8秒（100行で1〜3分） |
| 保存ポリシー | PDFは処理後即時削除。CSVはブラウザ経由でのみ取得し、サーバーには保管しない |

---

## 2. コンポーネント構成

1. **Webフロント (webapp/index.html)**
   - 利用者がPDFをドロップ/選択し、解析エンジンを指定するUI。
   - `fetch`でRailway上のFastAPI (`/api` 配下) にリクエストを送信。
   - 処理進捗、完了時間、件数を表示し、CSVをダウンロードさせる。

2. **FastAPIバックエンド (backend/app/main.py)**
   - `/api/jobs`でジョブを受け付け、`JobManager`へ登録。
   - `_process_job_record`がバックグラウンドでPDF解析→取引補正→CSV生成を実行。
   - `/api/jobs/{id}`で進捗確認、`/api/jobs/{id}/result`で最終成果物を返却。

3. **ジョブ管理 (backend/app/job_manager.py)**
   - 受け取ったPDFを一時ファイルに保存し、UUIDベースの`job_id`を発行。
   - 各ジョブは専用スレッドで`_process_job_record`を実行。
   - ジョブ情報は`JobRecord`（ステータス/進捗/結果ファイル等）で保持。

4. **Gemini解析 (backend/app/gemini.py + main.py内ヘルパー)**
   - Gemini 2.5 Pro/FlashへPDFページを投げ、テキスト行・構造化取引を取得。
   - 4スレッドまでの`ThreadPoolExecutor`でページ並列解析し、全件を結合。

5. **後処理・CSV出力 (backend/app/main.py & exporter.py)**
   - `_enforce_continuity`などの補正ロジックで入出金/残高の矛盾を解消。
   - `export_to_csv_strings`でBOM付きCSV文字列を生成し、Base64化して返却。

---

## 3. リクエスト〜CSV出力の詳細フロー

### 3.1 Webフロントからの処理
1. 利用者がタブ「通帳・取引履歴」を選択し、解析エンジン（Pro/Flash）と日付形式を指定。
2. `FormData`にPDFとメタ情報を格納し、`POST /api/jobs`を実行。
3. 受領した`job_id`を`/api/jobs/{id}`でポーリングし、`status=completed`で`/api/jobs/{id}/result`からファイル群を取得。
4. フロントでは処理時間・名義人・件数を表示し、CSVダウンロードボタンを提供。

### 3.2 FastAPI受信〜ジョブ登録
1. `/api/jobs`は以下を受け取る:
   - `file`: PDF
   - `document_type`: 任意のタイプ（取引履歴/通帳など）
   - `date_format`: `auto / western / wareki`
   - `processing_mode`: 常に`gemini`に強制（v0.8より）。他値は警告ログを出し`gemini`へフォールバック。
   - `gemini_model`: `gemini-2.5-pro`または`gemini-2.5-flash`
2. 入力検証後、`JobManager.submit`へ引き渡し、`job_id`を返却。
3. `JobManager`は一時ファイルを作成、`JobRecord`を初期化し、専用スレッドで`_process_job_record`を走らせる。

### 3.3 `_process_job_record` の処理手順
1. **初期化**
   - PDFを読み込み、`JobHandle`で`stage="queued"`→`"analyzing"`へ更新。
   - `PdfChunkingPlan(max_pages=1, max_bytes=settings.gemini_max_document_bytes)`でチャンク分割。

2. **Geminiページ並列処理**
   - `ThreadPoolExecutor(max_workers=min(4, ページ数))`で各ページを`_analyze_page_with_gemini`へ投入。
   - 1ページ毎にGeminiへ2.5 Pro/Flash呼び出し (`chunk_page_limit_override=1`)。
   - 返却された`GeminiExtraction`を `_convert_gemini_structured_transactions` or `build_transactions_from_lines`で`TransactionLine`へ変換。
   - 並列実行の完了順で結果を受け取り、ページ番号でソート→連結。
   - 進捗 (`processed_chunks / total_chunks`) をハンドル経由で更新。

3. **補正・整合パイプライン**
   1. `_enforce_continuity`: 残高差と摘要キーワードから入出金欄を補正し、必要に応じて残高を補完。
   2. `_finalize_transaction_directions`: 残高推移を基準に入出金方向を再判定。
3. `post_process_transactions`: 日付/摘要の整形（AI出力のカタカナ→常用表現変換、数字のみの括弧除去、`手数料(1)`→`手数料`統合 など）を行う。
   4. `_finalize_transactions_from_balance`: 最終残高を絶対値とし、入出金金額の矛盾を再度解消（残高は変更しない）。

4. **CSV/JSON生成・レスポンス作成**
   - `export_to_csv_strings`が`assets.csv`と`bank_transactions.csv`を生成。
   - `bank_transactions.json`は `{"version": "2.0", "exported_at": "...", "accounts": [...], "transactions": [...]}` 形式で、各取引には `transaction_date/date`・`description/memo`・`withdrawal(_amount)`・`deposit(_amount)`・`accountId` などを同居させ、Excel/入出金検討表ツールの双方が同じファイルを参照できる。
   - すべてのCSV/JSONをUTF-8(BOM付き)文字列→Base64エンコードし、`result_files`として保存。
   - `handle.update`で`status="completed"`、`processed_chunks=total_chunks`、`assets_payload`などをセット。
   - `JOB_TOTAL`などのタイミングログを`logger.info("TIMING|...")`で出力。

5. **後処理**
   - `JobManager`はジョブ終了後に一時ファイルを削除。
   - `JobResultResponse`には`document_type`と`files`（Base64 CSV）、`assets`（構造化JSON）を含める。

---

## 4. 補正ロジックの詳細

| 関数 | 主な処理 | 効果 |
| --- | --- | --- |
| `_enforce_continuity(prev_balance, transactions)` | 摘要キーワード/残高差/前行残高を利用し、入出金欄の入れ替え・再算出・残高補完を実施。 | 摘要由来の誤判定や空欄を早期に補正 |
| `_finalize_transaction_directions(transactions)` | 各行の残高差から入出金方向を再判定。必要に応じて再設定メモを追記。 | 残高を真値として入出金の符号を確定 |
| `post_process_transactions(transactions)` | 日付整形・テキスト正規化（半角→全角、カタカナ略語→常用語、手数料（1/2）→手数料 など） | CSV整形前に項目をクリーンに保つ |
| `_finalize_transactions_from_balance(transactions)` | 前行との差額と入出金欄だけを照合し、矛盾時には金額を入れ替える／再設定する。残高は一切変更しない。 | 最終的に残高が常に連続する状態を保証 |

補正メモ (`correction_note`) には「残高差から入金額を再算出」「入出金欄を残高整合の結果として入れ替えました」などの履歴を残し、CSVで可視化できるようにしています。

---

## 5. セキュリティ・保守方針

1. **ファイル保存**
   - PDFはユニークな一時ファイルとして保存し、ジョブ終了後に即削除。
   - CSVはBase64文字列としてAPIレスポンスで返却するのみ。サーバー側に永続保存しない。

2. **環境変数 & 設定**
   - `backend/app/config.py`で`.env`を読み込み、`GEMINI_API_KEYS`, `GEMINI_MODEL`, `GEMINI_DOCUMENT_MAX_MB`などを設定。
   - Azureキーはv0.8では未使用だが、互換性のため設定項目は残っている。

3. **ログ**
   - `logger.info("TIMING|job_id|component|page|duration_ms")`形式でPDF分割/各ページGemini処理/補正/CSV出力などの所要時間を出力し、ボトルネック解析に活用。
   - 重大な例外は`handle.update(status="failed")`経由でユーザにも通知。

4. **再処理時の推奨手順**
   - ページを再読み込み→PDFを再度アップロードすることで、通信・一時混雑による失敗を改善できるケースが多い。
   - 巨大ファイルは複数回に分割することで各ジョブの安定性が向上する。

---

## 6. 開発・計測メモ

- ローカル検証は`uvicorn backend.app.main:app --reload`で起動。`run_server.sh`と`benchmark.py`を利用するとAPIベンチマーク（平均・p95・Gemini/Azure各フェーズの所要時間）が採取できる。
- `benchmark_results/`には最近の計測結果 (`benchmark_detail.json`, `benchmark_summary.csv`) を保存済み。
- コード改修時は`AGENTS.md`のProgress Logに必ず対応内容を記載し、変更が未検証の場合はその旨を明記すること。

---

## 7. よくある質問（開発向け）

1. **なぜGemini単独構成に統一したのか？**  
   Azure+Geminiハイブリッドは残高精度が低く、最終的に入出金調整にも影響したため。Gemini 2.5 Pro単独＋Parallel Page処理で速度を確保しつつ精度99%を維持できた。

2. **処理時間を短縮するには？**  
   1ページ1チャンク＋4スレッドでGemini呼び出しを並列化することで、ページ数に対してほぼ一定時間で処理可能。更なる改善にはGemini API側のバッチ/並列度調整が対象となる。

3. **CSV整形で留意すべき点は？**  
   - BOM付きUTF-8でエクスポートし、Excelで文字化けしないよう保持。
   - `correction_note`は後工程の監査に利用されるため、補正内容を必ず追記する。

この仕様書は v0.8 時点の内容に基づきます。機能追加・構成変更が行われた場合は、同ファイルを更新し、`AGENTS.md`にも履歴を残してください。

---

## 8. 入出金検討表ツール (ledger_frontend)

- `ledger_frontend/` は React + Vite で構築された帳票ツールで、`npm run build` すると `webapp/ledger/` に静的成果物が生成されます。
- Firebase 依存を撤廃し、Railway 上の FastAPI に追加された `/api/ledger/*` REST エンドポイントへ `fetch` で直接アクセスします。
- ブラウザごとに `POST /api/ledger/session` で匿名トークンを取得し、以降のリクエストで `X-Ledger-Token` として送信することでユーザー領域を判別します。
- 口座/取引の取得は `GET /api/ledger/state?case_id=...`、追加・削除は `/api/ledger/accounts*` / `/api/ledger/transactions*`、順序更新は `/reorder` エンドポイントで完結します。
- 案件（case）は `/api/ledger/cases` で管理し、UI の案件切替ドロップダウンから既存案件を選択または新規案件を作成できます。
- OCR完了後は `webapp/index.html` のCTAが `job_id` 付きで `/ledger/?job_id=...` に遷移します。Ledger側では `GET /api/ledger/jobs/{job_id}/preview` で検出口座を取得し、ユーザーが「新規口座として登録」または「既存口座へマージ」を選択→`POST /api/ledger/jobs/{job_id}/import` で取引を案件へ反映します。
- バージョン0.9では `webapp/index.html` から `POST /api/ledger/jobs/{job_id}/import` を自動実行し、ファイル名を案件名とした新規ケースに口座／取引を登録してから `/ledger/?case_id=...` へリンクします。Ledger画面では必要に応じて別案件への移動や手動マージも可能です。
- `LEDGER_DB_PATH`（既定: `data/ledger.db`）で指定されたSQLiteに永続化され、Cloudflare Pages + Railway だけで入出金検討表の保存・復元が可能になりました。
