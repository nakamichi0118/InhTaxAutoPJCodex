# Cloud Run SQLite 永続化ガイド（GCS Volume Mount）

## 概要

Cloud Run はエフェメラルストレージのため、デプロイ・再起動のたびに `data/analytics.db` / `data/ledger.db` が消える。
GCS バケットをコンテナにマウントすることで SQLite ファイルを永続化する。

---

## 1. GCS バケット作成（初回のみ）

```bash
gcloud storage buckets create gs://sorobocr-data-dxs-pj \
  --project=dxs-pj \
  --location=asia-northeast1 \
  --uniform-bucket-level-access
```

---

## 2. Cloud Run サービスアカウントへの権限付与（初回のみ）

```bash
gcloud projects add-iam-policy-binding dxs-pj \
  --member="serviceAccount:441106995152-compute@developer.gserviceaccount.com" \
  --role="roles/storage.objectUser"
```

---

## 3. Volume Mount 付きデプロイ

```bash
gcloud run deploy sorobocr \
  --source=. \
  --region=asia-northeast1 \
  --project=dxs-pj \
  --max-instances=1 \
  --memory=2Gi \
  --timeout=600 \
  --allow-unauthenticated \
  --add-volume=name=db-data,type=cloud-storage,bucket=sorobocr-data-dxs-pj \
  --add-volume-mount=volume=db-data,mount-path=/mnt/data \
  --set-env-vars=ANALYTICS_DB_PATH=/mnt/data/analytics.db,LEDGER_DB_PATH=/mnt/data/ledger.db
```

初回マウント時はバケットが空なので DB ファイルは存在しない。
アプリ起動時に `AnalyticsStore._initialize()` / `LedgerStore._initialize()` が自動で DB を作成する。

---

## 4. WAL モードと GCS Fuse に関する注意点

- `backend/app/analytics/store.py` では `PRAGMA journal_mode=WAL;` を設定済み
- GCS Fuse 上の SQLite WAL モードは **`--max-instances=1` の場合のみ** 安全に動作する
- 複数インスタンスが同一 GCS ファイルに並行書き込みすると WAL ジャーナルが壊れる
- **`--max-instances=1` は必須設定** — チャンクアップロードの一貫性のためにも必要

---

## 5. Railway からのデータ移行手順

### 5-1. Railway 側でエクスポート

```bash
# Railway CLI をインストール済みであること
railway login
railway link   # プロジェクトとサービスを選択

# analytics.db の全ログを JSONL として出力
railway run python scripts/export_railway_logs.py > /tmp/railway_logs.jsonl

# 件数確認（stderrに出力される）
# 例: Exported 13508 rows
```

### 5-2. ローカルからCloud Runへインポート

```bash
# 基本的な使い方
python scripts/import_to_cloudrun.py /tmp/railway_logs.jsonl

# パスワードや URL を明示する場合
python scripts/import_to_cloudrun.py /tmp/railway_logs.jsonl \
  --base-url https://sorobocr-441106995152.asia-northeast1.run.app \
  --password sorobocr2024 \
  --batch-size 500
```

重複判定は `timestamp + endpoint + ip_address` の組み合わせで行うため、
同じデータを複数回投入しても重複は自動スキップされる（冪等性あり）。

---

## 6. 動作確認

```bash
# ヘルスチェック
curl https://sorobocr-441106995152.asia-northeast1.run.app/api/ping

# インポートエンドポイントのテスト（1件）
curl -X POST \
  "https://sorobocr-441106995152.asia-northeast1.run.app/api/analytics/import?password=sorobocr2024" \
  -H "Content-Type: application/json" \
  -d '[{"timestamp":"2026-01-01T00:00:00","endpoint":"/test","method":"GET","status_code":200}]'
# 期待レスポンス: {"inserted":1,"skipped":0,"total":1}
```
