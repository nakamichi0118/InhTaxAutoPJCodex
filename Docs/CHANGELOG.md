# 変更履歴 (CHANGELOG)

本ドキュメントは [Keep a Changelog](https://keepachangelog.com/ja/1.0.0/) に準拠します。

---

## [0.10.0] - 2026-01-19

### Added（追加）

- **同日取引の順序保持** - OCR読み取り順序を `line_order` フィールドで追跡
  - `TransactionLine` に `line_order`, `account_type` フィールド追加
  - ソートキーに `line_order` を追加し、同日取引の正確な順序を保持

- **読取開始日の指定機能** - 資金移動表の開始日を指定可能に
  - フロントエンド: 通帳・取引履歴タブに「読取開始日」入力フィールド追加
  - バックエンド: `/api/jobs` に `start_date` パラメータ追加
  - 指定日以前の取引を自動的にフィルタリング

- **総合口座の普通預金/定期預金自動分離**
  - Geminiプロンプトに口座種別検出を追加
  - `account_type` に基づいて取引を自動分離
  - 普通預金・定期預金それぞれ別の `AssetRecord` として出力

### Fixed（修正）

- **ゆうちょ銀行の利子・税金処理改善**
  - 利子・税金キーワード検出（利息、源泉税、所得税等）
  - 小額の利子・税金取引で残高計算が誤る問題を修正

- **ゆうちょ銀行摘要の取扱店番号除去**
  - 先頭の4-5桁店番号パターン（例: `03050 1,000,000通帳`）を除去
  - `YUUCHO_BRANCH_PREFIX_PATTERN` 追加

### Changed（変更）

- **カタカナ→漢字変換拡充**
  - 50+の変換パターンを追加（金融機関名、税金関連、公共料金など）
  - 例: ﾌﾘｺﾐ→振込、ﾐﾂｲｽﾐﾄﾓ→三井住友、ﾈﾝｷﾝ→年金 等

- **入出金検討表ツールをβ版表記に変更**
  - ヘッダーリンク、セクションタイトル、ボタンに「β版」ラベルを追加

### Technical Details

- `backend/app/models.py` - `TransactionLine` に `line_order`, `account_type` 追加
- `backend/app/azure_analyzer.py` - ソートキーに `line_order` 追加
- `backend/app/main.py` - `_separate_assets_by_account_type()`, `_filter_transactions_by_start_date()` 追加
- `backend/app/gemini.py` - プロンプトに `account_type` 検出を追加
- `backend/app/job_manager.py` - `start_date` フィールド追加
- `description_utils.py` - ゆうちょパターン、カタカナ変換拡充
- `webapp/index.html` - 開始日UI追加、β版表記追加

---

## [0.9.1] - 2026-01-16

### Fixed（修正）

- 名寄帳ジョブ処理の修正 - `document_type_hint == "nayose"` の場合に専用処理を実行するよう修正
  - 以前: すべてのGeminiジョブが `bank_deposit` として処理されていた
  - 修正後: 名寄帳は `_analyze_nayose_with_gemini` で正しく処理される

### Changed（変更）

- `backend/app/main.py` - `_process_job_record` に名寄帳専用分岐を追加
- `webapp/index.html` - 土地・家屋タブ有効化（v0.9.0でフロントのみ対応済み）
- `CLAUDE.md` - デプロイ後検証手順を追加

---

## [0.9.0] - 2026-01-16

### Added（追加）
- **名寄帳（土地・家屋）対応** - Phase 1 不動産機能
  - `backend/app/prompts/nayose_prompt.py` - 名寄帳解析用Geminiプロンプト
  - `backend/app/parsers/nayose_parser.py` - 名寄帳レスポンスパーサー
  - `tests/test_nayose_parser.py` - ユニットテスト（22件）
- **土地・家屋専用CSV出力**
  - `land.csv` - 土地情報（市区町村、所在地・地番、地目、地積、評価額、持分欄）
  - `building.csv` - 家屋情報（市区町村、所在地・家屋番号、構造、床面積、建築年、評価額、持分欄）
- **仕様書**
  - `Docs/ROADMAP_SPEC.md` - 財産種別拡張ロードマップ
  - `Docs/REAL_ESTATE_MATCHING_SPEC.md` - 不動産マッチング検討資料
  - `Docs/VERSIONING.md` - バージョン管理規約
  - `Docs/CHANGELOG.md` - 変更履歴
- `CLAUDE.md` - Claude Code用プロジェクトガイド

### Changed（変更）
- `backend/app/models.py` - DocumentTypeに`nayose`追加
- `backend/app/parser.py` - 名寄帳書類の自動検出ロジック追加
- `backend/app/gemini.py` - `extract_nayose_from_pdf`メソッド追加
- `backend/app/main.py` - 名寄帳用ルーティング追加
- `backend/app/exporter.py` - 土地・家屋CSV出力対応
- `src/export_csv.py` - `LAND_EXPORT_COLUMNS`, `BUILDING_EXPORT_COLUMNS` 追加
- `VBA/transaction_import.bas` - 検索行数を1000→10000に拡張

### Design Decisions（設計判断）
- 名寄帳のみ方式を採用（登記簿とのマッチングは見送り）
- 持分・登記地目はExcel側で手入力とする
- 名寄帳から100%取得可能な項目: 所在地、地番、地目（課税）、地積、評価額、構造、建築年

---

## [0.8.0] - 2026-01-XX（以前のリリース）

### Features
- 預貯金（通帳・取引履歴）対応
- Gemini AI OCR による取引抽出
- 日付推論エンジン（DateInferenceEngine）
- VBA取込マクロ（transaction_import.bas）
- FastAPI非同期ジョブシステム
- UTF-8 BOM CSV出力

---

## バージョン履歴サマリー

| バージョン | リリース日 | 主要機能 |
|-----------|-----------|---------|
| 0.10.0 | 2026-01-19 | 同日順序保持、期間指定、総合口座分離、ゆうちょ対応改善 |
| 0.9.1 | 2026-01-16 | 名寄帳ジョブ処理修正 |
| 0.9.0 | 2026-01-16 | 名寄帳（土地・家屋）対応 |
| 0.8.0 | - | 預貯金対応（初回リリース） |
