# 変更履歴 (CHANGELOG)

本ドキュメントは [Keep a Changelog](https://keepachangelog.com/ja/1.0.0/) に準拠します。

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
| 0.9.0 | 2026-01-16 | 名寄帳（土地・家屋）対応 |
| 0.8.0 | - | 預貯金対応（初回リリース） |
