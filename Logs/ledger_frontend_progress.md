# Ledger Frontend Progress Log

## 2025-11-20 16:56 JST
- SOROBOCR風ヒーロー・統合UIをReact側で再構築。名義人と口座名を個別に表示し、案件/口座/取引の指標をヒーローに表示。
- TransactionTable/Integratedタブへ汎用ソート機能を実装。並び替えモードと列ヘッダ切替、手動順序との排他制御を追加。
- 取引タブにもローカルソートを追加し、入出金の昇降順切替に対応。
- 使い方モーダル+`public/guide.html` を追加し、pending import〜PDF出力までの手順とFAQを整理。
- `npm run build` を実行し、`webapp/ledger/`へ最新成果物を配置済み。

## 2025-11-20 17:16 JST
- ヘッダーや注意書きを既存SOROBOCRテイストへ戻し、グラデーション/サマリーカード/余計なAPI説明を撤去。ガイドはシンプルなボタンとモーダルでアクセスできるよう整理。
- Job Previewに「統合グループキー」オプションを追加し、同じキーを与えた複数口座を1口座として取り込めるようフロント/バックエンド両方を拡張。
- `LedgerJobImportMapping` と `/jobs/{id}/import` を改修し、グループ単位でまとめて口座作成＆取引一括投入が可能に。UIにも説明追記。
- `npm run build` を再実行し、新しいバンドル(`index-bR01U9YX.css` / `index-DbGv-S05.js`)を `webapp/ledger/` に配置。

## 2025-11-20 18:00 JST
- `/api/ledger/import` 置換問題の対策として、Pending Import時はフロント側で既存口座/取引と読み取り結果をマージしてから送信するよう変更。口座ID競合時はUUIDを自動採番して重複を避ける。
- 統合タブにフィルターUIを追加（口座/入出金区分/色/金額レンジ/キーワード）。フィルタ結果に対してソート・PDF出力が行われるよう組み合わせ処理も更新。
- `npm run build` を再実行し、フィルター入りのUIと差し替え済みバンドル (`index-CwhQkWVL.css`, `index-C_GWTxPS.js`) を `webapp/ledger/` へ配置。
