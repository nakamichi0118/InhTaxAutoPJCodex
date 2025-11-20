# Ledger Frontend Progress Log

## 2025-11-20 16:56 JST
- SOROBOCR風ヒーロー・統合UIをReact側で再構築。名義人と口座名を個別に表示し、案件/口座/取引の指標をヒーローに表示。
- TransactionTable/Integratedタブへ汎用ソート機能を実装。並び替えモードと列ヘッダ切替、手動順序との排他制御を追加。
- 取引タブにもローカルソートを追加し、入出金の昇降順切替に対応。
- 使い方モーダル+`public/guide.html` を追加し、pending import〜PDF出力までの手順とFAQを整理。
- `npm run build` を実行し、`webapp/ledger/`へ最新成果物を配置済み。
