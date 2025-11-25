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

## 2025-11-20 18:10 JST
- ヘッダー/背景/アクセントカラーをSOROBOCR本体に合わせ、ダークネイビーベース＋白テキストで統一。説明文も本家に合わせたトーンへ調整。
- 口座編集モーダルを追加し、名義・口座名・番号を直接修正できるようにした（`/accounts/{id}` PATCH 連携）。
- 比較パネルを統合タブへ実装。複数口座を選択→入出金/差額/最終取引日を並べて確認でき、選択した口座でリストを絞り込む操作もワンクリック化。
- フィルターUIを比較選択と連動させ、「比較中の口座のみ」オプションやチップ選択を追加。全体の背景も#f1f5f9に合わせて統一感を出した。
- `npm run build` を再実行し、新バンドル (`index-n9tr1DDO.css`, `index-xnLHgbK-.js`) を `webapp/ledger/` に配置。`pytest` も4件パスを確認。

## 2025-11-20 21:12 JST
- 統合タブにAI分析（ヒューリスティック）ボタンを追加。保険会社名の有無・贈与税疑義（個人宛の年間入金110万円超）を自動抽出し、チェックリストで提示。
- 口座比較フィルタを強化し、選択した口座のみでリスト表示／ハイライトできるように調整。比較チップは新カード群と連動。
- すべての改修後 `npm run build` & `pytest` を再実行、`index-v4_1zm5r.css` / `index-PdHtSuWN.js` の最新ビルドを `webapp/ledger/` に反映。

## 2025-11-20 22:08 JST
- LedgerヘッダーをSOROBOCR本体と同じ線形グラデーション/ハイライトへ合わせ、ボタンやコピーもトーンを統一。
- Pending Import時のデフォルト口座名にファイル名(entry.name)を採用し、無名の通帳でも識別しやすく修正。
- `npm run build` を再実行し、新しいバンドル(`index-Bzuudqbz.css` / `index-DgToVgW4.js`)を反映。

## 2025-11-25 16:00 JST
- 取引にタグを付与/編集できるようUIとAPIを拡張。新規・編集フォームにタグ入力(カンマ区切り)を追加し、一覧でもバッジ表示＆タグキーワードでフィルター可能に。
- Ledger DBへ`tags`カラムを追加し、Create/Update/Import/Exportの全経路でタグを保存・復元。既存データは空文字として安全に移行。
- ヘッダーをSOROBOCR本体のグラデーションに揃えつつ、pending import時の口座名はファイル名を既定とするよう調整。
- `npm run build`&`pytest`再実行。新バンドル(`index-Czm85cG5.css`/`index-Dn7NBBUK.js`)を`webapp/ledger/`に反映。
