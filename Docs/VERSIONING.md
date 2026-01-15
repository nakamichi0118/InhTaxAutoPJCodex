# バージョン管理規約

## セマンティックバージョニング

本プロジェクトは [Semantic Versioning 2.0.0](https://semver.org/lang/ja/) に準拠します。

```
MAJOR.MINOR.PATCH
```

| 区分 | 更新タイミング | 例 |
|------|--------------|-----|
| MAJOR | 後方互換性のない変更（API破壊的変更、データ形式変更） | 1.0.0 → 2.0.0 |
| MINOR | 後方互換性のある機能追加 | 0.8.0 → 0.9.0 |
| PATCH | バグ修正、軽微な改善 | 0.9.0 → 0.9.1 |

## バージョン管理ファイル

| ファイル | 用途 |
|---------|------|
| `VERSION` | 現在のバージョン番号（単一行） |
| `Docs/CHANGELOG.md` | 変更履歴 |
| `backend/app/main.py` | FastAPIアプリのバージョン定義 |

## リリース命名規則

### ブランチ名
- 機能追加: `feature/<機能名>` (例: `feature/nayose-parser`)
- バグ修正: `fix/<問題名>` (例: `fix/date-parsing`)
- リリース: `release/v<バージョン>` (例: `release/v0.9.0`)

### コミットメッセージ
```
<タイプ>: <概要>

<詳細説明（任意）>
```

| タイプ | 用途 |
|--------|------|
| `feat` | 新機能追加 |
| `fix` | バグ修正 |
| `docs` | ドキュメント変更 |
| `refactor` | リファクタリング |
| `test` | テスト追加・修正 |
| `chore` | ビルド・設定変更 |

### タグ
リリース時は `v<バージョン>` 形式でタグ付け:
```bash
git tag -a v0.9.0 -m "Release v0.9.0: 名寄帳対応"
git push origin v0.9.0
```

## 機能別バージョン対応表

| バージョン | 主要機能 |
|-----------|---------|
| v0.8.x | 預貯金（通帳・取引履歴）対応 |
| v0.9.x | 名寄帳（土地・家屋）対応 |
| v0.10.x | 有価証券対応（予定） |
| v0.11.x | 保険対応（予定） |
| v1.0.0 | 全財産種別対応・正式リリース |

## 更新手順

1. `VERSION` ファイルを更新
2. `backend/app/main.py` のバージョンを更新
3. `Docs/CHANGELOG.md` に変更内容を記載
4. コミット & プッシュ
5. タグ付け（リリース時のみ）

```bash
# 例: v0.9.0 リリース
echo "0.9.0" > VERSION
# main.py のバージョンを更新
git add -A
git commit -m "feat: 名寄帳（土地・家屋）対応を追加"
git push origin main
git tag -a v0.9.0 -m "Release v0.9.0"
git push origin v0.9.0
```
