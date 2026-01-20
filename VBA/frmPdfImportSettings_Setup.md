# PDF取込設定フォーム セットアップガイド

## 概要

このUserFormは、PDF取込時の複数のダイアログを1つのフォームに統合します。

## フォームの作成手順

### 1. Excel VBAエディタでUserFormを作成

1. Excel VBAエディタを開く（Alt + F11）
2. プロジェクトを右クリック → 挿入 → ユーザーフォーム
3. フォーム名を `frmPdfImportSettings` に変更

### 2. フォームプロパティの設定

| プロパティ      | 値                   |
| --------------- | -------------------- |
| Name            | frmPdfImportSettings |
| Caption         | PDF取込設定          |
| Width           | 360                  |
| Height          | 320                  |
| StartUpPosition | 1 - CenterOwner      |

### 3. コントロールの配置

以下のコントロールを配置してください：

#### 書類タイプ（Frame + OptionButton）

**フレーム1**

- Name: fraDocType
- Caption: 書類タイプ
- Top: 12
- Left: 12
- Width: 156
- Height: 60

**OptionButton1（フレーム内）**

- Name: optDocTypeBank
- Caption: 通帳
- Top: 14　　
- Left: 12
- Width: 60

**OptionButton2（フレーム内）**

- Name: optDocTypeHistory
- Caption: 取引履歴
- Top: 36
- Left: 70
- Width: 80

#### 日付形式（Frame + OptionButton）

**フレーム2**

- Name: fraDateFormat
- Caption: 日付形式
- Top: 96
- Left: 12
- Width: 156
- Height: 60

**OptionButton3（フレーム内）**

- Name: optDateFormatAuto
- Caption: 自動
- Top: 14
- Left: 12
- Width: 48

**OptionButton4（フレーム内）**

- Name: optDateFormatWareki
- Caption: 和暦
- Top: 14
- Left: 60
- Width: 48

**OptionButton5（フレーム内）**

- Name: optDateFormatWestern
- Caption: 西暦
- Top: 14
- Left: 108
- Width: 48

#### 最小金額（Label + TextBox）

**Label1**

- Name: lblMinAmount
- Caption: 最小金額（円）：
- Top: 84
- Left: 12
- Width: 84

**TextBox1**

- Name: txtMinAmount
- Top: 84
- Left: 102
- Width: 120
- Text: 500000

#### 開始日（Label + TextBox）

**Label2**

- Name: lblStartDate
- Caption: 開始日（任意）：
- Top: 114
- Left: 12
- Width: 84

**TextBox2**

- Name: txtStartDate
- Top: 114
- Left: 102
- Width: 120
- Text: （空）

**Label3**

- Name: lblStartDateHint
- Caption: YYYY-MM-DD
- Top: 114
- Left: 228
- Width: 72
- ForeColor: &H808080

#### 終了日（Label + TextBox）

**Label4**

- Name: lblEndDate
- Caption: 終了日（任意）：
- Top: 144
- Left: 12
- Width: 84

**TextBox3**

- Name: txtEndDate
- Top: 144
- Left: 102
- Width: 120
- Text: （空）

**Label5**

- Name: lblEndDateHint
- Caption: 空欄＝最新まで
- Top: 144
- Left: 228
- Width: 84
- ForeColor: &H808080

#### ボタン

**CommandButton1**

- Name: cmdOK
- Caption: 取込開始
- Top: 192
- Left: 72
- Width: 96
- Height: 30
- Default: True

**CommandButton2**

- Name: cmdCancel
- Caption: キャンセル
- Top: 192
- Left: 180
- Width: 96
- Height: 30
- Cancel: True

### 4. コードの貼り付け

フォームをダブルクリックしてコードウィンドウを開き、`frmPdfImportSettings.frm` ファイルのコードを貼り付けてください。

### 5. 動作確認

VBAエディタでF5キーを押してフォームをテスト実行できます。

## 使用方法

```vba
Dim frm As frmPdfImportSettings
Set frm = New frmPdfImportSettings
frm.Show

If frm.Cancelled Then
    ' ユーザーがキャンセルした
    Exit Sub
End If

' 取得した値を使用
Debug.Print "書類タイプ: " & frm.DocType
Debug.Print "日付形式: " & frm.DateFormat
Debug.Print "最小金額: " & frm.MinAmount
Debug.Print "開始日: " & frm.StartDate
Debug.Print "終了日: " & frm.EndDate

Unload frm
```
