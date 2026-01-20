Option Explicit

'================ Global Declarations ================
' Windows API declarations (module先頭に配置する必要あり)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private gDebugWs As Worksheet
Private gDebugRow As Long
Private Const DEBUG_LOG_ENABLED As Boolean = True

'===============================================================================
' PDF取込設定の型定義（Typeはモジュール先頭で定義が必要）
'===============================================================================
Public Type PdfImportSettings
    Cancelled As Boolean
    DocType As String
    DateFormat As String
    MinAmount As Long
    StartDate As String
    EndDate As String
End Type

'CSV取込ボタンクリック時のメインプロシージャ
Sub CSV取込ボタン_Click()
    Dim filePath As String
    Dim buttonCol As Long
    Dim buttonRow As Long
    Dim minAmount As Long
    Dim csvData As Variant
    Dim usageData As Variant
    Dim rawText As String
    Dim ws As Worksheet
    Dim callerName As String
    Dim shp As Shape
    Dim retryCount As Integer
    Dim dummy As Long

    ' ブック開直後の初期化待ち（シェイプコレクション準備）
    DoEvents
    Sleep 100
    DoEvents

    Set ws = ActiveSheet

    ' シェイプコレクションを事前にアクセスして初期化を促す
    On Error Resume Next
    dummy = ws.Shapes.Count
    Err.Clear
    On Error GoTo 0

    '1. ボタン位置の取得（リトライ付き）
    For retryCount = 1 To 3
        On Error Resume Next
        callerName = Application.Caller
        If Err.Number = 0 And Len(callerName) > 0 Then
            Set shp = ws.Shapes(callerName)
            If Err.Number = 0 And Not shp Is Nothing Then
                On Error GoTo 0
                Exit For
            End If
        End If
        Err.Clear
        On Error GoTo 0

        ' リトライ前に少し待つ
        DoEvents
        Sleep 200
        DoEvents
    Next retryCount

    If shp Is Nothing Then
        MsgBox "初回セットアップが完了しました。" & vbCrLf & vbCrLf & _
               "再度ボタンを押して実行してください。", vbInformation
        Exit Sub
    End If

    buttonCol = shp.TopLeftCell.Column
    buttonRow = shp.TopLeftCell.Row

    '2. ファイルダイアログを開いてCSVファイルを選択
    filePath = SelectCSVFile()
    If filePath = "" Then
        Exit Sub
    End If

    '3. 金額フィルタの入力
    minAmount = GetMinimumAmount()
    If minAmount = -1 Then
        Exit Sub
    End If

    '4. CSVファイル全体を読み込む
    rawText = ReadUtf8File(filePath)
    If Len(rawText) = 0 Then
        MsgBox "CSVファイルの読み込みに失敗しました。", vbExclamation
        Exit Sub
    End If
    usageData = ParseTransactionCsvContent(rawText, 0)
    csvData = ParseTransactionCsvContent(rawText, minAmount)
    If IsEmpty(csvData) Then
        MsgBox "指定金額以上の取引は見つかりませんでした。", vbInformation
        Exit Sub
    End If

    '5. データをExcelに反映
    Call ImportDataToExcel(csvData, buttonCol, buttonRow, usageData)

    MsgBox "CSV取込が完了しました。", vbInformation

End Sub

'CSVファイル選択ダイアログ
Function SelectCSVFile() As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .TITLE = "CSVファイルを選択してください"
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv", 1
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectCSVFile = .SelectedItems(1)
        Else
            SelectCSVFile = ""
        End If
    End With

    Set fd = Nothing
End Function

'最小金額の入力ダイアログ
Function GetMinimumAmount() As Long
    Dim inputValue As String
    Dim amount As Long

    inputValue = InputBox("取り込む最小金額を入力してください（円単位）" & vbCrLf & _
                         "例：50万円以上の場合は「500000」と入力", _
                         "金額フィルタ", "500000")

    If inputValue = "" Then
        GetMinimumAmount = -1
        Exit Function
    End If

    If Not IsNumeric(inputValue) Then
        MsgBox "数値を入力してください。", vbExclamation
        GetMinimumAmount = -1
        Exit Function
    End If

    GetMinimumAmount = CLng(inputValue)
End Function

'日付形式を変換（YYYY-MM-DD → 和暦形式）
'2桁年号の場合はスマート推論を適用（相続税案件向け）
Function ConvertDateFormat(dateStr As String) As String
    Dim dateParts() As String
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    Dim gengo As String
    Dim warekiYear As Integer

    dateParts = Split(dateStr, "-")
    If UBound(dateParts) <> 2 Then
        ConvertDateFormat = ""
        Exit Function
    End If

    year = CInt(dateParts(0))
    month = CInt(dateParts(1))
    day = CInt(dateParts(2))

    '2桁年号のスマート推論（相続税案件では直近の日付が多い）
    If year < 100 Then
        year = InferFullYear(year)
    End If

    '和暦変換
    If year >= 2019 Then
        gengo = "R"
        warekiYear = year - 2018
    ElseIf year >= 1989 Then
        gengo = "H"
        warekiYear = year - 1988
    Else
        gengo = "S"
        warekiYear = year - 1925
    End If

    ConvertDateFormat = gengo & warekiYear & "(" & year & ")/" & Format(month, "00") & "/" & Format(day, "00")
End Function

'2桁年号から西暦4桁を推論する
'相続税案件では直近数年の取引が多いため、1桁年号は令和と推定
Private Function InferFullYear(shortYear As Integer) As Integer
    Dim currentReiwaYear As Integer
    Dim currentYear As Integer

    currentYear = VBA.year(Now)
    currentReiwaYear = currentYear - 2018  ' 令和何年か（2025年なら令和7年）

    '推論ロジック:
    ' 1-現在の令和年 → 令和（例: 1-7なら令和1-7年 = 2019-2025年）
    ' 8-31 → 平成（H8-H31 = 1996-2019年、令和8年以降は未来なので平成）
    ' 32-64 → 昭和（S32-S64 = 1957-1989年）
    ' 65-99 → 1900年代後半（1965-1999年）として扱う

    If shortYear >= 1 And shortYear <= currentReiwaYear Then
        ' 令和の範囲内（今日が令和7年なら1-7は令和）
        InferFullYear = shortYear + 2018
    ElseIf shortYear <= 31 Then
        ' 平成の範囲（H1-H31 = 1989-2019年）
        ' ただし1-7は上で令和として処理済み
        InferFullYear = shortYear + 1988
    ElseIf shortYear <= 64 Then
        ' 昭和の範囲（S32-S64 = 1957-1989年）
        InferFullYear = shortYear + 1925
    Else
        ' 65-99: 1900年代後半として解釈
        InferFullYear = 1900 + shortYear
    End If
End Function

'データをExcelに反映
Sub ImportDataToExcel(csvData As Variant, buttonCol As Long, buttonRow As Long, usageData As Variant)
    Dim ws As Worksheet
    Dim i As Long
    Dim targetRow As Long
    Dim existingRow As Long
    Dim dateToFind As String
    Dim gyouhajime As Long
    Dim gyousaigo As Long
    Dim retuhajime As Long
    Dim retusaigo As Long
    Dim FoundCell As Range
    Dim insertRow As Long
    Dim emptyFound As Boolean
    Dim summarySource As Variant

    Set ws = Worksheets("預金推移")

    '範囲を取得
    gyouhajime = ws.Range("A1:A10000").Find("資金移動始").row
    gyousaigo = ws.Range("A1:A10000").Find("資金移動終").row
    retuhajime = ws.Range("C1:BZ1").Find("資金移動始").Column
    retusaigo = ws.Range("C1:DZ1").Find("資金移動終").Column

    '画面更新停止
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'データを1件ずつ処理
    For i = 1 To UBound(csvData, 1)
        dateToFind = csvData(i, 1)

        '同じ日付を検索
        Set FoundCell = ws.Range(ws.Cells(gyouhajime + 1, 3), ws.Cells(gyousaigo - 3, 3)).Find(What:=dateToFind, LookAt:=xlWhole)

        If FoundCell Is Nothing Then
            '同じ日付がない場合、新規行を挿入
            insertRow = gyousaigo - 2
            ws.Rows(insertRow).Insert Shift:=xlUp
            ws.Rows(insertRow).Interior.ColorIndex = 0
            ws.Rows(insertRow).NumberFormatLocal = "#,##0"
            ws.Rows(insertRow).Font.Bold = False

            '日付とデータを入力
            ws.Cells(insertRow, 3).value = dateToFind
            Call PopulateTransactionRow(ws, insertRow, buttonCol, csvData(i, 2), csvData(i, 3), csvData(i, 4))

        Else
            '同じ日付がある場合
            existingRow = FoundCell.row
            emptyFound = False

            '同じ日付の行を確認
            Do
                '空白セルをチェック
                If ws.Cells(existingRow, buttonCol).value = "" And ws.Cells(existingRow, buttonCol + 1).value = "" Then
                    '空白セルがある場合はそこに入力
                    Call PopulateTransactionRow(ws, existingRow, buttonCol, csvData(i, 2), csvData(i, 3), csvData(i, 4))

                    emptyFound = True
                    Exit Do
                End If

                '次の行が同じ日付かチェック
                If existingRow < gyousaigo - 3 Then
                    If ws.Cells(existingRow + 1, 3).value = dateToFind Then
                        existingRow = existingRow + 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            '空白セルが見つからなかった場合は新規行を挿入
            If Not emptyFound Then
                insertRow = existingRow + 1
                ws.Rows(insertRow).Insert Shift:=xlUp
                ws.Rows(insertRow).Interior.ColorIndex = 0
                ws.Rows(insertRow).NumberFormatLocal = "#,##0"
                ws.Rows(insertRow).Font.Bold = False

                '日付とデータを入力
                ws.Cells(insertRow, 3).value = dateToFind
                Call PopulateTransactionRow(ws, insertRow, buttonCol, csvData(i, 2), csvData(i, 3), csvData(i, 4))
            End If
        End If
    Next i

    '範囲を再取得してソート
    gyouhajime = ws.Range("A1:A10000").Find("資金移動始").row
    gyousaigo = ws.Range("A1:A10000").Find("資金移動終").row
    retuhajime = ws.Range("C1:BZ1").Find("資金移動始").Column
    retusaigo = ws.Range("C1:DZ1").Find("資金移動終").Column

    With ws
        .Range(.Cells(gyouhajime + 1, retuhajime), .Cells(gyousaigo - 3, retusaigo)).Sort Key1:=.Columns(3)
    End With

    '罫線の更新
    Call UpdateBorders

    '用途サマリをボタン行の1つ上へ表示
    If IsEmpty(usageData) Then
        summarySource = csvData
    Else
        summarySource = usageData
    End If
    Call WriteUsageSummary(summarySource, buttonCol, buttonRow)

    '画面更新再開
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Private Sub PopulateTransactionRow(ws As Worksheet, targetRow As Long, buttonCol As Long, _
    withdrawValue As Variant, depositValue As Variant, descriptionValue As Variant)

    Dim withdrawAmount As Long
    Dim depositAmount As Long
    Dim description As String

    withdrawAmount = ToLongValue(CStr(withdrawValue))
    depositAmount = ToLongValue(CStr(depositValue))
    description = CStr(descriptionValue)

    If withdrawAmount > 0 Then
        ws.Cells(targetRow, buttonCol).value = Format(withdrawAmount, "#,##0")
        ws.Cells(targetRow, buttonCol).HorizontalAlignment = xlRight
    End If

    If depositAmount > 0 Then
        ws.Cells(targetRow, buttonCol + 1).value = Format(depositAmount, "#,##0")
        ws.Cells(targetRow, buttonCol + 1).HorizontalAlignment = xlRight
    End If

    Call ApplyDescriptionToOppositeCell(ws, targetRow, buttonCol, withdrawAmount, depositAmount, description)
End Sub

Private Sub ApplyDescriptionToOppositeCell(ws As Worksheet, targetRow As Long, buttonCol As Long, _
    withdrawAmount As Long, depositAmount As Long, description As String)
    Dim descCol As Long
    Dim trimmedDesc As String

    trimmedDesc = Trim$(description)
    If Len(trimmedDesc) = 0 Then Exit Sub

    If withdrawAmount > 0 And depositAmount <= 0 Then
        descCol = buttonCol + 1
    ElseIf depositAmount > 0 And withdrawAmount <= 0 Then
        descCol = buttonCol
    Else
        Exit Sub
    End If

    If ws.Cells(targetRow, descCol).value = "" Then
        ws.Cells(targetRow, descCol).value = trimmedDesc
        ws.Cells(targetRow, descCol).HorizontalAlignment = xlLeft
    End If
End Sub

'罫線を更新（資金移動.basの処理を参考）
Sub UpdateBorders()
    Dim ws As Worksheet
    Dim gyouhajime As Long
    Dim gyousaigo As Long
    Dim retuhajime As Long
    Dim retusaigo As Long
    Dim i As Long
    Dim currentYear As String
    Dim nextYear As String

    Set ws = Worksheets("預金推移")

    gyouhajime = ws.Range("A1:A10000").Find("資金移動始").row
    gyousaigo = ws.Range("A1:A10000").Find("資金移動終").row
    retuhajime = ws.Range("C1:BZ1").Find("資金移動始").Column
    retusaigo = ws.Range("C1:DZ1").Find("資金移動終").Column

    '横罫線を設定
    ws.Range(ws.Cells(gyouhajime, retuhajime), ws.Cells(gyousaigo, retusaigo)).Borders(xlInsideHorizontal).LineStyle = xlContinuous

    '年度境界に色付き罫線
    For i = gyouhajime + 1 To gyousaigo - 4
        currentYear = ExtractYear(ws.Cells(i, 3).value)
        nextYear = ExtractYear(ws.Cells(i + 1, 3).value)

        If currentYear <> nextYear Then
            With ws.Range(ws.Cells(i, 3), ws.Cells(i, retusaigo)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(255, 0, 0)
                .Weight = xlThin
            End With
        End If
    Next i

    '縦罫線を設定
    ws.Range(ws.Cells(gyouhajime, retuhajime), ws.Cells(gyousaigo, retuhajime)).Borders(xlEdgeLeft).Weight = xlThick
    ws.Range(ws.Cells(gyouhajime, retusaigo), ws.Cells(gyousaigo, retusaigo)).Borders(xlEdgeRight).Weight = xlThick

End Sub

Sub WriteUsageSummary(csvData As Variant, buttonCol As Long, buttonRow As Long)
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim withdrawalSummary As String
    Dim depositSummary As String

    If IsEmpty(csvData) Then Exit Sub
    If buttonRow <= 1 Then Exit Sub

    targetRow = buttonRow - 1
    Set ws = Worksheets("預金推移")

    withdrawalSummary = BuildUsageSummary(csvData, True)
    depositSummary = BuildUsageSummary(csvData, False)

    ApplySummaryToCell ws.Cells(targetRow, buttonCol), withdrawalSummary
    ApplySummaryToCell ws.Cells(targetRow, buttonCol + 1), depositSummary
End Sub

Private Function BuildUsageSummary(csvData As Variant, isWithdrawal As Boolean) As String
    Dim rowCount As Long
    Dim columnCount As Long
    Dim dict As Object
    Dim i As Long
    Dim amount As Long
    Dim descValue As String
    Dim key As String
    Dim keys As Variant
    Dim j As Long
    Dim tempKey As Variant
    Dim parts() As String

    On Error GoTo ExitFunc
    rowCount = UBound(csvData, 1)
    columnCount = UBound(csvData, 2)
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To rowCount
        amount = 0
        If isWithdrawal Then
            If columnCount >= 2 And IsNumeric(csvData(i, 2)) Then
                amount = CLng(csvData(i, 2))
            End If
        Else
            If columnCount >= 3 And IsNumeric(csvData(i, 3)) Then
                amount = CLng(csvData(i, 3))
            End If
        End If

        If amount > 0 Then
            If columnCount >= 4 Then
                descValue = CStr(csvData(i, 4))
            Else
                descValue = ""
            End If
            key = NormalizeSummaryKey(descValue)
            If dict.exists(key) Then
                dict(key) = dict(key) + 1
            Else
                dict.Add key, 1
            End If
        End If
    Next i

    If dict.count = 0 Then GoTo ExitFunc

    keys = dict.keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If dict(keys(j)) > dict(keys(i)) _
                Or (dict(keys(j)) = dict(keys(i)) And StrComp(keys(j), keys(i), vbTextCompare) < 0) Then
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i

    ReDim parts(LBound(keys) To UBound(keys))
    For i = LBound(keys) To UBound(keys)
        parts(i) = keys(i)
    Next i
    BuildUsageSummary = Join(parts, " / ")

    ' 文字数制限（Excelの行高さ制限対策: 最大200文字）
    Const MAX_SUMMARY_LENGTH As Long = 200
    If Len(BuildUsageSummary) > MAX_SUMMARY_LENGTH Then
        BuildUsageSummary = Left$(BuildUsageSummary, MAX_SUMMARY_LENGTH - 3) & "..."
    End If

ExitFunc:
    Exit Function
End Function

' 用途サマリのみを書き込む（取引データがない場合用）
Sub WriteUsageSummaryOnly(csvData As Variant, buttonCol As Long, buttonRow As Long)
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim withdrawalSummary As String
    Dim depositSummary As String

    If IsEmpty(csvData) Then Exit Sub
    If buttonRow <= 1 Then Exit Sub

    targetRow = buttonRow - 1
    Set ws = Worksheets("預金推移")

    withdrawalSummary = BuildUsageSummary(csvData, True)
    depositSummary = BuildUsageSummary(csvData, False)

    ApplySummaryToCell ws.Cells(targetRow, buttonCol), withdrawalSummary
    ApplySummaryToCell ws.Cells(targetRow, buttonCol + 1), depositSummary
End Sub

Private Function ReadUtf8File(filePath As String) As String
    On Error GoTo Failed
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2 'adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        ReadUtf8File = .ReadText
        .Close
    End With
    Exit Function
Failed:
    ReadUtf8File = ""
End Function

Private Function RemoveUtf8Bom(textVal As String) As String
    Dim cleaned As String
    cleaned = textVal
    If Len(cleaned) = 0 Then
        RemoveUtf8Bom = cleaned
        Exit Function
    End If
    If AscW(Left$(cleaned, 1)) = &HFEFF Then
        cleaned = Mid$(cleaned, 2)
    ElseIf Len(cleaned) >= 3 Then
        If Mid$(cleaned, 1, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then
            cleaned = Mid$(cleaned, 4)
        End If
    End If
    RemoveUtf8Bom = cleaned
End Function

Private Function NormalizeHeaderName(rawText As String) As String
    NormalizeHeaderName = LCase$(Trim$(RemoveUtf8Bom(rawText)))
End Function

Private Function SplitCsvFields(lineText As String) As Variant
    Dim results As Object
    Dim current As String
    Dim i As Long
    Dim ch As String
    Dim nextChar As String
    Dim inQuotes As Boolean
    Set results = CreateObject("System.Collections.ArrayList")
    current = ""
    For i = 1 To Len(lineText)
        ch = Mid$(lineText, i, 1)
        If ch = """" Then
            nextChar = ""
            If i < Len(lineText) Then
                nextChar = Mid$(lineText, i + 1, 1)
            End If
            If inQuotes And nextChar = """" Then
                current = current & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            results.Add current
            current = ""
        Else
            current = current & ch
        End If
    Next i
    results.Add current
    SplitCsvFields = results.ToArray
End Function

Private Function GetFieldIndex(fields As Variant, fieldName As String, defaultIndex As Long) As Long
    Dim resolved As Long
    resolved = FindHeaderIndex(fields, fieldName)
    If resolved <> -1 Then
        GetFieldIndex = resolved
    Else
        GetFieldIndex = defaultIndex
    End If
End Function

Private Function FindHeaderIndex(fields As Variant, fieldName As String) As Long
    Dim i As Long
    Dim normalizedTarget As String
    normalizedTarget = LCase$(Trim$(fieldName))
    For i = LBound(fields) To UBound(fields)
        If NormalizeHeaderName(CStr(fields(i))) = normalizedTarget Then
            FindHeaderIndex = i
            Exit Function
        End If
    Next i
    FindHeaderIndex = -1
End Function

Private Function ResolveFieldIndex(fields As Variant, aliases As Variant, defaultIndex As Long) As Long
    Dim aliasName As Variant
    Dim resolved As Long
    For Each aliasName In aliases
        resolved = FindHeaderIndex(fields, CStr(aliasName))
        If resolved <> -1 Then
            ResolveFieldIndex = resolved
            Exit Function
        End If
    Next aliasName
    ResolveFieldIndex = defaultIndex
End Function

Private Function GetArrayValue(fields As Variant, index As Long) As String
    If index >= LBound(fields) And index <= UBound(fields) Then
        GetArrayValue = Trim$(fields(index))
    Else
        GetArrayValue = ""
    End If
End Function

Private Function ToLongValue(valueText As String) As Long
    Dim cleaned As String
    cleaned = Replace(valueText, ",", "")
    cleaned = Replace(cleaned, """", "")
    cleaned = Trim$(cleaned)
    If Len(cleaned) = 0 Then
        ToLongValue = 0
    ElseIf InStr(cleaned, ".") > 0 Then
        cleaned = Left$(cleaned, InStr(cleaned, ".") - 1)
    End If
    If IsNumeric(cleaned) Then
        ToLongValue = CLng(cleaned)
    Else
        ToLongValue = 0
    End If
End Function

Private Function NormalizeSummaryKey(rawText As String) As String
    Dim cleaned As String
    Dim lastChar As String

    cleaned = CleanDescriptionText(rawText)

    ' 特定パターンの正規化（決算利息、受取利子など日付・番号付きの項目をまとめる）
    cleaned = NormalizeCommonPatterns(cleaned)

    ' 末尾の数字を除去
    Do While Len(cleaned) > 0 And IsNumericCharacter(Right$(cleaned, 1))
        cleaned = Left$(cleaned, Len(cleaned) - 1)
        cleaned = Trim$(cleaned)
    Loop
    ' 先頭の数字を除去
    Do While Len(cleaned) > 0 And IsNumericCharacter(Left$(cleaned, 1))
        cleaned = Mid$(cleaned, 2)
        cleaned = Trim$(cleaned)
    Loop
    ' 末尾の記号を除去
    Do While Len(cleaned) > 0
        lastChar = Right$(cleaned, 1)
        If lastChar = "-" Or lastChar = "ー" Or lastChar = "/" Or lastChar = " " Then
            cleaned = Left$(cleaned, Len(cleaned) - 1)
            cleaned = Trim$(cleaned)
        Else
            Exit Do
        End If
    Loop

    If Len(cleaned) = 0 Then
        cleaned = "(摘要なし)"
    End If
    NormalizeSummaryKey = cleaned
End Function

' 共通パターンを正規化（日付や番号付きの摘要をベース名にまとめる）
Private Function NormalizeCommonPatterns(rawText As String) As String
    Dim result As String
    result = rawText

    ' 決算利息 XX-YYマデ/マテ → 決算利息
    If InStr(1, result, "決算利息", vbTextCompare) > 0 Then
        result = "決算利息"
        NormalizeCommonPatterns = result
        Exit Function
    End If

    ' 受取利子 (利子 XX, 税金 YY) → 受取利子
    If InStr(1, result, "受取利子", vbTextCompare) > 0 Then
        result = "受取利子"
        NormalizeCommonPatterns = result
        Exit Function
    End If

    ' 利息 XX → 利息
    If Left$(result, 2) = "利息" Then
        result = "利息"
        NormalizeCommonPatterns = result
        Exit Function
    End If

    NormalizeCommonPatterns = result
End Function

Private Function IsNumericCharacter(ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsNumericCharacter = False
    Else
        IsNumericCharacter = (ch Like "[0-9]") Or (AscW(ch) >= &HFF10 And AscW(ch) <= &HFF19)
    End If
End Function

Private Sub ApplySummaryToCell(targetCell As Range, summaryText As String)
    targetCell.value = summaryText
    targetCell.Font.Color = RGB(0, 0, 0)
    targetCell.Font.Bold = False
    targetCell.WrapText = True
    If Len(summaryText) = 0 Then Exit Sub
    HighlightAttentionKeywords targetCell, summaryText
End Sub

Private Sub HighlightAttentionKeywords(targetCell As Range, textValue As String)
    Dim keywords As Variant
    Dim keyword As Variant
    Dim startPos As Long

    keywords = Array("保険", "ホケン", "ﾎｹﾝ")
    For Each keyword In keywords
        startPos = InStr(1, textValue, keyword, vbTextCompare)
        Do While startPos > 0
            targetCell.Characters(startPos, Len(keyword)).Font.Color = vbRed
            startPos = InStr(startPos + Len(keyword), textValue, keyword, vbTextCompare)
        Loop
    Next keyword
End Sub

Public Function CleanDescriptionText(rawText As String) As String
    Dim textVal As String
    If IsNull(rawText) Then
        CleanDescriptionText = ""
        Exit Function
    End If
    textVal = CStr(rawText)
    textVal = Replace(textVal, "*", "")
    textVal = Replace(textVal, "＊", "")
    textVal = Replace(textVal, "　", " ")
    textVal = Trim$(textVal)
    Do While InStr(textVal, "  ") > 0
        textVal = Replace(textVal, "  ", " ")
    Loop
    CleanDescriptionText = textVal
End Function
'年度を抽出する関数（資金移動.basから流用）
Function ExtractYear(inputText As String) As String
    Dim regex As Object
    Dim match As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "\((\d{4})\)"
    regex.Global = False

    If regex.Test(inputText) Then
        Set match = regex.Execute(inputText)
        ExtractYear = match(0).SubMatches(0)
    Else
        ExtractYear = ""
    End If
End Function

Public Function ParseTransactionCsvContent(csvContent As String, minAmount As Long) As Variant
    Dim normalized As String
    Dim lines() As String
    Dim headerFields As Variant
    Dim lineFields As Variant
    Dim idxDate As Long
    Dim idxDesc As Long
    Dim idxWithdraw As Long
    Dim idxDeposit As Long
    Dim resultData() As Variant
    Dim finalData() As Variant
    Dim dataCount As Long
    Dim i As Long
    Dim transDate As String
    Dim description As String
    Dim withdrawAmount As Long
    Dim depositAmount As Long
    Dim dateAliases As Variant
    Dim descAliases As Variant
    Dim withdrawAliases As Variant
    Dim depositAliases As Variant
    Dim rawWithdraw As String
    Dim rawDeposit As String

    normalized = Replace(csvContent, vbCr, "")
    normalized = RemoveUtf8Bom(normalized)
    lines = Split(normalized, vbLf)
    If UBound(lines) < 0 Then
        ParseTransactionCsvContent = Empty
        Exit Function
    End If
    headerFields = SplitCsvFields(lines(0))
    If UBound(headerFields) < 3 Then
        ParseTransactionCsvContent = Empty
        Exit Function
    End If

    dateAliases = Array("transaction_date", "date", "取引日", "年月日", "日付")
    descAliases = Array("description", "摘要", "内容", "備考")
    withdrawAliases = Array("withdrawal_amount", "withdrawal", "debit", "出金", "支払", "支払金額")
    depositAliases = Array("deposit_amount", "deposit", "credit", "入金", "預入", "入金金額", "預り")

    idxDate = ResolveFieldIndex(headerFields, dateAliases, 0)
    idxDesc = ResolveFieldIndex(headerFields, descAliases, 1)
    idxWithdraw = ResolveFieldIndex(headerFields, withdrawAliases, 2)
    idxDeposit = ResolveFieldIndex(headerFields, depositAliases, 3)

    dataCount = 0
    ReDim resultData(1 To 10000, 1 To 4)

    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            lineFields = SplitCsvFields(lines(i))
            If UBound(lineFields) >= idxDeposit Then
                transDate = GetArrayValue(lineFields, idxDate)
                description = CleanDescriptionText(GetArrayValue(lineFields, idxDesc))
                rawWithdraw = Trim$(GetArrayValue(lineFields, idxWithdraw))
                rawDeposit = Trim$(GetArrayValue(lineFields, idxDeposit))
                withdrawAmount = ToLongValue(rawWithdraw)
                depositAmount = ToLongValue(rawDeposit)

                ' 出金・入金両方が空の行はスキップ（繰越行など）
                If Len(rawWithdraw) = 0 And Len(rawDeposit) = 0 Then
                    ' Skip this row
                ElseIf Abs(withdrawAmount) >= minAmount Or Abs(depositAmount) >= minAmount Then
                    dataCount = dataCount + 1
                    resultData(dataCount, 1) = ConvertDateFormat(transDate)
                    resultData(dataCount, 2) = withdrawAmount
                    resultData(dataCount, 3) = depositAmount
                    resultData(dataCount, 4) = description
                End If
            End If
        End If
    Next i

    If dataCount = 0 Then
        ParseTransactionCsvContent = Empty
    Else
        ReDim finalData(1 To dataCount, 1 To 4)
        For i = 1 To dataCount
            finalData(i, 1) = resultData(i, 1)
            finalData(i, 2) = resultData(i, 2)
            finalData(i, 3) = resultData(i, 3)
            finalData(i, 4) = resultData(i, 4)
        Next i
        ParseTransactionCsvContent = finalData
    End If
End Function

'================ PDF Import Helpers ================

Sub PDF取込ボタン_Click()
    RunPdfImportWorkflow ""
End Sub

Sub 取引履歴取込ボタン_Click()
    RunPdfImportWorkflow "transaction_history"
End Sub

Sub 通帳取込ボタン_Click()
    RunPdfImportWorkflow "bank_deposit"
End Sub

Private Sub RunPdfImportWorkflow(targetDocType As String)
    Dim ws As Worksheet
    Dim buttonCol As Long
    Dim buttonRow As Long
    Dim pdfPath As String
    Dim payloadText As String
    Dim filteredData As Variant
    Dim usageData As Variant
    Dim callerName As String
    Dim shp As Shape
    Dim retryCount As Integer
    Dim dummy As Long
    Dim settings As PdfImportSettings

    Call InitPdfDebugLog(targetDocType)

    ' ブック開直後の初期化待ち（シェイプコレクション準備）
    DoEvents
    Sleep 100
    DoEvents

    Set ws = ActiveSheet

    ' シェイプコレクションを事前にアクセスして初期化を促す
    On Error Resume Next
    dummy = ws.Shapes.Count
    Err.Clear
    On Error GoTo 0

    ' Application.Callerからボタン情報を取得（リトライ付き）
    For retryCount = 1 To 3
        On Error Resume Next
        callerName = Application.Caller
        If Err.Number = 0 And Len(callerName) > 0 Then
            Set shp = ws.Shapes(callerName)
            If Err.Number = 0 And Not shp Is Nothing Then
                On Error GoTo 0
                Exit For
            End If
        End If
        Err.Clear
        On Error GoTo 0

        ' リトライ前に少し待つ
        DoEvents
        Sleep 200
        DoEvents
    Next retryCount

    If shp Is Nothing Then
        MsgBox "初回セットアップが完了しました。" & vbCrLf & vbCrLf & _
               "再度ボタンを押して実行してください。", vbInformation
        Exit Sub
    End If

    buttonCol = shp.TopLeftCell.Column
    buttonRow = shp.TopLeftCell.Row

    ' PDFファイルの選択
    pdfPath = SelectPdfFile()
    If pdfPath = "" Then Exit Sub

    ' 設定を収集（UserFormまたはフォールバックダイアログ）
    settings = CollectPdfImportSettingsFallback(targetDocType)
    If settings.Cancelled Then Exit Sub

    ' API呼び出し
    payloadText = FetchTransactionsJsonText(pdfPath, settings.DocType, settings.DateFormat, _
                                            settings.StartDate, settings.EndDate)
    If Len(payloadText) = 0 Then
        MsgBox "PDF の読み取りに失敗しました。設定値とネットワークを確認してください。", vbExclamation
        Exit Sub
    End If

    LogPdfRawJsonSample payloadText
    usageData = ParseTransactionPayload(payloadText, 0)
    filteredData = ParseTransactionPayload(payloadText, settings.MinAmount)

    If IsEmpty(usageData) Then
        usageData = filteredData
    End If

    ' 指定金額以上の取引がなくても用途サマリは書き込む
    If IsEmpty(filteredData) Then
        ' 取引データなしでも用途サマリのみ書き込み
        If Not IsEmpty(usageData) Then
            Call WriteUsageSummaryOnly(usageData, buttonCol, buttonRow)
            MsgBox "指定金額以上の取引は見つかりませんでした。" & vbCrLf & _
                   "用途欄のみ転記しました。", vbInformation
        Else
            MsgBox "取引データが見つかりませんでした。", vbInformation
        End If
        Exit Sub
    End If

    Call ImportDataToExcel(filteredData, buttonCol, buttonRow, usageData)
    MsgBox "PDF の取り込みが完了しました。", vbInformation
End Sub

Private Function PromptPdfMinimumAmount() As Long
    Dim inputValue As String
    Dim amount As Long

    inputValue = InputBox("取り込む最小金額を入力してください（円単位）" & vbCrLf & _
                          "例：50万円以上の場合は「500000」と入力", _
                          "金額フィルタ", "500000")

    If inputValue = "" Then
        PromptPdfMinimumAmount = -1
        Exit Function
    End If

    If Not IsNumeric(inputValue) Then
        MsgBox "数値を入力してください。", vbExclamation
        PromptPdfMinimumAmount = -1
        Exit Function
    End If

    amount = CLng(inputValue)
    PromptPdfMinimumAmount = amount
End Function

Public Sub InitPdfDebugLog(Optional ByVal scenarioName As String = "")
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim found As Boolean
    Dim i As Long

    If Not DEBUG_LOG_ENABLED Then Exit Sub

    Set wb = ThisWorkbook
    found = False

    For i = 1 To wb.Worksheets.count
        If wb.Worksheets(i).Name = "PDF取込ログ" Then
            Set ws = wb.Worksheets(i)
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = "PDF取込ログ"
    End If

    ws.Cells.Clear
    ws.Range("A1").value = "Step"
    ws.Range("B1").value = "LineIndex"
    ws.Range("C1").value = "RawLine"
    ws.Range("D1").value = "RawWithdraw"
    ws.Range("E1").value = "RawDeposit"
    ws.Range("F1").value = "ParsedWithdraw"
    ws.Range("G1").value = "ParsedDeposit"
    ws.Range("H1").value = "MinAmount"
    ws.Range("I1").value = "PassedFilter"
    ws.Range("J1").value = "Note"

    gDebugRow = 1
    Set gDebugWs = ws
    gDebugRow = gDebugRow + 1
    gDebugWs.Cells(gDebugRow, "A").value = "Start"
    gDebugWs.Cells(gDebugRow, "C").value = "Scenario"
    gDebugWs.Cells(gDebugRow, "D").value = scenarioName
    gDebugWs.Cells(gDebugRow, "E").value = Now
    gDebugRow = gDebugRow + 1
End Sub

Private Sub LogPdfParseRow( _
    ByVal stepName As String, _
    ByVal lineIndex As Long, _
    ByVal rawLine As String, _
    ByVal rawWithdraw As String, _
    ByVal rawDeposit As String, _
    ByVal parsedWithdraw As Long, _
    ByVal parsedDeposit As Long, _
    ByVal minAmount As Long, _
    ByVal passed As Boolean, _
    Optional ByVal note As String = "")

    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    If gDebugWs Is Nothing Then Exit Sub
    On Error GoTo 0

    gDebugRow = gDebugRow + 1

    With gDebugWs
        .Cells(gDebugRow, "A").value = stepName
        .Cells(gDebugRow, "B").value = lineIndex
        .Cells(gDebugRow, "C").value = rawLine
        .Cells(gDebugRow, "D").value = rawWithdraw
        .Cells(gDebugRow, "E").value = rawDeposit
        .Cells(gDebugRow, "F").value = parsedWithdraw
        .Cells(gDebugRow, "G").value = parsedDeposit
        .Cells(gDebugRow, "H").value = minAmount
        .Cells(gDebugRow, "I").value = IIf(passed, "TRUE", "FALSE")
        .Cells(gDebugRow, "J").value = note
    End With
End Sub

Private Sub LogPdfRawJsonSample(jsonText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    Call LogPdfParseRow("RawJson", -1, Left$(jsonText, 200), "", "", 0, 0, 0, True, "preview")
End Sub

Private Function SelectPdfFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .TITLE = "PDF ファイルを選択してください"
        .Filters.Clear
        .Filters.Add "PDF", "*.pdf"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SelectPdfFile = .SelectedItems(1)
        Else
            SelectPdfFile = ""
        End If
    End With
    Set fd = Nothing
End Function

Private Function FetchTransactionsJsonText(pdfPath As String, overrideDocType As String, dateFormat As String, _
    Optional startDate As String = "", Optional endDate As String = "") As String
    On Error GoTo ErrHandler
    Dim baseUrl As String
    Dim apiKey As String
    Dim docType As String
    Dim dateFmt As String
    Dim normalizedBase As String
    Dim jobId As String
    Dim statusUrl As String
    Dim resultUrl As String
    Dim statusJson As String
    Dim jobStatus As String
    Dim jobStatusRaw As String
    Dim stage As String
    Dim detail As String
    Dim startTime As Date
    Dim pollIntervalMs As Long
    Dim maxWaitSeconds As Long
    Dim resultJson As String
    Dim jsonBase64 As String
    Dim displayName As String

    baseUrl = GetConfigValue("BASE_URL")
    apiKey = GetConfigValue("API_KEY")
    docType = ResolveDocumentType(overrideDocType)
    If Len(docType) = 0 Then GoTo Cleanup
    ' 引数で渡された日付形式を使用（空の場合は設定シートから取得）
    If Len(dateFormat) > 0 Then
        dateFmt = dateFormat
    Else
        dateFmt = GetOptionalConfigValue("DATE_FORMAT", "auto")
    End If
    maxWaitSeconds = CLng(GetOptionalConfigValue("JOB_MAX_WAIT_SECONDS", "900"))
    pollIntervalMs = CLng(GetOptionalConfigValue("JOB_POLL_INTERVAL_MS", "4000"))
    If pollIntervalMs < 500 Then pollIntervalMs = 500

    normalizedBase = NormalizeBaseUrl(baseUrl)
    displayName = ExtractFileName(pdfPath)
    jobId = CreateAnalysisJob(normalizedBase & "/jobs", pdfPath, docType, dateFmt, apiKey, startDate, endDate)
    If Len(jobId) = 0 Then GoTo Cleanup

    Application.StatusBar = "ファイル: " & displayName & " ｜ 解析を開始しました"
    statusUrl = normalizedBase & "/jobs/" & jobId
    resultUrl = statusUrl & "/result?format=json"
    startTime = Now

    Do
        statusJson = GetJobStatus(statusUrl, apiKey)
        If Len(statusJson) = 0 Then GoTo Cleanup

        jobStatusRaw = GetJsonStringValue(statusJson, "status")
        jobStatus = LCase$(jobStatusRaw)
        stage = GetJsonStringValue(statusJson, "stage")
        detail = NormalizeDetailText(GetJsonStringValue(statusJson, "detail"))
        UpdateJobStatusBar displayName, jobStatusRaw, stage, detail

        If jobStatus = "completed" Then Exit Do
        If jobStatus = "failed" Then
            MsgBox "解析ジョブが失敗しました: " & IIf(Len(detail) = 0, "(詳細なし)", detail), vbExclamation
            GoTo Cleanup
        End If

        If DateDiff("s", startTime, Now) >= maxWaitSeconds Then
            MsgBox "解析ジョブがタイムアウトしました。Web アプリをご利用ください。", vbExclamation
            GoTo Cleanup
        End If

        Sleep pollIntervalMs
        DoEvents
    Loop

    resultJson = GetJobResult(resultUrl, apiKey)
    If Len(resultJson) = 0 Then GoTo Cleanup

    jsonBase64 = ExtractJobFileBase64(resultJson, "bank_transactions.json")
    If Len(jsonBase64) = 0 Then
        jsonBase64 = ExtractFirstFileBase64(resultJson)
    End If
    FetchTransactionsJsonText = Utf8BytesToString(Base64ToBytes(jsonBase64))

Cleanup:
    Application.StatusBar = False
    Exit Function

ErrHandler:
    MsgBox "API 呼び出しでエラー: " & Err.description, vbCritical
    FetchTransactionsJsonText = ""
    Resume Cleanup
End Function

Private Function CreateAnalysisJob(endpoint As String, pdfPath As String, docType As String, _
    dateFmt As String, apiKey As String, _
    Optional startDate As String = "", Optional endDate As String = "") As String
    Dim boundary As String
    Dim body() As Byte
    Dim http As Object
    Dim jobId As String
    Dim responseText As String

    boundary = "----SOROBOCR" & Format(Now, "yymmddhhmmss")
    body = BuildMultipartBody(pdfPath, boundary, docType, dateFmt, startDate, endDate)

    Set http = CreateHttpClient(120000)
    http.Open "POST", endpoint, False
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.setRequestHeader "Accept", "application/json"
    If Len(apiKey) > 0 Then
        http.setRequestHeader "X-API-Key", apiKey
    End If
    http.send body
    responseText = ReadUtf8Response(http)

    If http.Status <> 200 And http.Status <> 202 Then
        MsgBox "ジョブ作成に失敗しました: " & http.Status & " " & http.statusText & _
               vbCrLf & responseText, vbExclamation
        Exit Function
    End If

    jobId = GetJsonStringValue(responseText, "job_id")
    If Len(jobId) = 0 Then
        MsgBox "ジョブ ID を取得できませんでした。応答: " & responseText, vbExclamation
        Exit Function
    End If

    CreateAnalysisJob = jobId
End Function

Private Function GetJobStatus(statusUrl As String, apiKey As String) As String
    GetJobStatus = SendJsonRequest("GET", statusUrl, apiKey, 60000)
End Function

Private Function GetJobResult(resultUrl As String, apiKey As String) As String
    GetJobResult = SendJsonRequest("GET", resultUrl, apiKey, 180000)
End Function

Private Function SendJsonRequest(method As String, url As String, apiKey As String, _
    receiveTimeoutMs As Long) As String
    On Error GoTo ErrHandler
    Dim http As Object
    Dim responseText As String

    Set http = CreateHttpClient(receiveTimeoutMs)
    http.Open method, url, False
    http.setRequestHeader "Accept", "application/json"
    If Len(apiKey) > 0 Then
        http.setRequestHeader "X-API-Key", apiKey
    End If
    http.send
    responseText = ReadUtf8Response(http)

    If http.Status <> 200 Then
        MsgBox "API応答: " & http.Status & " " & http.statusText & _
               vbCrLf & responseText, vbExclamation
        Exit Function
    End If

    SendJsonRequest = responseText
    Exit Function

ErrHandler:
    MsgBox "API 呼び出しでエラー: " & Err.description, vbCritical
    SendJsonRequest = ""
End Function

Private Function CreateHttpClient(receiveTimeoutMs As Long) As Object
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Option(6) = True
    http.Option(12) = False
    http.Option(4) = 13056
    On Error Resume Next
    http.setTimeouts 30000, 60000, 60000, receiveTimeoutMs
    On Error GoTo 0
    Set CreateHttpClient = http
End Function

Private Sub UpdateJobStatusBar(displayName As String, jobStatus As String, stage As String, detail As String)
    Dim message As String
    Dim stageLabel As String
    Dim statusLabel As String
    Dim hint As String

    statusLabel = TranslateJobStatus(jobStatus)
    stageLabel = TranslateStage(stage)
    hint = detail
    If Len(hint) = 0 Then
        hint = StageDefaultHint(stageLabel)
    End If
    If Len(hint) = 0 Then
        hint = "少々お待ちください..."
    End If

    message = "ファイル: " & displayName
    If Len(stageLabel) > 0 Then
        message = message & " ｜ " & stageLabel
    End If
    If Len(statusLabel) > 0 Then
        message = message & " (" & statusLabel & ")"
    End If
    message = message & " - " & hint

    Application.StatusBar = message
End Sub

Private Function TranslateJobStatus(jobStatus As String) As String
    Select Case LCase$(jobStatus)
        Case "pending": TranslateJobStatus = "待機中"
        Case "running": TranslateJobStatus = "実行中"
        Case "completed": TranslateJobStatus = "完了"
        Case "failed": TranslateJobStatus = "失敗"
        Case Else: TranslateJobStatus = jobStatus
    End Select
End Function

Private Function TranslateStage(stage As String) As String
    Select Case LCase$(stage)
        Case "queued": TranslateStage = "キュー投入"
        Case "analyzing": TranslateStage = "レイアウト解析"
        Case "balance_probe": TranslateStage = "残高スキャン"
        Case "balance_refine": TranslateStage = "AI補正"
        Case "exporting": TranslateStage = "CSV出力"
        Case "completed": TranslateStage = "完了"
        Case "failed": TranslateStage = "失敗"
        Case Else: TranslateStage = stage
    End Select
End Function

Private Function StageDefaultHint(stageLabel As String) As String
    Select Case stageLabel
        Case "キュー投入": StageDefaultHint = "順番待ちです"
        Case "レイアウト解析": StageDefaultHint = "少々お待ちください"
        Case "残高スキャン": StageDefaultHint = "中間と最終残高を確認しています"
        Case "AI補正": StageDefaultHint = "AI 補正を適用しています"
        Case "CSV出力": StageDefaultHint = "結果を作成中です"
        Case Else: StageDefaultHint = ""
    End Select
End Function

Private Function NormalizeDetailText(detail As String) As String
    Dim cleaned As String
    cleaned = Replace(detail, vbCr, " ")
    cleaned = Replace(cleaned, vbLf, " ")
    cleaned = Trim$(cleaned)
    If Len(cleaned) > 80 Then
        NormalizeDetailText = Left$(cleaned, 77) & "..."
    Else
        NormalizeDetailText = cleaned
    End If
End Function

Private Function GetJsonStringValue(json As String, keyName As String) As String
    Dim token As String
    Dim startPos As Long

    token = """" & keyName & """:"
    startPos = InStr(1, json, token, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(token)
    Do While startPos <= Len(json) And Mid$(json, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    If startPos > Len(json) Then Exit Function
    If Mid$(json, startPos, 1) <> """" Then Exit Function
    startPos = startPos + 1

    GetJsonStringValue = ExtractJsonStringAt(json, startPos)
End Function

Private Function ExtractJsonStringAt(json As String, startPos As Long) As String
    Dim i As Long
    Dim ch As String
    For i = startPos To Len(json)
        ch = Mid$(json, i, 1)
        If ch = """" Then
            If i = startPos Then Exit For
            If Mid$(json, i - 1, 1) <> "\" Then Exit For
        End If
        ExtractJsonStringAt = ExtractJsonStringAt & ch
    Next i
End Function

Private Function ExtractFileName(filePath As String) As String
    On Error GoTo Fallback
    ExtractFileName = Mid$(filePath, InStrRev(filePath, Application.PathSeparator) + 1)
    Exit Function
Fallback:
    ExtractFileName = filePath
End Function

Private Function ReadUtf8Response(http As Object) As String
    On Error GoTo Fallback
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    ReadUtf8Response = stream.ReadText
    stream.Close
    Set stream = Nothing
    Exit Function
Fallback:
    ReadUtf8Response = http.responseText
End Function

Private Function NormalizeBaseUrl(baseUrl As String) As String
    Dim tmp As String
    tmp = Trim(baseUrl)
    If Right(tmp, 1) = "/" Then tmp = Left(tmp, Len(tmp) - 1)
    NormalizeBaseUrl = tmp
End Function

Private Function BuildMultipartBody(pdfPath As String, boundary As String, _
    Optional docType As String = "", Optional dateFmt As String = "", _
    Optional startDate As String = "", Optional endDate As String = "") As Byte()
    Dim fileBytes() As Byte
    Dim fileName As String
    Dim stream As Object

    fileBytes = ReadBinaryFile(pdfPath)
    fileName = Mid$(pdfPath, InStrRev(pdfPath, Application.PathSeparator) + 1)

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 'binary
    stream.Open
    If Len(docType) > 0 Then
        stream.Write StringToBytes(BuildTextPart(boundary, "document_type", docType))
    End If
    If Len(dateFmt) > 0 Then
        stream.Write StringToBytes(BuildTextPart(boundary, "date_format", dateFmt))
    End If
    If Len(startDate) > 0 Then
        stream.Write StringToBytes(BuildTextPart(boundary, "start_date", startDate))
    End If
    If Len(endDate) > 0 Then
        stream.Write StringToBytes(BuildTextPart(boundary, "end_date", endDate))
    End If
    stream.Write StringToBytes(BuildFileHeader(boundary, fileName))
    stream.Write fileBytes
    stream.Write StringToBytes(BuildClosingBoundary(boundary))

    stream.Position = 0
    BuildMultipartBody = stream.Read
    stream.Close
    Set stream = Nothing
End Function

Private Function BuildTextPart(boundary As String, fieldName As String, fieldValue As String) As String
    BuildTextPart = "--" & boundary & vbCrLf & _
                    "Content-Disposition: form-data; name=""" & fieldName & """" & vbCrLf & vbCrLf & _
                    fieldValue & vbCrLf
End Function

Private Function BuildFileHeader(boundary As String, fileName As String) As String
    BuildFileHeader = "--" & boundary & vbCrLf & _
                      "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf & _
                      "Content-Type: application/pdf" & vbCrLf & vbCrLf
End Function

Private Function BuildClosingBoundary(boundary As String) As String
    BuildClosingBoundary = vbCrLf & "--" & boundary & "--" & vbCrLf
End Function

Private Function ReadBinaryFile(filePath As String) As Byte()
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile filePath
    stream.Position = 0
    ReadBinaryFile = stream.Read
    stream.Close
    Set stream = Nothing
End Function

Private Function StringToBytes(textValue As String) As Byte()
    Dim stream As Object
    Dim raw() As Byte
    Dim trimmed() As Byte
    Dim i As Long
    Dim startIndex As Long

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' text mode
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textValue

    stream.Position = 0
    stream.Type = 1 ' binary
    raw = stream.Read
    stream.Close
    Set stream = Nothing

    Dim upperBound As Long
    On Error Resume Next
    upperBound = UBound(raw)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        StringToBytes = raw
        Exit Function
    End If
    On Error GoTo 0

    startIndex = 0
    If upperBound >= 2 Then
        If raw(0) = &HEF And raw(1) = &HBB And raw(2) = &HBF Then
            startIndex = 3 ' skip BOM
        End If
    End If

    If startIndex = 0 Then
        StringToBytes = raw
    Else
        ReDim trimmed(upperBound - startIndex)
        For i = 0 To UBound(trimmed)
            trimmed(i) = raw(i + startIndex)
        Next i
        StringToBytes = trimmed
    End If
End Function

Private Function Base64ToBytes(base64Value As String) As Byte()
    Dim xml As Object
    Dim node As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.dataType = "bin.base64"
    node.text = base64Value
    Base64ToBytes = node.nodeTypedValue
End Function

Private Function Utf8BytesToString(data() As Byte) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 1
        .Open
        .Write data
        .Position = 0
        .Type = 2
        .Charset = "UTF-8"
        Utf8BytesToString = .ReadText
        .Close
    End With
End Function

Private Function ExtractJobFileBase64(json As String, fileName As String) As String
    Dim token As String
    Dim startPos As Long

    token = """" & fileName & """:"""
    startPos = InStr(1, json, token, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(token)

    ExtractJobFileBase64 = ExtractJsonStringAt(json, startPos)
End Function

Private Function ExtractFirstFileBase64(json As String) As String
    Dim token As String
    Dim pos As Long
    Dim colonPos As Long

    token = """files"":{"
    pos = InStr(1, json, token, vbTextCompare)
    If pos = 0 Then Exit Function
    pos = pos + Len(token)

    colonPos = InStr(pos, json, Chr$(34) & ":" & Chr$(34), vbTextCompare)
    If colonPos = 0 Then Exit Function
    pos = colonPos + 3

    ExtractFirstFileBase64 = ExtractJsonStringAt(json, pos)
End Function

Private Function ParseCsvText(csvContent As String, minAmount As Long) As Variant
    Dim lines() As String
    Dim lineData() As String
    Dim resultData() As Variant
    Dim finalData() As Variant
    Dim i As Long
    Dim dataCount As Long
    Dim transDate As String
    Dim description As String
    Dim withdrawAmount As Long
    Dim depositAmount As Long
    Dim rawLine As String
    Dim rawWithdraw As String
    Dim rawDeposit As String
    Dim passed As Boolean

    lines = Split(csvContent, vbLf)
    dataCount = 0
    ReDim resultData(1 To 10000, 1 To 4)

    For i = 1 To UBound(lines)
        rawLine = NormalizePdfCsvLine(RemoveUtf8Bom(lines(i)))
        If Len(Trim$(rawLine)) > 0 Then
            lineData = PdfSplitCsvLine(rawLine)
            If UBound(lineData) >= 3 Then
                transDate = lineData(0)
                description = CleanDescriptionText(lineData(1))
                rawWithdraw = Trim$(lineData(2))
                rawDeposit = Trim$(lineData(3))
                withdrawAmount = PdfToLong(rawWithdraw)
                depositAmount = PdfToLong(rawDeposit)

                ' 出金・入金両方が空の行はスキップ（繰越行など）
                If Len(rawWithdraw) = 0 And Len(rawDeposit) = 0 Then
                    passed = False
                Else
                    passed = (Abs(withdrawAmount) >= minAmount Or Abs(depositAmount) >= minAmount)
                End If

                Call LogPdfParseRow( _
                    "ParseCsvText", _
                    i, _
                    rawLine, _
                    rawWithdraw, _
                    rawDeposit, _
                    withdrawAmount, _
                    depositAmount, _
                    minAmount, _
                    passed, _
                    "" _
                )
                If passed Then
                    dataCount = dataCount + 1
                    resultData(dataCount, 1) = ConvertDateFormat(transDate)
                    resultData(dataCount, 2) = withdrawAmount
                    resultData(dataCount, 3) = depositAmount
                    resultData(dataCount, 4) = description
                End If
            Else
                Call LogPdfParseRow( _
                    "ParseCsvText", _
                    i, _
                    rawLine, _
                    "", _
                    "", _
                    0, _
                    0, _
                    minAmount, _
                    False, _
                    "Skipped: columns=" & (UBound(lineData) + 1))
            End If
        End If
    Next i

    Call LogPdfParseRow("Summary", -1, "", "", "", dataCount, 0, minAmount, (dataCount > 0), "dataCount")

    If dataCount = 0 Then
        ParseCsvText = Empty
    Else
        ReDim finalData(1 To dataCount, 1 To 4)
        For i = 1 To dataCount
            finalData(i, 1) = resultData(i, 1)
            finalData(i, 2) = resultData(i, 2)
            finalData(i, 3) = resultData(i, 3)
            finalData(i, 4) = resultData(i, 4)
        Next i
        ParseCsvText = finalData
    End If
End Function

Private Function PdfSplitCsvLine(lineText As String) As String()
    Dim results As Object
    Dim token As String
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim arr() As String

    Set results = CreateObject("System.Collections.ArrayList")
    token = ""
    For i = 1 To Len(lineText)
        ch = Mid$(lineText, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf ch = "," And Not inQuotes Then
            results.Add token
            token = ""
        Else
            token = token & ch
        End If
    Next i
    results.Add token
    If results.count = 0 Then
        ReDim arr(0 To 0)
        arr(0) = ""
    Else
        ReDim arr(0 To results.count - 1)
        For i = 0 To results.count - 1
            arr(i) = CStr(results(i))
        Next i
    End If
    PdfSplitCsvLine = arr
End Function

Private Function PdfToLong(valueText As String) As Long
    Dim cleaned As String
    cleaned = RemoveUtf8Bom(valueText)
    cleaned = Replace(cleaned, "－", "-")
    cleaned = Replace(cleaned, "?", "-")
    cleaned = Replace(cleaned, ",", "")
    cleaned = Replace(cleaned, """", "")
    cleaned = Trim$(cleaned)
    cleaned = Replace(cleaned, Chr$(160), "")
    If Len(cleaned) = 0 Then
        PdfToLong = 0
    Else
        If InStr(cleaned, ".") > 0 Then
            cleaned = Left$(cleaned, InStr(cleaned, ".") - 1)
        End If
        If IsNumeric(cleaned) Then
            PdfToLong = CLng(cleaned)
        Else
            PdfToLong = 0
        End If
    End If
End Function

Private Function NormalizePdfCsvLine(lineText As String) As String
    Dim cleaned As String
    cleaned = Trim$(lineText)
    If Len(cleaned) >= 1 Then
        If Left$(cleaned, 1) = """" Then
            cleaned = Mid$(cleaned, 2)
        End If
    End If
    If Len(cleaned) >= 1 Then
        If Right$(cleaned, 1) = """" Then
            cleaned = Left$(cleaned, Len(cleaned) - 1)
        End If
    End If
    cleaned = Replace(cleaned, ChrW(&HFF0C), ",")
    cleaned = Replace(cleaned, ChrW(&HFF1A), ":")
    cleaned = Replace(cleaned, ChrW(&H3001), ",")
    NormalizePdfCsvLine = cleaned
End Function

Private Function ExtractTransactionsArray(jsonText As String) As String
    Dim pos As Long
    Dim keyName As String
    Dim arrayStart As Long
    Dim ch As String
    Dim skipped As Variant

    pos = 1
    JsonSkipWhitespace jsonText, pos
    If JsonPeek(jsonText, pos) <> "{" Then Exit Function
    pos = pos + 1

    Do
        JsonSkipWhitespace jsonText, pos
        ch = JsonPeek(jsonText, pos)
        If ch = "" Then Exit Do
        If ch = "}" Then Exit Do
        keyName = JsonParseString(jsonText, pos)
        JsonSkipWhitespace jsonText, pos
        If JsonPeek(jsonText, pos) <> ":" Then Exit Function
        pos = pos + 1
        JsonSkipWhitespace jsonText, pos
        If LCase$(keyName) = "transactions" Then
            If JsonPeek(jsonText, pos) = "[" Then
                arrayStart = pos
                JsonSkipArray jsonText, pos
                ExtractTransactionsArray = Mid$(jsonText, arrayStart, pos - arrayStart)
                Exit Function
            End If
        End If
        skipped = JsonParseValue(jsonText, pos)
        JsonSkipWhitespace jsonText, pos
        ch = JsonPeek(jsonText, pos)
        If ch = "," Then
            pos = pos + 1
        ElseIf ch = "}" Then
            Exit Do
        End If
    Loop
End Function

Private Function ParseTransactionJsonContent(jsonText As String, minAmount As Long) As Variant
    Dim pos As Long
    Dim ch As String
    Dim capacity As Long
    Dim count As Long
    Dim temp() As Variant
    Dim objectIndex As Long
    Dim txnDate As String
    Dim description As String
    Dim withdrawAmount As Long
    Dim depositAmount As Long
    Dim rawLine As String
    Dim passed As Boolean
    Dim value As Variant
    Dim key As String
    Dim balanceValue As Variant
    Dim hasWithdraw As Boolean
    Dim hasDeposit As Boolean
    Dim finalData() As Variant
    Dim i As Long

    pos = 1
    capacity = 64
    count = 0
    ReDim temp(1 To capacity, 1 To 4)

    JsonSkipWhitespace jsonText, pos
    If JsonPeek(jsonText, pos) <> "[" Then
        ParseTransactionJsonContent = Empty
        Exit Function
    End If
    pos = pos + 1

    Do
ContinueArrayLoop:
        JsonSkipWhitespace jsonText, pos
        ch = JsonPeek(jsonText, pos)
        If ch = "" Then Exit Do
        If ch = "]" Then
            pos = pos + 1
            Exit Do
        End If
        If ch = "," Then
            pos = pos + 1
            GoTo ContinueArrayLoop
        End If
        If ch <> "{" Then Exit Do
        objectIndex = objectIndex + 1
        pos = pos + 1

        txnDate = ""
        description = ""
        withdrawAmount = 0
        depositAmount = 0
        balanceValue = Null
        hasWithdraw = False
        hasDeposit = False

        Do
ContinueObjectLoop:
            JsonSkipWhitespace jsonText, pos
            ch = JsonPeek(jsonText, pos)
            If ch = "}" Then
                pos = pos + 1
                Exit Do
            End If
            key = JsonParseString(jsonText, pos)
            JsonSkipWhitespace jsonText, pos
            If JsonPeek(jsonText, pos) <> ":" Then Exit Do
            pos = pos + 1
            JsonSkipWhitespace jsonText, pos
            value = JsonParseValue(jsonText, pos)
            Select Case LCase$(key)
                Case "transaction_date", "date"
                    If Not IsNull(value) Then txnDate = CStr(value)
                Case "description"
                    If Not IsNull(value) Then description = CStr(value)
                Case "memo"
                    If Len(description) = 0 And Not IsNull(value) Then description = CStr(value)
                Case "withdrawal_amount", "withdrawal"
                    If Not IsNull(value) Then
                        withdrawAmount = CLng(value)
                        hasWithdraw = True
                    End If
                Case "deposit_amount", "deposit"
                    If Not IsNull(value) Then
                        depositAmount = CLng(value)
                        hasDeposit = True
                    End If
                Case "balance"
                    If Not IsNull(value) Then balanceValue = CLng(value)
                Case "correction_note", "correctionnote"
                    ' ignore for now
            End Select
            JsonSkipWhitespace jsonText, pos
            ch = JsonPeek(jsonText, pos)
            If ch = "," Then
                pos = pos + 1
                GoTo ContinueObjectLoop
            ElseIf ch = "}" Then
                pos = pos + 1
                Exit Do
            End If
        Loop

        ' 出金・入金両方がnullの行はスキップ（繰越行など）
        If Not hasWithdraw And Not hasDeposit Then
            passed = False
        Else
            passed = (Abs(withdrawAmount) >= minAmount Or Abs(depositAmount) >= minAmount)
        End If
        rawLine = txnDate & "," & description
        LogPdfParseRow "ParseJson", objectIndex, rawLine, CStr(withdrawAmount), CStr(depositAmount), _
            withdrawAmount, depositAmount, minAmount, passed, ""

        If passed Then
            count = count + 1
            If count > capacity Then
                capacity = capacity * 2
                temp = ResizeTransactionArray(temp, capacity)
            End If
            temp(count, 1) = ConvertDateFormat(txnDate)
            temp(count, 2) = withdrawAmount
            temp(count, 3) = depositAmount
            temp(count, 4) = description
        End If
    Loop

    LogPdfParseRow "Summary(JSON)", -1, "", "", "", count, 0, minAmount, (count > 0), "dataCount"

    If count = 0 Then
        ParseTransactionJsonContent = Empty
    Else
        ReDim finalData(1 To count, 1 To 4)
        For i = 1 To count
            finalData(i, 1) = temp(i, 1)
            finalData(i, 2) = temp(i, 2)
            finalData(i, 3) = temp(i, 3)
            finalData(i, 4) = temp(i, 4)
        Next i
        ParseTransactionJsonContent = finalData
    End If
End Function

Private Function ParseTransactionPayload(rawText As String, minAmount As Long) As Variant
    Dim trimmed As String
    Dim firstChar As String
    Dim transactionsJson As String

    trimmed = LTrim$(rawText)
    If Len(trimmed) = 0 Then
        ParseTransactionPayload = Empty
        Exit Function
    End If
    firstChar = Left$(trimmed, 1)
    If firstChar = "[" Then
        ParseTransactionPayload = ParseTransactionJsonContent(trimmed, minAmount)
    ElseIf firstChar = "{" Then
        transactionsJson = ExtractTransactionsArray(trimmed)
        If Len(transactionsJson) = 0 Then
            ParseTransactionPayload = Empty
        Else
            ParseTransactionPayload = ParseTransactionJsonContent(transactionsJson, minAmount)
        End If
    Else
        ParseTransactionPayload = ParseCsvText(rawText, minAmount)
    End If
End Function

Private Function ResizeTransactionArray(oldArray As Variant, newCapacity As Long) As Variant
    Dim newArray() As Variant
    Dim i As Long
    Dim oldCount As Long

    On Error Resume Next
    oldCount = UBound(oldArray, 1)
    On Error GoTo 0
    If oldCount = 0 Then oldCount = 0

    ReDim newArray(1 To newCapacity, 1 To 4)
    For i = 1 To oldCount
        newArray(i, 1) = oldArray(i, 1)
        newArray(i, 2) = oldArray(i, 2)
        newArray(i, 3) = oldArray(i, 3)
        newArray(i, 4) = oldArray(i, 4)
    Next i

    ResizeTransactionArray = newArray
End Function

Private Sub JsonSkipWhitespace(ByVal text As String, ByRef pos As Long)
    Dim length As Long
    length = Len(text)
    Do While pos <= length
        Select Case Mid$(text, pos, 1)
            Case " ", vbTab, vbCr, vbLf
                pos = pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Function JsonPeek(ByVal text As String, ByVal pos As Long) As String
    If pos > Len(text) Or pos <= 0 Then
        JsonPeek = ""
    Else
        JsonPeek = Mid$(text, pos, 1)
    End If
End Function

Private Function JsonParseString(ByVal text As String, ByRef pos As Long) As String
    Dim result As String
    Dim ch As String
    Dim code As String
    result = ""
    If JsonPeek(text, pos) <> """" Then
        JsonParseString = ""
        Exit Function
    End If
    pos = pos + 1
    Do While pos <= Len(text)
        ch = Mid$(text, pos, 1)
        If ch = "\" Then
            pos = pos + 1
            ch = Mid$(text, pos, 1)
            Select Case ch
                Case """", "\", "/"
                    result = result & ch
                Case "b": result = result & vbBack
                Case "f": result = result & vbFormFeed
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u"
                    code = Mid$(text, pos + 1, 4)
                    result = result & ChrW(CLng("&H" & code))
                    pos = pos + 4
                Case Else
                    result = result & ch
            End Select
            pos = pos + 1
        ElseIf ch = """" Then
            pos = pos + 1
            Exit Do
        Else
            result = result & ch
            pos = pos + 1
        End If
    Loop
    JsonParseString = result
End Function

Private Function JsonParseValue(ByVal text As String, ByRef pos As Long) As Variant
    Dim ch As String
    ch = JsonPeek(text, pos)
    Select Case ch
        Case """"
            JsonParseValue = JsonParseString(text, pos)
        Case "-", "0" To "9"
            JsonParseValue = JsonParseNumber(text, pos)
        Case "n"
            If Mid$(text, pos, 4) = "null" Then
                pos = pos + 4
                JsonParseValue = Null
            End If
        Case "t"
            If Mid$(text, pos, 4) = "true" Then
                pos = pos + 4
                JsonParseValue = True
            End If
        Case "f"
            If Mid$(text, pos, 5) = "false" Then
                pos = pos + 5
                JsonParseValue = False
            End If
        Case "{"
            JsonSkipObject text, pos
            JsonParseValue = Null
        Case "["
            JsonSkipArray text, pos
            JsonParseValue = Null
        Case Else
            JsonParseValue = Null
    End Select
End Function

Private Function JsonParseNumber(ByVal text As String, ByRef pos As Long) As Double
    Dim startPos As Long
    Dim ch As String
    startPos = pos
    Do While pos <= Len(text)
        ch = Mid$(text, pos, 1)
        If InStr("0123456789+-eE.", ch) = 0 Then Exit Do
        pos = pos + 1
    Loop
    JsonParseNumber = val(Mid$(text, startPos, pos - startPos))
End Function

Private Sub JsonSkipObject(ByVal text As String, ByRef pos As Long)
    Dim depth As Long
    Dim ch As String
    depth = 0
    Do While pos <= Len(text)
        ch = Mid$(text, pos, 1)
        If ch = "{" Then
            depth = depth + 1
            pos = pos + 1
        ElseIf ch = "}" Then
            depth = depth - 1
            pos = pos + 1
            If depth = 0 Then Exit Do
        ElseIf ch = """" Then
            JsonParseString text, pos
        Else
            pos = pos + 1
        End If
    Loop
End Sub

Private Sub JsonSkipArray(ByVal text As String, ByRef pos As Long)
    Dim depth As Long
    Dim ch As String
    depth = 0
    Do While pos <= Len(text)
        ch = Mid$(text, pos, 1)
        If ch = "[" Then
            depth = depth + 1
            pos = pos + 1
        ElseIf ch = "]" Then
            depth = depth - 1
            pos = pos + 1
            If depth = 0 Then Exit Do
        ElseIf ch = """" Then
            JsonParseString text, pos
        Else
            pos = pos + 1
        End If
    Loop
End Sub
Private Function GetConfigValue(keyName As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("設定")
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1, , "設定シートが見つかりません。"
    End If

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 1 To lastRow
        If Trim$(ws.Cells(i, 1).value) = keyName Then
            GetConfigValue = Trim$(ws.Cells(i, 2).value)
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 2, , "設定シートに " & keyName & " が定義されていません。"
End Function

Private Function GetOptionalConfigValue(keyName As String, defaultValue As String) As String
    On Error GoTo MissingKey
    Dim value As String
    value = GetConfigValue(keyName)
    If Len(value) = 0 Then
        GetOptionalConfigValue = defaultValue
    Else
        GetOptionalConfigValue = value
    End If
    Exit Function
MissingKey:
    Err.Clear
    GetOptionalConfigValue = defaultValue
End Function

Private Function ResolveDocumentType(overrideDocType As String) As String
    Dim defaultType As String
    If Len(overrideDocType) > 0 Then
        ResolveDocumentType = overrideDocType
        Exit Function
    End If

    defaultType = GetOptionalConfigValue("DOC_TYPE", "transaction_history")

    ResolveDocumentType = PromptDocTypeSelection(defaultType)
End Function

Private Function PromptDocTypeSelection(defaultType As String) As String
    Dim prompt As String
    Dim choice As VbMsgBoxResult
    Dim defaultHint As String

    defaultHint = IIf(LCase$(defaultType) = "bank_deposit", "（既定: 通帳）", "（既定: 取引履歴）")
    prompt = "読み取るPDFは通帳ですか？" & vbCrLf & _
             "・通帳の場合は「はい」を選択" & vbCrLf & _
             "・取引履歴の場合は「いいえ」を選択" & vbCrLf & _
             "・処理を中止する場合は「キャンセル」を選択" & vbCrLf & _
             defaultHint

    Do
        choice = MsgBox(prompt, vbQuestion + vbYesNoCancel, "書類タイプの選択")
        Select Case choice
            Case vbYes
                PromptDocTypeSelection = "bank_deposit"
                Exit Function
            Case vbNo
                PromptDocTypeSelection = "transaction_history"
                Exit Function
            Case vbCancel
                PromptDocTypeSelection = ""
                Exit Function
        End Select
    Loop
End Function

Private Function PromptDateFormatSelection() As String
    Dim prompt As String
    Dim choice As Integer

    prompt = "通帳の日付表記を選択してください：" & vbCrLf & vbCrLf & _
             "　1 = 自動判定（推奨）" & vbCrLf & _
             "　2 = 和暦（三菱UFJ、ゆうちょなど）" & vbCrLf & _
             "　　　例: 01-12-06 → 令和1年12月6日" & vbCrLf & _
             "　3 = 西暦（みずほ銀行など）" & vbCrLf & _
             "　　　例: 20-02-14 → 2020年2月14日"

    Dim inputValue As String
    inputValue = InputBox(prompt, "日付形式の選択", "1")

    If Len(inputValue) = 0 Then
        PromptDateFormatSelection = ""
        Exit Function
    End If

    Select Case Trim$(inputValue)
        Case "1"
            PromptDateFormatSelection = "auto"
        Case "2"
            PromptDateFormatSelection = "auto"  ' 和暦も auto で処理
        Case "3"
            PromptDateFormatSelection = "western"
        Case Else
            PromptDateFormatSelection = "auto"
    End Select
End Function

'===============================================================================
' PDF取込設定を収集（UserFormまたはフォールバックダイアログ）
'===============================================================================
Public Function CollectPdfImportSettings(defaultDocType As String) As PdfImportSettings
    Dim settings As PdfImportSettings
    settings.Cancelled = True

    ' UserFormが利用可能か試行
    On Error Resume Next
    Dim frm As Object
    Set frm = Nothing

    ' frmPdfImportSettingsが存在する場合は使用
    Set frm = UserForms.Add("frmPdfImportSettings")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' UserFormが使えない場合はフォールバック
        settings = CollectPdfImportSettingsFallback(defaultDocType)
        Exit Function
    End If
    On Error GoTo 0

    ' UserFormを表示
    frm.Show

    If frm.Cancelled Then
        Unload frm
        Exit Function
    End If

    settings.Cancelled = False
    settings.DocType = frm.DocType
    settings.DateFormat = frm.DateFormat
    settings.MinAmount = frm.MinAmount
    settings.StartDate = frm.NormalizeDateForApi()
    settings.EndDate = frm.NormalizeEndDateForApi()

    Unload frm
    CollectPdfImportSettings = settings
End Function

'===============================================================================
' フォールバック：個別ダイアログで設定を収集
'===============================================================================
Private Function CollectPdfImportSettingsFallback(defaultDocType As String) As PdfImportSettings
    Dim settings As PdfImportSettings
    settings.Cancelled = True

    ' 書類タイプの選択（通帳ボタンから呼ばれた場合はスキップ）
    If Len(defaultDocType) > 0 Then
        settings.DocType = defaultDocType
    Else
        settings.DocType = ResolveDocumentType("")
        If Len(settings.DocType) = 0 Then Exit Function
    End If

    ' 日付形式の選択
    settings.DateFormat = PromptDateFormatSelection()
    If Len(settings.DateFormat) = 0 Then Exit Function

    ' 最小金額の入力
    Dim minAmountResult As Long
    minAmountResult = PromptPdfMinimumAmount()
    If minAmountResult = -1 Then Exit Function
    settings.MinAmount = minAmountResult

    ' 開始日の入力（任意）
    settings.StartDate = PromptDateInput("開始日", "取込開始日を入力してください（任意）" & vbCrLf & _
        "空欄にすると全期間を対象にします。" & vbCrLf & vbCrLf & _
        "形式: YYYY-MM-DD （例: 2021-04-01）")

    ' 終了日の入力（任意）
    settings.EndDate = PromptDateInput("終了日", "取込終了日を入力してください（任意）" & vbCrLf & _
        "空欄にすると最新まで取り込みます。" & vbCrLf & vbCrLf & _
        "形式: YYYY-MM-DD （例: 2024-12-31）")

    settings.Cancelled = False
    CollectPdfImportSettingsFallback = settings
End Function

'===============================================================================
' 日付入力ダイアログ（任意入力用）
'===============================================================================
Private Function PromptDateInput(title As String, prompt As String) As String
    Dim inputValue As String

    inputValue = InputBox(prompt, title, "")

    If Len(Trim$(inputValue)) = 0 Then
        PromptDateInput = ""
        Exit Function
    End If

    ' 簡易バリデーション: 数字とハイフン/スラッシュのみ
    inputValue = Replace(inputValue, "/", "-")
    If Len(inputValue) > 0 Then
        ' YYYY-MM-DD形式かどうか簡易チェック
        Dim parts() As String
        parts = Split(inputValue, "-")
        If UBound(parts) = 2 Then
            If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
                PromptDateInput = inputValue
                Exit Function
            End If
        End If
        MsgBox "日付形式が正しくありません。空欄として扱います。" & vbCrLf & _
               "正しい形式: YYYY-MM-DD", vbExclamation
    End If

    PromptDateInput = ""
End Function


