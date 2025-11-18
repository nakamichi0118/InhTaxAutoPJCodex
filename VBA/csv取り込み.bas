Option Explicit

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

    Set ws = ActiveSheet

    '1. ボタン位置の取得
    buttonCol = ws.Shapes(Application.Caller).TopLeftCell.Column
    buttonRow = ws.Shapes(Application.Caller).TopLeftCell.Row

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

    Set ws = Worksheets("預金推移")

    '範囲を取得
    gyouhajime = ws.Range("A1:A1000").Find("資金移動始").Row
    gyousaigo = ws.Range("A1:A1000").Find("資金移動終").Row
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
            ws.Cells(insertRow, 3).Value = dateToFind

            If csvData(i, 2) > 0 Then '出金
                ws.Cells(insertRow, buttonCol).Value = Format(csvData(i, 2), "#,##0")
                ws.Cells(insertRow, buttonCol).HorizontalAlignment = xlRight
            End If

            If csvData(i, 3) > 0 Then '入金
                ws.Cells(insertRow, buttonCol + 1).Value = Format(csvData(i, 3), "#,##0")
                ws.Cells(insertRow, buttonCol + 1).HorizontalAlignment = xlRight
            End If

        Else
            '同じ日付がある場合
            existingRow = FoundCell.Row
            emptyFound = False

            '同じ日付の行を確認
            Do
                '空白セルをチェック
                If ws.Cells(existingRow, buttonCol).Value = "" And ws.Cells(existingRow, buttonCol + 1).Value = "" Then
                    '空白セルがある場合はそこに入力
                    If csvData(i, 2) > 0 Then '出金
                        ws.Cells(existingRow, buttonCol).Value = Format(csvData(i, 2), "#,##0")
                        ws.Cells(existingRow, buttonCol).HorizontalAlignment = xlRight
                    End If

                    If csvData(i, 3) > 0 Then '入金
                        ws.Cells(existingRow, buttonCol + 1).Value = Format(csvData(i, 3), "#,##0")
                        ws.Cells(existingRow, buttonCol + 1).HorizontalAlignment = xlRight
                    End If

                    emptyFound = True
                    Exit Do
                End If

                '次の行が同じ日付かチェック
                If existingRow < gyousaigo - 3 Then
                    If ws.Cells(existingRow + 1, 3).Value = dateToFind Then
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
                ws.Cells(insertRow, 3).Value = dateToFind

                If csvData(i, 2) > 0 Then '出金
                    ws.Cells(insertRow, buttonCol).Value = Format(csvData(i, 2), "#,##0")
                    ws.Cells(insertRow, buttonCol).HorizontalAlignment = xlRight
                End If

                If csvData(i, 3) > 0 Then '入金
                    ws.Cells(insertRow, buttonCol + 1).Value = Format(csvData(i, 3), "#,##0")
                    ws.Cells(insertRow, buttonCol + 1).HorizontalAlignment = xlRight
                End If
            End If
        End If
    Next i

    '範囲を再取得してソート
    gyouhajime = ws.Range("A1:A1000").Find("資金移動始").Row
    gyousaigo = ws.Range("A1:A1000").Find("資金移動終").Row
    retuhajime = ws.Range("C1:BZ1").Find("資金移動始").Column
    retusaigo = ws.Range("C1:DZ1").Find("資金移動終").Column

    With ws
        .Range(.Cells(gyouhajime + 1, retuhajime), .Cells(gyousaigo - 3, retusaigo)).Sort Key1:=.Columns(3)
    End With

    '罫線の更新
    Call UpdateBorders

    '用途サマリをボタン行の1つ上へ表示
    Dim summarySource As Variant
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

    gyouhajime = ws.Range("A1:A1000").Find("資金移動始").Row
    gyousaigo = ws.Range("A1:A1000").Find("資金移動終").Row
    retuhajime = ws.Range("C1:BZ1").Find("資金移動始").Column
    retusaigo = ws.Range("C1:DZ1").Find("資金移動終").Column

    '横罫線を設定
    ws.Range(ws.Cells(gyouhajime, retuhajime), ws.Cells(gyousaigo, retusaigo)).Borders(xlInsideHorizontal).LineStyle = xlContinuous

    '年度境界に色付き罫線
    For i = gyouhajime + 1 To gyousaigo - 4
        currentYear = ExtractYear(ws.Cells(i, 3).Value)
        nextYear = ExtractYear(ws.Cells(i + 1, 3).Value)

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
            If dict.Exists(key) Then
                dict(key) = dict(key) + 1
            Else
                dict.Add key, 1
            End If
        End If
    Next i

    If dict.Count = 0 Then GoTo ExitFunc

    keys = dict.Keys
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
        parts(i) = keys(i) & "(" & dict(keys(i)) & ")"
    Next i
    BuildUsageSummary = Join(parts, " / ")

ExitFunc:
    Exit Function
End Function

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
    If Len(textVal) > 0 Then
        If AscW(Left$(textVal, 1)) = &HFEFF Then
            RemoveUtf8Bom = Mid$(textVal, 2)
        Else
            RemoveUtf8Bom = textVal
        End If
    Else
        RemoveUtf8Bom = textVal
    End If
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
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        If StrComp(Trim$(fields(i)), fieldName, vbTextCompare) = 0 Then
            GetFieldIndex = i
            Exit Function
        End If
    Next i
    GetFieldIndex = defaultIndex
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
    cleaned = CleanDescriptionText(rawText)
    Do While Len(cleaned) > 0 And IsNumericCharacter(Right$(cleaned, 1))
        cleaned = Left$(cleaned, Len(cleaned) - 1)
        cleaned = Trim$(cleaned)
    Loop
    Do While Len(cleaned) > 0 And IsNumericCharacter(Left$(cleaned, 1))
        cleaned = Mid$(cleaned, 2)
        cleaned = Trim$(cleaned)
    Loop
    If Len(cleaned) = 0 Then
        cleaned = "(摘要なし)"
    End If
    NormalizeSummaryKey = cleaned
End Function

Private Function IsNumericCharacter(ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsNumericCharacter = False
    Else
        IsNumericCharacter = (ch Like "[0-9]") Or (AscW(ch) >= &HFF10 And AscW(ch) <= &HFF19)
    End If
End Function

Private Sub ApplySummaryToCell(targetCell As Range, summaryText As String)
    targetCell.Value = summaryText
    targetCell.Font.Color = RGB(0, 0, 0)
    targetCell.Font.Bold = False
    targetCell.WrapText = False
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
    idxDate = GetFieldIndex(headerFields, "transaction_date", 0)
    idxDesc = GetFieldIndex(headerFields, "description", 1)
    idxWithdraw = GetFieldIndex(headerFields, "withdrawal_amount", 2)
    idxDeposit = GetFieldIndex(headerFields, "deposit_amount", 3)

    dataCount = 0
    ReDim resultData(1 To 10000, 1 To 4)

    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            lineFields = SplitCsvFields(lines(i))
            If UBound(lineFields) >= idxDeposit Then
                transDate = GetArrayValue(lineFields, idxDate)
                description = CleanDescriptionText(GetArrayValue(lineFields, idxDesc))
                withdrawAmount = ToLongValue(GetArrayValue(lineFields, idxWithdraw))
                depositAmount = ToLongValue(GetArrayValue(lineFields, idxDeposit))
                If withdrawAmount >= minAmount Or depositAmount >= minAmount Then
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
