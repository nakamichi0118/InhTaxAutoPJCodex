Option Explicit

Private gDebugWs As Worksheet
Private gDebugRow As Long
Private Const DEBUG_LOG_ENABLED As Boolean = True

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

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
    Dim minAmount As Long
    Dim csvText As String
    Dim csvData As Variant
    Dim usageData As Variant

    Call InitPdfDebugLog(targetDocType)

    Set ws = ActiveSheet
    buttonCol = ws.Shapes(Application.Caller).TopLeftCell.Column
    buttonRow = ws.Shapes(Application.Caller).TopLeftCell.Row

    pdfPath = SelectPdfFile()
    If pdfPath = "" Then Exit Sub

    minAmount = GetMinimumAmount()
    If minAmount = -1 Then Exit Sub

    csvText = FetchCsvTextFromApi(pdfPath, targetDocType)
    If Len(csvText) = 0 Then
        MsgBox "PDF の読み取りに失敗しました。設定値とネットワークを確認してください。", vbExclamation
        Exit Sub
    End If

    usageData = ParseTransactionCsvContent(csvText, 0)
    csvData = ParseCsvText(csvText, minAmount)
    If IsEmpty(csvData) Then
        MsgBox "指定金額以上の取引は見つかりませんでした。", vbInformation
        Exit Sub
    End If

    Call ImportDataToExcel(csvData, buttonCol, buttonRow, usageData)
    MsgBox "PDF の取り込みが完了しました。", vbInformation
End Sub

Public Sub InitPdfDebugLog(Optional ByVal scenarioName As String = "")
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim found As Boolean
    Dim i As Long

    If Not DEBUG_LOG_ENABLED Then Exit Sub

    Set wb = ThisWorkbook
    found = False

    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).Name = "PDF取込ログ" Then
            Set ws = wb.Worksheets(i)
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "PDF取込ログ"
    End If

    ws.Cells.Clear
    ws.Range("A1").Value = "Step"
    ws.Range("B1").Value = "LineIndex"
    ws.Range("C1").Value = "RawLine"
    ws.Range("D1").Value = "RawWithdraw"
    ws.Range("E1").Value = "RawDeposit"
    ws.Range("F1").Value = "ParsedWithdraw"
    ws.Range("G1").Value = "ParsedDeposit"
    ws.Range("H1").Value = "MinAmount"
    ws.Range("I1").Value = "PassedFilter"
    ws.Range("J1").Value = "Note"

    gDebugRow = 1
    Set gDebugWs = ws
    gDebugRow = gDebugRow + 1
    gDebugWs.Cells(gDebugRow, "A").Value = "Start"
    gDebugWs.Cells(gDebugRow, "C").Value = "Scenario"
    gDebugWs.Cells(gDebugRow, "D").Value = scenarioName
    gDebugWs.Cells(gDebugRow, "E").Value = Now
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
        .Cells(gDebugRow, "A").Value = stepName
        .Cells(gDebugRow, "B").Value = lineIndex
        .Cells(gDebugRow, "C").Value = rawLine
        .Cells(gDebugRow, "D").Value = rawWithdraw
        .Cells(gDebugRow, "E").Value = rawDeposit
        .Cells(gDebugRow, "F").Value = parsedWithdraw
        .Cells(gDebugRow, "G").Value = parsedDeposit
        .Cells(gDebugRow, "H").Value = minAmount
        .Cells(gDebugRow, "I").Value = IIf(passed, "TRUE", "FALSE")
        .Cells(gDebugRow, "J").Value = note
    End With
End Sub

Private Function SelectPdfFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "PDF ファイルを選択してください"
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

Private Function FetchCsvTextFromApi(pdfPath As String, overrideDocType As String) As String
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
    Dim csvBase64 As String
    Dim displayName As String

    baseUrl = GetConfigValue("BASE_URL")
    apiKey = GetConfigValue("API_KEY")
    docType = ResolveDocumentType(overrideDocType)
    If Len(docType) = 0 Then GoTo Cleanup
    dateFmt = GetOptionalConfigValue("DATE_FORMAT", "auto")
    maxWaitSeconds = CLng(GetOptionalConfigValue("JOB_MAX_WAIT_SECONDS", "900"))
    pollIntervalMs = CLng(GetOptionalConfigValue("JOB_POLL_INTERVAL_MS", "4000"))
    If pollIntervalMs < 500 Then pollIntervalMs = 500

    normalizedBase = NormalizeBaseUrl(baseUrl)
    displayName = ExtractFileName(pdfPath)
    jobId = CreateAnalysisJob(normalizedBase & "/jobs", pdfPath, docType, dateFmt, apiKey)
    If Len(jobId) = 0 Then GoTo Cleanup

    Application.StatusBar = "ファイル: " & displayName & " ｜ 解析を開始しました"
    statusUrl = normalizedBase & "/jobs/" & jobId
    resultUrl = statusUrl & "/result"
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

    csvBase64 = ExtractCsvBase64(resultJson, "bank_transactions.csv")
    If Len(csvBase64) = 0 Then
        csvBase64 = ExtractFirstFileBase64(resultJson)
    End If
    FetchCsvTextFromApi = Base64ToUtf8(csvBase64)

Cleanup:
    Application.StatusBar = False
    Exit Function

ErrHandler:
    MsgBox "API 呼び出しでエラー: " & Err.Description, vbCritical
    FetchCsvTextFromApi = ""
    Resume Cleanup
End Function

Private Function CreateAnalysisJob(endpoint As String, pdfPath As String, docType As String, _
    dateFmt As String, apiKey As String) As String
    Dim boundary As String
    Dim body() As Byte
    Dim http As Object
    Dim jobId As String
    Dim responseText As String

    boundary = "----SOROBOCR" & Format(Now, "yymmddhhmmss")
    body = BuildMultipartBody(pdfPath, boundary, docType, dateFmt)

    Set http = CreateHttpClient(120000)
    http.Open "POST", endpoint, False
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.SetRequestHeader "Accept", "application/json"
    If Len(apiKey) > 0 Then
        http.SetRequestHeader "X-API-Key", apiKey
    End If
    http.Send body
    responseText = ReadUtf8Response(http)

    If http.Status <> 200 And http.Status <> 202 Then
        MsgBox "ジョブ作成に失敗しました: " & http.Status & " " & http.StatusText & _
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
    http.SetRequestHeader "Accept", "application/json"
    If Len(apiKey) > 0 Then
        http.SetRequestHeader "X-API-Key", apiKey
    End If
    http.Send
    responseText = ReadUtf8Response(http)

    If http.Status <> 200 Then
        MsgBox "API応答: " & http.Status & " " & http.StatusText & _
               vbCrLf & responseText, vbExclamation
        Exit Function
    End If

    SendJsonRequest = responseText
    Exit Function

ErrHandler:
    MsgBox "API 呼び出しでエラー: " & Err.Description, vbCritical
    SendJsonRequest = ""
End Function

Private Function CreateHttpClient(receiveTimeoutMs As Long) As Object
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Option(6) = True
    http.Option(12) = False
    http.Option(4) = 13056
    On Error Resume Next
    http.SetTimeouts 30000, 60000, 60000, receiveTimeoutMs
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
    stream.Write http.ResponseBody
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    ReadUtf8Response = stream.ReadText
    stream.Close
    Set stream = Nothing
    Exit Function
Fallback:
    ReadUtf8Response = http.ResponseText
End Function

Private Function NormalizeBaseUrl(baseUrl As String) As String
    Dim tmp As String
    tmp = Trim(baseUrl)
    If Right(tmp, 1) = "/" Then tmp = Left(tmp, Len(tmp) - 1)
    NormalizeBaseUrl = tmp
End Function

Private Function BuildMultipartBody(pdfPath As String, boundary As String, _
    Optional docType As String = "", Optional dateFmt As String = "") As Byte()
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

Private Function Base64ToUtf8(base64Value As String) As String
    If Len(base64Value) = 0 Then Exit Function
    Dim xml As Object
    Dim node As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = base64Value
    Dim bytes() As Byte
    bytes = node.nodeTypedValue
    Base64ToUtf8 = StrConv(bytes, vbUnicode)
    If Left$(Base64ToUtf8, 1) = ChrW(&HFEFF) Then
        Base64ToUtf8 = Mid$(Base64ToUtf8, 2)
    End If
End Function

Private Function ExtractCsvBase64(json As String, fileName As String) As String
    Dim token As String
    Dim startPos As Long

    token = """" & fileName & """:"""
    startPos = InStr(1, json, token, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(token)

    ExtractCsvBase64 = ExtractJsonStringAt(json, startPos)
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
        rawLine = RemoveUtf8Bom(lines(i))
        If Len(Trim$(rawLine)) > 0 Then
            lineData = PdfSplitCsvLine(rawLine)
            If UBound(lineData) >= 3 Then
                transDate = lineData(0)
                description = CleanDescriptionText(lineData(1))
                rawWithdraw = lineData(2)
                rawDeposit = lineData(3)
                withdrawAmount = PdfToLong(rawWithdraw)
                depositAmount = PdfToLong(rawDeposit)
                passed = (Abs(withdrawAmount) >= minAmount Or Abs(depositAmount) >= minAmount)
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
    PdfSplitCsvLine = results.ToArray
End Function

Private Function PdfToLong(valueText As String) As Long
    Dim cleaned As String
    cleaned = RemoveUtf8Bom(valueText)
    cleaned = Replace(cleaned, "－", "-")
    cleaned = Replace(cleaned, "−", "-")
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

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        If Trim$(ws.Cells(i, 1).Value) = keyName Then
            GetConfigValue = Trim$(ws.Cells(i, 2).Value)
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
