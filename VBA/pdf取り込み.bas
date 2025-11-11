Option Explicit

'====================================================================================
' PDF 直接取り込みモジュール
'
' ・隠しシート「設定」の A 列にキー、B 列に値を格納しておくこと。
'   例) A1=BASE_URL, B1=https://inhtaxautopjcodex-production.up.railway.app/api
'       A2=API_KEY,   B2= (必要なら任意のキー。未使用なら空欄で可)
' ・Cloudflare Pages 経由/ Railway 直割りどちらでも BASE_URL で切り替え可能。
' ・参照設定は不要。(WinHTTP/ADODB/DOMDocument は Late Binding)
'
' ※既存の CSV 取り込みロジック (GetMinimumAmount, ImportDataToExcel など) を流用。
'====================================================================================

Sub PDF取込ボタン_Click()
    Dim ws As Worksheet
    Dim buttonCol As Long
    Dim pdfPath As String
    Dim minAmount As Long
    Dim csvText As String
    Dim csvData As Variant

    Set ws = ActiveSheet
    buttonCol = ws.Shapes(Application.Caller).TopLeftCell.Column

    pdfPath = SelectPdfFile()
    If pdfPath = "" Then Exit Sub

    minAmount = GetMinimumAmount()
    If minAmount = -1 Then Exit Sub

    csvText = FetchCsvTextFromApi(pdfPath)
    If Len(csvText) = 0 Then
        MsgBox "PDF の読み取りに失敗しました。設定値とネットワークを確認してください。", vbExclamation
        Exit Sub
    End If

    csvData = ParseCsvText(csvText, minAmount)
    If IsEmpty(csvData) Then
        MsgBox "指定金額以上の取引は見つかりませんでした。", vbInformation
        Exit Sub
    End If

    Call ImportDataToExcel(csvData, buttonCol)
    MsgBox "PDF の取り込みが完了しました。", vbInformation
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

Private Function FetchCsvTextFromApi(pdfPath As String) As String
    On Error GoTo ErrHandler
    Dim baseUrl As String
    Dim apiKey As String
    Dim endpoint As String
    Dim boundary As String
    Dim body() As Byte
    Dim http As Object
    Dim responseText As String
    Dim csvBase64 As String

    baseUrl = GetConfigValue("BASE_URL")
    apiKey = GetConfigValue("API_KEY")
    endpoint = NormalizeBaseUrl(baseUrl) & "/documents/analyze-export"

    boundary = "----SOROBOCR" & Format(Now, "yymmddhhmmss")
    body = BuildMultipartBody(pdfPath, boundary)

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", endpoint, False
    http.Option(6) = True
    http.Option(12) = False
    http.Option(4) = 13056
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    If Len(apiKey) > 0 Then
        http.SetRequestHeader "X-API-Key", apiKey
    End If
    http.Send body

    If http.Status <> 200 Then
        MsgBox "API応答: " & http.Status & " " & http.StatusText, vbExclamation
        Exit Function
    End If

    responseText = http.ResponseText
    csvBase64 = ExtractCsvBase64(responseText, "bank_transactions.csv")
    If csvBase64 = "" Then
        csvBase64 = ExtractFirstFileBase64(responseText)
    End If
    FetchCsvTextFromApi = Base64ToUtf8(csvBase64)
    Exit Function

ErrHandler:
    MsgBox "API 呼び出しでエラー: " & Err.Description, vbCritical
    FetchCsvTextFromApi = ""
End Function

Private Function NormalizeBaseUrl(baseUrl As String) As String
    Dim tmp As String
    tmp = Trim(baseUrl)
    If Right(tmp, 1) = "/" Then tmp = Left(tmp, Len(tmp) - 1)
    NormalizeBaseUrl = tmp
End Function

Private Function BuildMultipartBody(pdfPath As String, boundary As String) As Byte()
    Dim fileBytes() As Byte
    Dim prefix As String
    Dim suffix As String
    Dim fileName As String
    Dim stream As Object

    fileBytes = ReadBinaryFile(pdfPath)
    fileName = Mid$(pdfPath, InStrRev(pdfPath, Application.PathSeparator) + 1)

    prefix = "--" & boundary & vbCrLf & _
             "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf & _
             "Content-Type: application/pdf" & vbCrLf & vbCrLf

    suffix = vbCrLf & "--" & boundary & "--" & vbCrLf

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 'binary
    stream.Open
    stream.Write StringToBytes(prefix)
    stream.Write fileBytes
    stream.Write StringToBytes(suffix)

    stream.Position = 0
    BuildMultipartBody = stream.Read
    stream.Close
    Set stream = Nothing
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
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 'text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textValue
    stream.Position = 0
    stream.Type = 1
    StringToBytes = stream.Read
    stream.Close
    Set stream = Nothing
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
    Dim i As Long
    Dim resultText As String

    token = """" & fileName & """:"""
    startPos = InStr(1, json, token, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(token)

    For i = startPos To Len(json)
        Dim ch As String
        ch = Mid$(json, i, 1)
        If ch = """" And Mid$(json, i - 1, 1) <> "\" Then Exit For
        resultText = resultText & ch
    Next i
    ExtractCsvBase64 = resultText
End Function

Private Function ExtractFirstFileBase64(json As String) As String
    Dim token As String
    Dim pos As Long

    token = """files"":{"
    pos = InStr(1, json, token, vbTextCompare)
    If pos = 0 Then Exit Function
    pos = pos + Len(token)
    pos = InStr(pos, json, """:"", vbTextCompare)
    If pos = 0 Then Exit Function
    pos = pos + 3
    Dim i As Long
    Dim resultText As String
    For i = pos To Len(json)
        Dim ch As String
        ch = Mid$(json, i, 1)
        If ch = """" And Mid$(json, i - 1, 1) <> "\" Then Exit For
        resultText = resultText & ch
    Next i
    ExtractFirstFileBase64 = resultText
End Function

Private Function ParseCsvText(csvContent As String, minAmount As Long) As Variant
    Dim lines() As String
    Dim resultData() As Variant
    Dim lineData() As String
    Dim i As Long
    Dim dataCount As Long
    Dim transDate As String
    Dim withdrawAmount As Long
    Dim depositAmount As Long

    lines = Split(csvContent, vbLf)
    dataCount = 0
    ReDim resultData(1 To 10000, 1 To 3)

    For i = 1 To UBound(lines) ' skip header at index 0
        If Trim$(lines(i)) <> "" Then
            lineData = SplitCsvLine(lines(i))
            If UBound(lineData) >= 4 Then
                transDate = lineData(0)
                withdrawAmount = ToLong(lineData(2))
                depositAmount = ToLong(lineData(3))
                If withdrawAmount >= minAmount Or depositAmount >= minAmount Then
                    dataCount = dataCount + 1
                    resultData(dataCount, 1) = ConvertDateFormat(transDate)
                    resultData(dataCount, 2) = withdrawAmount
                    resultData(dataCount, 3) = depositAmount
                End If
            End If
        End If
    Next i

    If dataCount = 0 Then
        ParseCsvText = Empty
    Else
        Dim finalData() As Variant
        ReDim finalData(1 To dataCount, 1 To 3)
        For i = 1 To dataCount
            finalData(i, 1) = resultData(i, 1)
            finalData(i, 2) = resultData(i, 2)
            finalData(i, 3) = resultData(i, 3)
        Next i
        ParseCsvText = finalData
    End If
End Function

Private Function SplitCsvLine(lineText As String) As String()
    Dim cleaned As String
    cleaned = Replace(lineText, vbCr, "")
    cleaned = Replace(cleaned, """", "")
    SplitCsvLine = Split(cleaned, ",")
End Function

Private Function ToLong(valueText As String) As Long
    Dim trimmed As String
    trimmed = Replace(Replace(Trim$(valueText), ",", ""), """", "")
    If trimmed = "" Then
        ToLong = 0
    Else
        If InStr(trimmed, ".") > 0 Then
            trimmed = Left$(trimmed, InStr(trimmed, ".") - 1)
        End If
        If IsNumeric(trimmed) Then
            ToLong = CLng(trimmed)
        Else
            ToLong = 0
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
