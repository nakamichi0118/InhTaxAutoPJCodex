VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPdfImportSettings
   Caption         =   "PDF取込設定"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5520
   OleObjectBlob   =   "frmPdfImportSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPdfImportSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' PDF取込設定フォーム
' 複数のダイアログを1つのUserFormに統合
'===============================================================================
Option Explicit

Private mCancelled As Boolean
Private mDocType As String
Private mDateFormat As String
Private mMinAmount As Long
Private mStartDate As String
Private mEndDate As String

'--- プロパティ ---

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

Public Property Get DocType() As String
    DocType = mDocType
End Property

Public Property Get DateFormat() As String
    DateFormat = mDateFormat
End Property

Public Property Get MinAmount() As Long
    MinAmount = mMinAmount
End Property

Public Property Get StartDate() As String
    StartDate = mStartDate
End Property

Public Property Get EndDate() As String
    EndDate = mEndDate
End Property

'--- イベントハンドラ ---

Private Sub UserForm_Initialize()
    ' デフォルト値を設定
    mCancelled = True
    mDocType = "bank_deposit"
    mDateFormat = "auto"
    mMinAmount = 500000
    mStartDate = ""
    mEndDate = ""

    ' コントロールの初期化
    optDocTypeBank.Value = True
    optDateFormatAuto.Value = True
    txtMinAmount.Value = "500000"
    txtStartDate.Value = ""
    txtEndDate.Value = ""
End Sub

Private Sub cmdOK_Click()
    ' 入力値の検証
    If Not ValidateInputs() Then
        Exit Sub
    End If

    ' 値を保存
    If optDocTypeBank.Value Then
        mDocType = "bank_deposit"
    Else
        mDocType = "transaction_history"
    End If

    If optDateFormatAuto.Value Then
        mDateFormat = "auto"
    ElseIf optDateFormatWareki.Value Then
        mDateFormat = "auto"  ' 和暦は auto で処理
    Else
        mDateFormat = "western"
    End If

    mMinAmount = CLng(txtMinAmount.Value)
    mStartDate = Trim$(txtStartDate.Value)
    mEndDate = Trim$(txtEndDate.Value)

    mCancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    mCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        mCancelled = True
    End If
End Sub

Private Function ValidateInputs() As Boolean
    ValidateInputs = False

    ' 最小金額の検証
    If Len(Trim$(txtMinAmount.Value)) = 0 Then
        MsgBox "最小金額を入力してください。", vbExclamation
        txtMinAmount.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtMinAmount.Value) Then
        MsgBox "最小金額は数値で入力してください。", vbExclamation
        txtMinAmount.SetFocus
        Exit Function
    End If

    If CLng(txtMinAmount.Value) < 0 Then
        MsgBox "最小金額は0以上で入力してください。", vbExclamation
        txtMinAmount.SetFocus
        Exit Function
    End If

    ' 開始日の検証（入力されている場合のみ）
    If Len(Trim$(txtStartDate.Value)) > 0 Then
        If Not IsValidDateFormat(txtStartDate.Value) Then
            MsgBox "開始日は YYYY-MM-DD または YYYY/MM/DD 形式で入力してください。" & vbCrLf & _
                   "例: 2021-04-01 または 2021/04/01", vbExclamation
            txtStartDate.SetFocus
            Exit Function
        End If
    End If

    ' 終了日の検証（入力されている場合のみ）
    If Len(Trim$(txtEndDate.Value)) > 0 Then
        If Not IsValidDateFormat(txtEndDate.Value) Then
            MsgBox "終了日は YYYY-MM-DD または YYYY/MM/DD 形式で入力してください。" & vbCrLf & _
                   "例: 2024-12-31 または 2024/12/31", vbExclamation
            txtEndDate.SetFocus
            Exit Function
        End If
    End If

    ' 開始日と終了日の整合性チェック
    If Len(Trim$(txtStartDate.Value)) > 0 And Len(Trim$(txtEndDate.Value)) > 0 Then
        Dim startDt As Date, endDt As Date
        startDt = ParseDateString(txtStartDate.Value)
        endDt = ParseDateString(txtEndDate.Value)
        If startDt > endDt Then
            MsgBox "終了日は開始日以降の日付を指定してください。", vbExclamation
            txtEndDate.SetFocus
            Exit Function
        End If
    End If

    ValidateInputs = True
End Function

Private Function IsValidDateFormat(dateStr As String) As Boolean
    Dim cleanDate As String
    cleanDate = Replace(dateStr, "/", "-")

    Dim parts() As String
    parts = Split(cleanDate, "-")

    If UBound(parts) <> 2 Then
        IsValidDateFormat = False
        Exit Function
    End If

    On Error GoTo InvalidDate

    Dim year As Integer, month As Integer, day As Integer
    year = CInt(parts(0))
    month = CInt(parts(1))
    day = CInt(parts(2))

    If year < 1900 Or year > 2100 Then GoTo InvalidDate
    If month < 1 Or month > 12 Then GoTo InvalidDate
    If day < 1 Or day > 31 Then GoTo InvalidDate

    ' 日付として有効かチェック
    Dim testDate As Date
    testDate = DateSerial(year, month, day)

    IsValidDateFormat = True
    Exit Function

InvalidDate:
    IsValidDateFormat = False
End Function

Private Function ParseDateString(dateStr As String) As Date
    Dim cleanDate As String
    cleanDate = Replace(dateStr, "/", "-")

    Dim parts() As String
    parts = Split(cleanDate, "-")

    ParseDateString = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
End Function

'--- 公開メソッド ---

Public Function NormalizeDateForApi() As String
    ' API送信用に日付を YYYY-MM-DD 形式に正規化
    If Len(mStartDate) > 0 Then
        NormalizeDateForApi = Replace(mStartDate, "/", "-")
    Else
        NormalizeDateForApi = ""
    End If
End Function

Public Function NormalizeEndDateForApi() As String
    ' API送信用に日付を YYYY-MM-DD 形式に正規化
    If Len(mEndDate) > 0 Then
        NormalizeEndDateForApi = Replace(mEndDate, "/", "-")
    Else
        NormalizeEndDateForApi = ""
    End If
End Function
