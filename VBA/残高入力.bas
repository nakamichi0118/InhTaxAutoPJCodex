Private m_clsResizer As CResizer
Private Sub UserForm_Initialize()
    Set m_clsResizer = New CResizer
    m_clsResizer.Add Me
    With ListBox1
        .AddItem "平成"
        .AddItem "令和"
    End With
    Me.ListBox1.ListIndex = 0
    Me.Height = 204
    Me.Width = 224
    Me.Zoom = 100
End Sub
Private Sub UserForm_Resize()
    If Me.Width < 240 Then
        Me.Width = 240
    End If
    If Me.Height < 203 Then
        Me.Height = 203
    End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "［閉じる］ボタンを使用してください"
        Cancel = True
    End If
End Sub
Private Sub money_Change()
    Static bProcessing As Boolean
    If bProcessing Then Exit Sub

    bProcessing = True
    If IsNumeric(Me.money.text) Then
        Dim formatted As String
        formatted = Format(Me.money.text, "#,##0")
        If Me.money.text <> formatted Then
            Me.money.text = formatted
        End If
    End If
    bProcessing = False
End Sub
Private Sub CommandButton2_Click()
Debug.Print "run commanButton2_click"
Dim retuhajime
Dim retusaigo
Dim gyouhajime
Dim gyousaigo
Dim gyou
Dim zengyou
Dim hidukekazu
Dim ColumnChS
Dim ColumnChN

Dim DayCo '日付の列
DayCo = "C"

Dim SerchCr '検索文字列
SerchCr = "A"

Dim SerchTop

Set suii = Worksheets("預金推移")
gyouhajime = Range("A1:A10000").Find("資金移動始").row
gyousaigo = Range("A1:A10000").Find("資金移動終").row
retuhajime = Range("C1:BZ1").Find("資金移動始").Column
retusaigo = Range("C1:DZ1").Find("資金移動終").Column

' データがない場合（資金移動始と資金移動終の間が3行以下）は終了
If gyousaigo - gyouhajime <= 3 Then
    Unload 残高入力
    Exit Sub
End If

retu = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column '入力列番号の取得

Dim rng As Range
Dim TempRng As Range
Dim SerchRng As Range
Dim j
j = 1
'閉じるときの処理　①i=0　②ソート　③合計の表示　④A列に検索の為の数式を入力　⑤閉じる
'①
i = 0
'②

' データがある場合のみソートとカラーサム処理を実行
If gyousaigo - gyouhajime > 3 Then
With ActiveWorkbook.Worksheets("預金推移")
    Range(.Cells(gyouhajime + 1, retuhajime), .Cells(gyousaigo - 3, retusaigo)).Sort Key1:=.Columns(3)
    gyouhajime = Range("A1:A10000").Find("資金移動始").row
    gyousaigo = Range("A1:A10000").Find("資金移動終").row
    retuhajime = Range("C1:BZ1").Find("資金移動始").Column
    retusaigo = Range("C1:DZ1").Find("資金移動終").Column
'③カラーサムの表示
    For i = 1 To retusaigo - 3
        '親族間移動（緑）
        .Cells(gyousaigo - 1, retuhajime + i).value = "=IFERROR(colorsum(" & .Cells(gyouhajime + 1, retuhajime + i).Address & ":" & .Cells(gyousaigo - 3, retuhajime + i).Address & ",$I$2),"""")"
    Next
    For i = 1 To retusaigo - 3
        '不明入出金（黄色）
        '親族間移動（緑）
        .Cells(gyousaigo - 2, retuhajime + i).value = "=IFERROR(colorsum(" & .Cells(gyouhajime + 1, retuhajime + i).Address & ":" & .Cells(gyousaigo - 3, retuhajime + i).Address & ",$I$2),"""")"
        
        If i Mod 2 = 0 Then
         '不明入出金（黄色）
         .Cells(gyousaigo - 1, retuhajime + i).value = "=IFERROR(colorsum(" & .Cells(gyouhajime + 1, retuhajime + i).Address & ":" & .Cells(gyousaigo - 3, retuhajime + i).Address & ",$F$3),"""")"
        Else
         '不明入出金（ピンク）
         .Cells(gyousaigo - 1, retuhajime + i).value = "=IFERROR(colorsum(" & .Cells(gyouhajime + 1, retuhajime + i).Address & ":" & .Cells(gyousaigo - 3, retuhajime + i).Address & ",$F$2),"""")"
        End If
    Next

    
End With
End If

ColumnChS = Replace(Cells(1, retu).Address(True, False), "$1", "") '出金列のアルファベット取得
ColumnChN = Replace(Cells(1, retu + 1).Address(True, False), "$1", "") '入金列のアルファベット取得

For i = 1 To gyousaigo - gyouhajime - 3
    suii.Cells(gyouhajime + i, 1) = "=mid(" & DayCo & gyouhajime + i & ",find(""(""," & DayCo & gyouhajime + i & ")+1,4)"
Next


Range(Cells(gyouhajime, retuhajime), Cells(gyousaigo, retusaigo)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'④合計のところに入力 　★この処理は資金移動の方も必要では？
Set FoundCell = Range(Cells(gyouhajime + 1, 3), Cells(gyousaigo - 3, 3)).CurrentRegion.Find(What:="の入出金")

' 「の入出金」行がある場合のみ処理
If Not FoundCell Is Nothing Then
    hidukekazu = WorksheetFunction.CountIf(Range("c1:c10000"), "*の入出金")

    Set SerchRng = Range("c1:c10000")
    Set rng = SerchRng.Find("*の入出金")

    If Not rng Is Nothing Then
        Set TempRng = rng
        gyou = rng.row  ' 重複Find呼び出しを削除

        Do
            ' セルの書式をクリアして数式として認識させる
            suii.Cells(gyou, retu).NumberFormat = "General"
            suii.Cells(gyou, retu + 1).NumberFormat = "General"
            suii.Cells(gyou, retu).Formula = "=SumIf(" & SerchCr & gyouhajime + 1 & ":" & SerchCr & gyou - 1 & ",mid(" & DayCo & gyou & ",FIND(""(""," & DayCo & gyou & ")+1,4)," & ColumnChS & gyouhajime + 1 & ":" & ColumnChS & gyou - 1 & ")"
            suii.Cells(gyou, retu + 1).Formula = "=SumIf(" & SerchCr & gyouhajime + 1 & ":" & SerchCr & gyou - 1 & ",mid(" & DayCo & gyou & ",FIND(""(""," & DayCo & gyou & ")+1,4)," & ColumnChN & gyouhajime + 1 & ":" & ColumnChN & gyou - 1 & ")"
            suii.Cells(gyou, retu).NumberFormat = "#,##0"
            suii.Cells(gyou, retu + 1).NumberFormat = "#,##0"
            suii.Cells(gyou, retu).HorizontalAlignment = xlRight '右寄せ
            suii.Cells(gyou, retu + 1).HorizontalAlignment = xlRight
            suii.Cells(gyou, retu).Font.Color = RGB(255, 0, 0)
            suii.Cells(gyou, retu + 1).Font.Color = RGB(31, 73, 125)
            suii.Range(suii.Cells(gyou, retuhajime), suii.Cells(gyou + 1, retusaigo)).Interior.Color = RGB(220, 220, 220)
            suii.Range(suii.Cells(gyou, retuhajime), suii.Cells(gyou + 1, retusaigo)).Font.Bold = True
            suii.Range(suii.Cells(gyou, retuhajime), suii.Cells(gyou + 1, retusaigo)).BorderAround LineStyle:=xlDouble

            Set rng = SerchRng.FindNext(rng)
            If rng Is Nothing Then Exit Do
            If rng.Address = TempRng.Address Then Exit Do

            zengyou = gyou
            gyou = rng.row

            suii.Cells(gyou + 1, retu) = "=" & ColumnChN & gyou + 1 & "-" & ColumnChN & zengyou + 1
            suii.Cells(gyou + 1, retu).NumberFormatLocal = """増減""#,##0_ ;""増減 ""-#,##0"

            j = j + 1
        Loop
    End If
End If
    suii.Range(suii.Cells(gyouhajime, retuhajime), suii.Cells(gyousaigo, retuhajime)).Borders(xlEdgeLeft).Weight = xlThick
    suii.Range(suii.Cells(gyouhajime, retusaigo), suii.Cells(gyousaigo, retusaigo)).Borders(xlEdgeRight).Weight = xlThick
Unload 残高入力
End Sub
Private Sub touroku_Click()
Debug.Print "run touroku_click"
Dim hiduke
Dim kakohiduke
Dim FoundCell As Range    ''またはバリアント型(Variant)とする
Dim lastRow As Long
Dim retu
Dim gyou
Dim hidukekazu
Dim suii
Dim retuhajime
Dim retusaigo
Dim gyouhajime
Dim gyousaigo

Dim DayCo '日付の列
DayCo = "C"

Dim SerchCr '検索文字列
SerchCr = "A"

Set suii = Worksheets("預金推移")
gyouhajime = Range("A1:A10000").Find("資金移動始").row
gyousaigo = Range("A1:A10000").Find("資金移動終").row
retuhajime = Range("C1:BZ1").Find("資金移動始").Column
retusaigo = Range("C1:DZ1").Find("資金移動終").Column

retu = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column '入力列番号の取得


'入力文字列が1文字だった場合0を付す

If ListBox1 = "平成" Then
    hiduke = "H" & year & "(" & year + 1988 & ")"
ElseIf ListBox1 = "令和" Then
    hiduke = "R" & year & "(" & year + 2018 & ")"
End If

'1.入力項目の確認
If year = "" Then
    MsgBox "年を入力してください"
ElseIf IsNull(ListBox1.value) Then
     MsgBox "年号を選択してくだいさい"
ElseIf money = "" Then
    MsgBox "金額を入力してください"

Else '入力データの登録
Set FoundCell = Range(Cells(gyouhajime + 1, 3), Cells(gyousaigo - 3, 3)).CurrentRegion.Find(What:=hiduke & "の入出金")
    If FoundCell Is Nothing Then '同じ日付がなかった場合
        lastRow = Cells(Rows.count, 3).End(xlUp).row - 2
        Rows(lastRow).Insert Shift:=xlUp
        Rows(lastRow).Insert Shift:=xlUp
        Rows(lastRow).Interior.ColorIndex = 0

        With Worksheets("預金推移")
            .Cells(lastRow, 3).value = hiduke & "の入出金"
            .Cells(lastRow + 1, 3).value = hiduke & "残高"
            .Cells(lastRow + 1, retu + 1) = money
            .Cells(lastRow + 1, retu + 1).NumberFormatLocal = "残高　@"
            .Cells(lastRow + 1, retu + 1).HorizontalAlignment = xlRight '右寄せ
            .Cells(lastRow + 1, retu).value = "-"
            .Cells(lastRow + 1, retu).HorizontalAlignment = xlCenter
        End With
    Else
    '同じ日付があった場合
    gyou = WorksheetFunction.match(hiduke & "の入出金", Range("C1:C10000"), 0)
        With Worksheets("預金推移")
            .Cells(gyou + 1, retu + 1).value = Format(money, "#,###") '
            .Cells(gyou + 1, retu + 1).HorizontalAlignment = xlRight '右寄せ
            .Cells(gyou + 1, retu).value = "-"
            .Cells(gyou + 1, retu).HorizontalAlignment = xlCenter
        End With
    End If
    year.value = ""
    money.value = ""
End If
End Sub
Private Sub UserForm_Terminate()
 
    Set m_clsResizer = Nothing
    
End Sub
