Private Sub money_Change()
If IsNumeric(Me.money.text) Then
Me.money.text = Format(Me.money.text, "#,##0")
End If
End Sub
Private Sub CommandButton2_Click()
    Dim retusaigo
    Dim gyouhajime
    Dim gyousaigo
    Dim gyou
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
    
    If gyouhajime = 13 And gyousaigo = 16 Then
        Unload 資金移動
        Exit Sub
    End If
    
    retu = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column '入力列番号の取得
    
    Dim rng As Range
    Dim TempRng As Range
    Dim SerchRng As Range
    
    '閉じるときの処理　①i=0　②ソート　③合計の表示　④A列に検索の為の数式を入力　⑤閉じる
    '①
    i = 0
    '②
    
    
    With ActiveWorkbook.Worksheets("預金推移")
        Range(.Cells(gyouhajime + 1, retuhajime), .Cells(gyousaigo - 3, retusaigo)).Sort Key1:=.Columns(3)
        gyouhajime = Range("A1:A10000").Find("資金移動始").row
        gyousaigo = Range("A1:A10000").Find("資金移動終").row
        retuhajime = Range("C1:BZ1").Find("資金移動始").Column
        retusaigo = Range("C1:DZ1").Find("資金移動終").Column
    '③カラーサムの表示
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

    
    ColumnChS = Replace(Cells(1, retu).Address(True, False), "$1", "") '出金列のアルファベット取得
    ColumnChN = Replace(Cells(1, retu + 1).Address(True, False), "$1", "") '入金列のアルファベット取得
    
    For i = 1 To gyousaigo - gyouhajime - 3
        suii.Cells(gyouhajime + i, 1) = "=mid(" & DayCo & gyouhajime + i & ",find(""(""," & DayCo & gyouhajime + i & ")+1,4)"
    Next
    
    
    Range(Cells(gyouhajime, retuhajime), Cells(gyousaigo, retusaigo)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    '④合計のところに入力 　★この処理は資金移動の方も必要では？
    Set FoundCell = Range(Cells(gyouhajime + 1, 3), Cells(gyousaigo - 3, 3)).CurrentRegion.Find(What:="の入出金")
    
    If FoundCell Is Nothing Then
    Else
        hidukekazu = WorksheetFunction.CountIf(Range("c1:c10000"), "*の入出金")
        
        Set SerchRng = Range("c1:c10000")
        
        Set rng = SerchRng.Find("*の入出金")
        
        Set TempRng = rng
        
        gyou = SerchRng.Find("*の入出金").row
        
        
        Do While Not rng Is Nothing
            
    
            suii.Cells(gyou, retu) = "=SumIf(" & SerchCr & gyouhajime + 1 & ":" & SerchCr & gyou - 1 & ",mid(" & DayCo & gyou & ",FIND(""(""" & "," & DayCo & gyou & ")+1,4)," & ColumnChS & gyouhajime + 1 & ":" & ColumnChS & gyou - 1 & ")"
            suii.Cells(gyou, retu + 1) = "=SumIf(" & SerchCr & gyouhajime + 1 & ":" & SerchCr & gyou - 1 & ",mid(" & DayCo & gyou & ",FIND(""(""" & "," & DayCo & gyou & ")+1,4)," & ColumnChN & gyouhajime + 1 & ":" & ColumnChN & gyou - 1 & ")"
            suii.Cells(gyou, retu).HorizontalAlignment = xlRight '右寄せ
            suii.Cells(gyou, retu + 1).HorizontalAlignment = xlRight
            suii.Cells(gyou, retu).Font.Color = RGB(255, 0, 0)
            suii.Cells(gyou, retu + 1).Font.Color = RGB(31, 73, 125)
            suii.Range(suii.Cells(gyou, retuhajime), suii.Cells(gyou + 1, retusaigo)).Interior.Color = RGB(191, 191, 191)
            suii.Range(suii.Cells(gyou, retuhajime), suii.Cells(gyou + 1, retusaigo)).Font.Bold = True
    
            suii.Range(suii.Cells(gyou, retuhajime), suii.Cells(gyou + 1, retusaigo)).BorderAround LineStyle:=xlDouble
            
            Set rng = SerchRng.FindNext(rng)
            zengyou = gyou
            gyou = rng.row
    
            
            If rng.Address = TempRng.Address Then
                Exit Do
            End If
            
            suii.Cells(gyou + 1, retu) = "=" & ColumnChN & gyou + 1 & "-" & ColumnChN & zengyou + 1
            suii.Cells(gyou + 1, retu).NumberFormatLocal = """増減""#,##0_ ;""増減 ""-#,##0"
            
            j = j + 1
        Loop
    End If
        suii.Range(suii.Cells(gyouhajime, retuhajime), suii.Cells(gyousaigo, retuhajime)).Borders(xlEdgeLeft).Weight = xlThick
        suii.Range(suii.Cells(gyouhajime, retusaigo), suii.Cells(gyousaigo, retusaigo)).Borders(xlEdgeRight).Weight = xlThick
        
    '3/5追加　罫線色付け
        Dim ButtonSheet As Worksheet
        Dim startCell As Range, endCell As Range, endColCell As Range
        Dim startRow As Long, endRow As Long, targetCol As Long, lastCol As Long
        Dim q As Long
        Dim currentYear As String
        Dim nextYear As String

        Set ButtonSheet = ThisWorkbook.Worksheets("預金推移")
    
        ' セルの検索
        Set startCell = ButtonSheet.Range("A:A").Find(What:="資金移動始", LookIn:=xlValues, LookAt:=xlWhole)
        Set endCell = ButtonSheet.Range("A:A").Find(What:="資金移動終", LookIn:=xlValues, LookAt:=xlWhole)
        Set endColCell = ButtonSheet.Rows(1).Find(What:="資金移動終", LookIn:=xlValues, LookAt:=xlWhole)
    
        ' いずれかが見つからなければ終了
        If startCell Is Nothing Or endCell Is Nothing Or endColCell Is Nothing Then
            MsgBox "「資金移動始」または「資金移動終」が見つかりませんでした。不具合報告へご連絡お願いします。", vbExclamation
                Unload 資金移動
            Exit Sub
        End If
    
        ' 見つかった場合は行・列番号を取得
        startRow = startCell.row
        endRow = endCell.row
        targetCol = 3
        lastCol = endColCell.Column
        
        '指定範囲内を繰り返し処理
        For q = startRow + 1 To endRow - 4
            currentYear = ExtractYear(ButtonSheet.Cells(q, targetCol).value)
            nextYear = ExtractYear(ButtonSheet.Cells(q + 1, targetCol).value)
        
            '年が異なる場合、赤色罫線    '年が同じ場合、黒色罫線
            If currentYear <> nextYear Then
            
                With ButtonSheet.Range(ButtonSheet.Cells(q, 3), ButtonSheet.Cells(q, lastCol)).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Color = RGB(255, 0, 0)
                    .Weight = xlThin
                End With
            Else
                With ButtonSheet.Range(ButtonSheet.Cells(q, 3), ButtonSheet.Cells(q, lastCol)).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Color = RGB(0, 0, 0)
                    .Weight = xlThin
                End With
            End If
        Next q
        
    '相続境界線（生前贈与7年ルール対応）
        '相続分シートから相続開始日を取得
        Dim inheritanceDate As Variant
        Dim boundaryYear As Integer
        Dim ER_FindRange As Range
        Dim inheritanceBoundary As Date
        Dim newRuleThreshold As Date
        Dim newRuleStartDate As Date

        '相続開始日の取得確認※相続開始日の入力がなければ罫線中断
        Set ER_FindRange = ThisWorkbook.Worksheets("相続分").Range("1:1").Find(What:="相続開始年月日", LookIn:=xlValues, LookAt:=xlWhole)

        If ER_FindRange Is Nothing Then
            MsgBox "相続分シートに「相続開始年月日」の文字が見つかりませんでした。シートをご確認ください。"
            Unload 資金移動
            Exit Sub
        End If

        '値を取得(Variant型で受け取る)
        inheritanceDate = ER_FindRange.Offset(0, 1).value
        '空欄の入力無しで処理する
        If Trim(CStr(inheritanceDate)) = "" Then
            MsgBox "相続分シートに「相続開始年月日の入力がありません。」"
            Unload 資金移動
            Exit Sub
        End If

        ' 2027.1.1を境に判定を変える（生前贈与7年ルール）
        newRuleThreshold = DateSerial(2027, 1, 1)
        newRuleStartDate = DateSerial(2024, 1, 1)

        If CDate(inheritanceDate) < newRuleThreshold Then
            ' 2027.1.1より前の相続開始 → 従来の3年ルール
            inheritanceBoundary = DateAdd("yyyy", -3, inheritanceDate)
        Else
            ' 2027.1.1以降の相続開始 → 7年ルール（経過措置考慮）
            inheritanceBoundary = DateAdd("yyyy", -7, inheritanceDate)
            If inheritanceBoundary < newRuleStartDate Then
                inheritanceBoundary = newRuleStartDate
            End If
        End If

        ' 境界年を取得
        boundaryYear = DatePart("yyyy", inheritanceBoundary)

        '検索範囲として、A列の資金移動始と資金移動終の間を範囲にする
        Dim yearList As Range
        Set yearList = ButtonSheet.Range(ButtonSheet.Range("A:A").Find(What:="資金移動始", LookIn:=xlValues, LookAt:=xlWhole).Offset(1, 0), ButtonSheet.Range("A:A").Find(What:="資金移動終", LookIn:=xlValues, LookAt:=xlWhole).Offset(-3, 0))
        '検索範囲にて、境界日を跨ぐ年度がある場合
        'まず１行目が境界年以上の場合のみ検索実施
        Dim firstFundsYear As Long
        firstFundsYear = ButtonSheet.Range("A:A").Find(What:="資金移動始", LookIn:=xlValues, LookAt:=xlWhole).Offset(1, 0)

        '一つ目の入力が境界年またはそれ以前の場合処理開始
        If firstFundsYear <= boundaryYear Then

                '必要な情報は
                Dim date1 As Variant
                Dim date2 As Variant
                Dim z As Range
                Dim v As Long
                Dim borderDate As Range
            '一つ目の入力項目が境界年の場合
            If firstFundsYear = boundaryYear Then
                'まず、境界年の範囲を作成する
                Dim RangeBoundary As Range
                For Each z In yearList
                    If z.value = boundaryYear Then
                        If RangeBoundary Is Nothing Then
                            Set RangeBoundary = z
                        Else
                            Set RangeBoundary = Union(RangeBoundary, z)
                        End If
                    End If
                Next z
                'この範囲で判定開始、境界日を跨ぐ場合を探す

                For v = 1 To RangeBoundary.Cells.count
                    date1 = dateConverter(RangeBoundary.Cells(v).Offset(0, 2).value)
                    date2 = dateConverter(RangeBoundary.Cells(v + 1).Offset(0, 2).value)

                    If date2 <> 0 Then
                        If IsDate(date1) And IsDate(date2) Then
                            If date1 < inheritanceBoundary And date2 >= inheritanceBoundary Then
                                Set borderDate = RangeBoundary(v)
                                Exit For
                            End If
                        End If
                    End If
                Next v

            '一つ目の入力項目が境界年より前だった場合
            ElseIf firstFundsYear < boundaryYear Then
                '一つ目のセルと二つ目のセルの条件を並べて、目標セルを探す。
                For Each z In yearList
                    date1 = dateConverter(z.Offset(0, 2).value)
                    date2 = dateConverter(z.Offset(1, 2).value)
                    If date2 <> 0 Then
                        If IsDate(date1) And IsDate(date2) Then
                            If date1 < inheritanceBoundary And date2 >= inheritanceBoundary Then
                                Set borderDate = z
                                Exit For
                            End If
                        End If
                    End If
                Next z

            Else
                '一つ目のセルが境界日と同日である場合
            End If

        '見つけた境界線セルの下側に緑色の線を引く

            If Not borderDate Is Nothing Then
                With ButtonSheet.Range(borderDate.Offset(0, 2), ButtonSheet.Cells(borderDate.row, lastCol)).Borders(xlEdgeBottom)
                     .LineStyle = xlContinuous
                     .Color = RGB(0, 255, 0)
                     .Weight = xlThin
                End With
            End If
        End If
    Unload 資金移動
End Sub

Private Sub touroku_Click()
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
    
    Set suii = Worksheets("預金推移")
    gyouhajime = Range("A1:A10000").Find("資金移動始").row
    gyousaigo = Range("A1:A10000").Find("資金移動終").row
    retuhajime = Range("C1:BZ1").Find("資金移動始").Column
    retusaigo = Range("C1:DZ1").Find("資金移動終").Column
    
    retu = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column '入力列番号の取得
    
    '入力文字列が1文字だった場合0を付す
    If Len(month) = 1 Then
        month = "0" & month
    End If
    
    If Len(day) = 1 Then
        day = "0" & day
    End If
    
    If ListBox1 = "平成" Then
        hiduke = Format("H" & year & "(" & year + 1988 & ")" & "/" & month & "/" & day, "ge / m / d")
    
    ElseIf ListBox1 = "令和" Then
        hiduke = Format("R" & year & "(" & year + 2018 & ")" & "/" & month & "/" & day, "ge / m / d")
    
    End If
    
    '1.入力項目の確認
    If day = "" Then
        MsgBox "年月日を入力してください"
    ElseIf IsNull(ListBox1.value) Then
        MsgBox "年号を選択してくだいさい"
    ElseIf money = "" Then
        MsgBox "金額を入力してください"
    
    Else '入力データの登録
    
        
        '3.日付欄に同じ日付がないかを検索
        Set FoundCell = Range(Cells(gyouhajime + 1, 3), Cells(gyousaigo - 3, 3)).CurrentRegion.Find(What:=hiduke)
        If FoundCell Is Nothing Then '同じ日付がなかった場合
            
            '4.最終行の検索　-2行のところから始める
            lastRow = Cells(Rows.count, 3).End(xlUp).row - 2
            Rows(lastRow).Insert Shift:=xlUp
            Rows(lastRow).Interior.ColorIndex = 0
            Rows(lastRow).NumberFormatLocal = "#,##0"
            Rows(lastRow).Font.Bold = False
            
            '5金額が正の値か負の値かの判定
            If money > 0 Then '正の値の場合
                
                '入金に転記する
                With Worksheets("預金推移")
                .Cells(lastRow, 3).value = hiduke '日付なので変えない
                .Cells(lastRow, retu + 1).value = Format(money, "#,###") '入金なので列番号+1
                .Cells(lastRow, retu + 1).HorizontalAlignment = xlRight  '右寄せ
                .Cells(lastRow, retu).value = tekiyou
                .Cells(lastRow, retu).HorizontalAlignment = xlRight  '右寄せ
                End With
                
            Else '負の値だった場合
                '出金に転記する
                With Worksheets("預金推移")
                .Cells(lastRow, 3).value = hiduke '日付なので変えない
                .Cells(lastRow, retu).value = Format(money * -1, "#,###") '出金なので列番号
                .Cells(lastRow, retu).HorizontalAlignment = xlRight   '右寄せ
                .Cells(lastRow, retu + 1).value = tekiyou
                .Cells(lastRow, retu + 1).HorizontalAlignment = xlLeft '左寄せ
                End With
            End If
            
         Else  '入力値があった場合→①同じ日付の数をカウント　②ifで同じ日付のところに入力値があるかカウント分ループさせる
                '①日付の行番号を取得
                gyou = WorksheetFunction.match(hiduke, Range("C1:C10000"), 0)
                hidukekazu = WorksheetFunction.CountIf(Range("c1:c10000"), hiduke)
                '②カウント分ループさせる ifで同じ日付のところに入力値があるか判定
                For i = 1 To hidukekazu
                If Cells(gyou + i - 1, retu) = "" And Cells(gyou + i - 1, retu + 1) = "" Then
                '5金額が正の値か負の値かの判定
                If money > 0 Then '正の値の場合
                '入金に転記する
                With Worksheets("預金推移")
                '日付は処理しない
                .Cells(gyou + i - 1, retu + 1).value = Format(money, "#,###") '同じ日付のところに入れる
                .Cells(gyou + i - 1, retu + 1).HorizontalAlignment = xlRight '右寄せ
                .Cells(gyou + i - 1, retu).value = tekiyou
                .Cells(gyou + i - 1, retu).HorizontalAlignment = xlRight '右寄せ
                End With
                Exit For
                Else '負の値だった場合
                '出金に転記する
                With Worksheets("預金推移")
                '日付は処理しない
                .Cells(gyou + i - 1, retu).value = Format(money * -1, "#,###") '同じ日付のところに入れる
                .Cells(gyou + i - 1, retu).HorizontalAlignment = xlRight   '右寄せ
                .Cells(gyou + i - 1, retu + 1).value = tekiyou
                .Cells(gyou + i - 1, retu + 1).HorizontalAlignment = xlLeft '左寄せ
                End With
                End If
                Exit For
                
                ElseIf i = hidukekazu Then '同じ日付があったが全て入力値が合った場合
                '4.最終行の検索　-2行のところから始める
                lastRow = 0
                lastRow = Cells(Rows.count, 3).End(xlUp).row - 2
                Rows(lastRow).Insert Shift:=xlUp
                Rows(lastRow).Interior.ColorIndex = 0
                Rows(lastRow).NumberFormatLocal = "#,##0"
                
                '5金額が正の値か負の値かの判定
                If money > 0 Then '正の値の場合
                '入金に転記する
                With Worksheets("預金推移")
                .Cells(lastRow, 3).value = hiduke '日付なので変えない
                .Cells(lastRow, retu + 1).value = Format(money, "#,###") '入金なので列番号+1
                .Cells(lastRow, retu + 1).HorizontalAlignment = xlRight   '右寄せ
                .Cells(lastRow, retu).value = tekiyou
                .Cells(lastRow, retu).HorizontalAlignment = xlRight    '右寄せ
                End With
                Exit For
                Else '負の値だった場合
                '出金に転記する
                With Worksheets("預金推移")
                .Cells(lastRow, 3).value = hiduke '日付なので変えない
                .Cells(lastRow, retu).value = Format(money * -1, "#,###") '出金なので列番号
                .Cells(lastRow, retu).HorizontalAlignment = xlRight    '右寄せ
                .Cells(lastRow, retu + 1).value = tekiyou
                .Cells(lastRow, retu + 1).HorizontalAlignment = xlLeft  '右寄せ
                End With
                End If
                Exit For
            End If
            Next
        End If
    End If
    '入力後の処理　①i=0　②フォーカスを年に　③テキストボックスを空欄に
    '①
    i = 0
    '②
    Me!year.SetFocus
    '③
    year.value = ""
    month.value = ""
    day.value = ""
    money.value = ""
    tekiyou = ""
    
    
    gyouhajime = Range("A1:A10000").Find("資金移動始").row
    gyousaigo = Range("A1:A10000").Find("資金移動終").row
    retuhajime = Range("C1:BZ1").Find("資金移動始").Column
    retusaigo = Range("C1:DZ1").Find("資金移動終").Column
    
    retu = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column '入力列番号の取得
    With ActiveWorkbook.Worksheets("預金推移")
    Range(.Cells(gyouhajime + 1, retuhajime), .Cells(gyousaigo - 3, retusaigo)).Sort Key1:=.Columns(3)
    End With

End Sub

Private Sub UserForm_Initialize()
retu = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column '入力列番号の取得
資金移動.Label15.Caption = Cells(7, retu).text
資金移動.Label16.Caption = Cells(6, retu).text

With ListBox1
.AddItem "令和"
.AddItem "平成"
End With
Me.ListBox1.ListIndex = 0

Me.Height = 274
Me.Width = 358 '適切な幅を設定
Me.Zoom = 100

End Sub

'通貨形式の文字列から年度を取り出す関数
Function ExtractYear(inputText As String) As String
    Dim regex As Object
    Dim match As Object
    Set regex = CreateObject("VBScript.RegExp")

    ' 正規表現でカッコ内の西暦を抽出
    regex.Pattern = "\((\d{4})\)" ' (2023) の形式を取得
    regex.Global = False

    If regex.Test(inputText) Then
        Set match = regex.Execute(inputText)
        ExtractYear = match(0).SubMatches(0) ' 抽出した西暦を返す
    Else
        ExtractYear = "" ' 該当なしの場合
    End If
End Function

'文字列「R〇(yyyy)/MM/DD」を、yyyy/MM/DDに変換する関数
Function dateConverter(targetCell As String) As Variant
    Dim rawDate As String
    Dim dateParts() As String
    Dim yearPart As Long, monthPart As Long, dayPart As Long

    On Error GoTo ErrorHandler

    ' フォーマットチェック：括弧とスラッシュがなければエラーへ
    If targetCell = "" Or InStr(targetCell, "(") = 0 Or InStr(targetCell, "/") = 0 Then GoTo ErrorHandler

    ' 括弧内を抽出 → "2019)/10/02"
    rawDate = Mid(targetCell, InStr(targetCell, "(") + 1)
    rawDate = Replace(rawDate, ")", "") ' 括弧閉じ削除 → "2019/10/02"

    ' 3つに分割されない場合もエラーへ
    dateParts = Split(rawDate, "/")
    If UBound(dateParts) < 2 Then GoTo ErrorHandler

    ' 各パーツが全て数値化できるかチェック（Valなら多少の文字混入にも強い）
    yearPart = val(dateParts(0))
    monthPart = val(dateParts(1))
    dayPart = val(dateParts(2))

    ' 0 は無効なので弾く
    If yearPart = 0 Or monthPart = 0 Or dayPart = 0 Then GoTo ErrorHandler

    ' 変換成功
    dateConverter = DateSerial(yearPart, monthPart, dayPart)
    Exit Function

ErrorHandler:
    dateConverter = 0
End Function


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
MsgBox "［閉じる］ボタンを使用してください"
Cancel = True
End If
End Sub





