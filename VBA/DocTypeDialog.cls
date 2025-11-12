VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocTypeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents btnOk As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private optBankbook As MSForms.OptionButton
Private optTransactions As MSForms.OptionButton
Private frm As Object
Private cancelled As Boolean
Private resultValue As String

Public Function ShowDialog(defaultType As String) As String
    BuildForm defaultType
    cancelled = True
    frm.Show
    If cancelled Then
        ShowDialog = ""
    Else
        ShowDialog = resultValue
    End If
    UnloadForm
End Function

Private Sub BuildForm(defaultType As String)
    Set frm = VBA.UserForms.Add()
    With frm
        .Caption = "書類タイプを選択"
        .Width = 360
        .Height = 180
        .StartUpPosition = 1 'CenterOwner
        .ControlBox = False
    End With

    Dim lbl As MSForms.Label
    Set lbl = frm.Controls.Add("Forms.Label.1", "lblInstruction", True)
    lbl.Caption = "処理したい書類の種類を選択してください。"
    lbl.Left = 10
    lbl.Top = 10
    lbl.Width = frm.InsideWidth - 20
    lbl.Height = 24

    Set optBankbook = frm.Controls.Add("Forms.OptionButton.1", "optBankbook", True)
    optBankbook.Caption = "通帳（預金残高）"
    optBankbook.Left = 20
    optBankbook.Top = 40
    optBankbook.Width = frm.InsideWidth - 40

    Set optTransactions = frm.Controls.Add("Forms.OptionButton.1", "optTransactions", True)
    optTransactions.Caption = "取引履歴（入出金明細）"
    optTransactions.Left = optBankbook.Left
    optTransactions.Top = optBankbook.Top + 24
    optTransactions.Width = optBankbook.Width

    If LCase$(defaultType) = "bank_deposit" Then
        optBankbook.Value = True
    Else
        optTransactions.Value = True
    End If

    Set btnOk = frm.Controls.Add("Forms.CommandButton.1", "btnOk", True)
    btnOk.Caption = "確定"
    btnOk.Width = 80
    btnOk.Height = 24
    btnOk.Left = frm.InsideWidth - btnOk.Width * 2 - 20
    btnOk.Top = optTransactions.Top + 32

    Set btnCancel = frm.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    btnCancel.Caption = "キャンセル"
    btnCancel.Width = 80
    btnCancel.Height = 24
    btnCancel.Left = frm.InsideWidth - btnCancel.Width - 10
    btnCancel.Top = btnOk.Top
End Sub

Private Sub btnCancel_Click()
    cancelled = True
    resultValue = ""
    frm.Hide
End Sub

Private Sub btnOk_Click()
    If optBankbook.Value Then
        resultValue = "bank_deposit"
    Else
        resultValue = "transaction_history"
    End If
    cancelled = False
    frm.Hide
End Sub

Private Sub UnloadForm()
    If Not frm Is Nothing Then
        VBA.UserForms.Remove frm
        Set frm = Nothing
    End If
End Sub
