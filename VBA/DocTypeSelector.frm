VERSION 5.00
Begin VB.UserForm DocTypeSelector 
   Caption         =   "書類タイプを選択"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnOK
      Caption         =   "確定"
      Height          =   360
      Left            =   1980
      TabIndex        =   3
      Top             =   1560
      Width           =   900
   End
   Begin VB.CommandButton btnCancel
      Caption         =   "キャンセル"
      Height          =   360
      Left            =   3000
      TabIndex        =   4
      Top             =   1560
      Width           =   900
   End
   Begin VB.OptionButton optTransactions
      Caption         =   "取引履歴（入出金明細）"
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   870
      Width           =   3100
   End
   Begin VB.OptionButton optBankbook
      Caption         =   "通帳（預金残高）"
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   540
      Width           =   3100
   End
   Begin VB.Label lblInstruction
      Caption         =   "処理したい書類の種類を選択してください。"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3300
   End
End
Attribute VB_Name = "DocTypeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_selectedType As String
Private m_cancelled As Boolean

Public Function ShowDialog(defaultType As String) As String
    If LCase$(defaultType) = "bank_deposit" Then
        optBankbook.Value = True
    Else
        optTransactions.Value = True
    End If
    m_cancelled = True
    Me.Show vbModal
    If m_cancelled Then
        ShowDialog = ""
    Else
        ShowDialog = m_selectedType
    End If
End Function

Private Sub btnCancel_Click()
    m_cancelled = True
    m_selectedType = ""
    Me.Hide
End Sub

Private Sub btnOK_Click()
    If optBankbook.Value Then
        m_selectedType = "bank_deposit"
    ElseIf optTransactions.Value Then
        m_selectedType = "transaction_history"
    Else
        MsgBox "書類タイプを選択してください。", vbExclamation
        Exit Sub
    End If
    m_cancelled = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        btnCancel_Click
    End If
End Sub
