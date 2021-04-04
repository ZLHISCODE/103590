VERSION 5.00
Begin VB.Form frmDefTableMerge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "合并单元格"
   ClientHeight    =   2160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4725
   Icon            =   "frmDefTableMerge.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -165
      TabIndex        =   5
      Top             =   1530
      Width           =   5265
   End
   Begin VB.TextBox txt 
      Height          =   300
      Left            =   2010
      TabIndex        =   1
      Top             =   1020
      Width           =   2370
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3375
      TabIndex        =   3
      Top             =   1725
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2160
      TabIndex        =   2
      Top             =   1725
      Width           =   1100
   End
   Begin VB.Label Label2 
      Caption         =   "内容(&T)"
      Height          =   225
      Left            =   1140
      TabIndex        =   0
      Top             =   1065
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   225
      Picture         =   "frmDefTableMerge.frx":000C
      Top             =   255
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "请在下面的输入框中输入任意表格的单元格合并后的文字（注意必须输入内容才能合并，若为空可以输入空格）。"
      Height          =   540
      Left            =   1095
      TabIndex        =   4
      Top             =   195
      Width           =   3420
   End
End
Attribute VB_Name = "frmDefTableMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private blnOK As Boolean
Private strText As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txt.Text = "" Then
        MsgBox "您必须输入内容，否则不能合并！", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    blnOK = True
    strText = txt.Text
    Unload Me
End Sub

Private Sub Form_Load()
    blnOK = False
End Sub

Public Function ShowMergeBox(frmParent As Form, Text As String) As Boolean
    txt.Text = Text
    frmDefTableMerge.Show 1, frmParent
    Text = strText
    ShowMergeBox = blnOK
    
End Function

Private Sub txt_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub

