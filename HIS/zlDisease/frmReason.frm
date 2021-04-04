VERSION 5.00
Begin VB.Form frmReason 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "不填写的理由"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6015
   Icon            =   "frmReason.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6015
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1320
      Width           =   950
   End
   Begin VB.TextBox txtReason 
      Height          =   735
      Left            =   240
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblTitle 
      Caption         =   "请输入不填写报告卡的理由："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrReason As String

Public Function ShowMe(ByRef frmObj As Object, Optional ByVal strDefault As String) As String
    txtReason.Text = strDefault
    Me.Show 1
    
    ShowMe = mstrReason
End Function

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Trim(txtReason.Text) = "" Then
        MsgBox "理由不能够为空，请输入。", vbInformation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    mstrReason = Trim(txtReason.Text)
End Sub

Private Sub txtReason_GotFocus()
    txtReason.SelStart = 0: txtReason.SelLength = 1000
    Call gobjComlib.ZLCommFun.OpenIme(True)
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call gobjComlib.ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
