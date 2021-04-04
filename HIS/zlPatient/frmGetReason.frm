VERSION 5.00
Begin VB.Form frmGetReason 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改原因"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4995
   Icon            =   "frmGetReason.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "确  定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3720
      TabIndex        =   2
      Top             =   1215
      Width           =   1100
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   500
      Width           =   4695
   End
   Begin VB.Label lblReason 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "请填写修改原因:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmGetReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrReason As String

Private Sub cmdOK_Click()
    If Trim(txtReason.Text) <> "" Then
        mstrReason = txtReason.Text
        Unload Me
    End If
End Sub

Public Function ShowMe(ByVal frmParent As Object, ByRef strReason As String) As Boolean
    Me.Show 1, frmParent
    strReason = mstrReason
End Function

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If zlCommFun.ActualLen(txtReason.Text) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
        KeyAscii = 0
    End If
End Sub
