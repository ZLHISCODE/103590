VERSION 5.00
Begin VB.Form frmMessage昭通 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "信息提示"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&O)"
      Height          =   405
      Left            =   3135
      TabIndex        =   1
      Top             =   3630
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   7245
   End
End
Attribute VB_Name = "frmMessage昭通"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub showMessage(strMessage As String)
    Text1.Text = strMessage
    Me.Show vbModal
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

