VERSION 5.00
Begin VB.Form frmErrAsk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "提示"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmErrAsk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.FileListBox File1 
      Height          =   2610
      Left            =   1980
      TabIndex        =   4
      Top             =   3000
      Width           =   3330
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   870
      TabIndex        =   2
      Top             =   30
      Width           =   3645
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "    可能是其他用户的独占或重新安装了操作系统带来的错误，排除独占使用因素仍不能运行，则需部分重装本系统。"
         Height          =   1050
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   3420
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2565
      TabIndex        =   1
      Top             =   1500
      Width           =   1100
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "重试(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1350
      TabIndex        =   0
      Top             =   1500
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmErrAsk.frx":000C
      Top             =   210
      Width           =   480
   End
End
Attribute VB_Name = "frmErrAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytReturn As Byte

Public Function ShowForm(ByVal strNumber As String, ByVal strNote As String) As Byte
    lblNote = strNote
    Me.Show 1
    ShowForm = mbytReturn
End Function

Private Sub cmdCancel_Click()
    mbytReturn = 0
    Unload Me
End Sub

Private Sub cmdRetry_Click()
    mbytReturn = 1
    Unload Me
End Sub
