VERSION 5.00
Begin VB.Form frmMediPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "frmMediPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5325
      TabIndex        =   0
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6675
      TabIndex        =   1
      Top             =   3840
      Width           =   1100
   End
End
Attribute VB_Name = "frmMediPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strPrivs As String
Dim mblnSetPara As Boolean      '是否具有参数设置权限

Public Sub ShowMe(ByVal strPrivss As String, ByVal frmParent As Object)
    strPrivs = strPrivss
    Me.Show 1, frmParent
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub
