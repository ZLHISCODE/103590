VERSION 5.00
Begin VB.Form frmErrAsk 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "提示"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ControlBox      =   0   'False
   Icon            =   "frmErrAsk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdRetry 
      Caption         =   "重试(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1200
      TabIndex        =   0
      Top             =   1350
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2145
      TabIndex        =   1
      Top             =   1350
      Width           =   900
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   3090
      TabIndex        =   2
      Top             =   1350
      Width           =   900
   End
   Begin VB.TextBox txtHelp 
      Height          =   765
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmErrAsk.frx":0E42
      Top             =   1815
      Width           =   4275
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frmErrAsk.frx":0EA7
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "序号："
      Height          =   180
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblScrip 
      AutoSize        =   -1  'True
      Caption         =   "说明："
      Height          =   180
      Left            =   945
      TabIndex        =   6
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblNote 
      Caption         =   "    可能是其他用户的独占或重新安装了操作系统带来的错误，排除独占使用因素仍不能运行，则需部分重装本系统。"
      Height          =   585
      Left            =   945
      TabIndex        =   5
      Top             =   330
      Width           =   3390
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAsk 
      AutoSize        =   -1  'True
      Caption         =   "再试一次吗"
      Height          =   180
      Left            =   945
      TabIndex        =   4
      Top             =   1020
      Width           =   900
   End
End
Attribute VB_Name = "frmErrAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bytReturn As Byte

Private Sub cmdCancel_Click()
    bytReturn = 0
    Hide
End Sub

Private Sub cmdHelp_Click()
    Height = Height + txtHelp.Height + 100
    cmdHelp.Enabled = False
End Sub

Private Sub cmdRetry_Click()
    bytReturn = 1
    Hide
End Sub

Private Sub Form_Load()
    bytReturn = 1
End Sub

