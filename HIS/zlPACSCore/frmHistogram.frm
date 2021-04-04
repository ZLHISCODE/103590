VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Begin VB.Form frmHistogram 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "直方图"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   DrawStyle       =   1  'Dash
   Icon            =   "frmHistogram.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "退出(&Q)"
      Height          =   350
      Left            =   5400
      TabIndex        =   12
      Top             =   4440
      Width           =   1100
   End
   Begin VB.Frame frmMaxAndValue 
      Height          =   975
      Left            =   228
      TabIndex        =   3
      Top             =   4050
      Width           =   4815
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "灰度值："
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "灰度值："
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "最多点数："
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "最少点数："
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3615
      Left            =   0
      OleObjectBlob   =   "frmHistogram.frx":000C
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label lblEnd 
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblStart 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   3840
      Width           =   1695
   End
End
Attribute VB_Name = "frmHistogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub MSChart1_Click()
    Me.Command1.SetFocus
End Sub

