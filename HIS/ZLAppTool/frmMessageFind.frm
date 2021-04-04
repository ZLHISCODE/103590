VERSION 5.00
Begin VB.Form frmMessageFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "frmMessageFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkBegin 
      Caption         =   "从头开始(&B)"
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   780
      Width           =   1575
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "区分大小写(&A)"
      Height          =   315
      Left            =   210
      TabIndex        =   3
      Top             =   780
      Width           =   1575
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   1170
      TabIndex        =   1
      Top             =   240
      Width           =   2595
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Left            =   4020
      TabIndex        =   4
      Top             =   180
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4020
      TabIndex        =   5
      Top             =   630
      Width           =   1100
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查找内容(&N)"
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   300
      Width           =   990
   End
End
Attribute VB_Name = "frmMessageFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmMain As frmMessageEdit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    With frmMain
        .mblnCase = chkCase.Value = 1
        .mblnBegin = chkBegin.Value = 1
        .mstrFind = txtFind.Text
        .FindText
    End With
    chkBegin.Value = 0
    frmMain.mblnBegin = False
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = txtFind.Text <> ""
End Sub
