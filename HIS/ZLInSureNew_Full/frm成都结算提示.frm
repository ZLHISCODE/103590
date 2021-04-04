VERSION 5.00
Begin VB.Form frm成都结算提示 
   Caption         =   "中联软件"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4800
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frm提示信息 
      Caption         =   "医保结算信息提示"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3600
         Top             =   360
      End
      Begin VB.Label Lbl信息 
         Caption         =   "提示信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frm成都结算提示"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    Timer1.Enabled = True
    Lbl信息.Caption = g成都结算信息
End Sub

Private Sub Timer1_Timer()
   Unload Me
End Sub
