VERSION 5.00
Begin VB.Form frm德阳票据打印 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印条件设置"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frm德阳票据打印.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   5
      Top             =   2325
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2085
      TabIndex        =   4
      Top             =   2325
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   2145
      Width           =   9300
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   2
      Top             =   750
      Width           =   7110
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   0
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1095
      Width           =   2760
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   1
      Left            =   1485
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "次单位"
      Top             =   1508
      Width           =   2760
   End
   Begin VB.Image img 
      Height          =   555
      Left            =   75
      Picture         =   "frm德阳票据打印.frx":020A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "请输入需要打印的住院号范围"
      Height          =   165
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   420
      Width           =   4965
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "开始住院号"
      Height          =   180
      Index           =   1
      Left            =   510
      TabIndex        =   7
      Top             =   1162
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "结束住院号"
      Height          =   180
      Index           =   2
      Left            =   510
      TabIndex        =   6
      Tag             =   "次单位"
      Top             =   1575
      Width           =   900
   End
End
Attribute VB_Name = "frm德阳票据打印"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrInfor As String
Dim mblnOK As Boolean
Dim mblnChange As Boolean
Dim mstr开始住院号 As String
Dim mstr结束住院号 As String
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim i As Long
    Dim strInfor As String
    
    mstr开始住院号 = txtEdit(0).Text
    mstr结束住院号 = txtEdit(1).Text
    
    mblnOK = True
    Unload Me
End Sub



Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub


Public Function ShowCard(ByRef str开始住院号 As String, ByRef str结束住院号 As String) As Boolean

    txtEdit(0).Text = str开始住院号
    txtEdit(1).Text = str结束住院号
    
    Me.Show 1
    str开始住院号 = mstr开始住院号
    str结束住院号 = mstr结束住院号
    ShowCard = mblnOK

End Function

