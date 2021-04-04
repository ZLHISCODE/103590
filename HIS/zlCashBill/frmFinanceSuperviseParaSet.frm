VERSION 5.00
Begin VB.Form frmFinanceSuperviseParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinanceSuperviseParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "备用金领用单打印设置(&B)"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   8
      Top             =   2130
      Width           =   2700
   End
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "收款收据打印设置(&S)"
      Height          =   375
      Index           =   0
      Left            =   195
      TabIndex        =   7
      Top             =   2130
      Width           =   2250
   End
   Begin VB.Frame fraSplit 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   0
      Left            =   -855
      TabIndex        =   3
      Top             =   975
      Width           =   9930
   End
   Begin VB.Frame fraSplit 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   1
      Left            =   -120
      TabIndex        =   2
      Top             =   3015
      Width           =   9525
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4455
      TabIndex        =   1
      Top             =   3375
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3270
      TabIndex        =   0
      Top             =   3375
      Width           =   1100
   End
   Begin VB.TextBox txtDrawMoney 
      Height          =   330
      Left            =   2235
      TabIndex        =   6
      Top             =   1500
      Width           =   1995
   End
   Begin VB.Label lblDrawMoney 
      AutoSize        =   -1  'True
      Caption         =   "备用金缺省领用金额                     元"
      Height          =   210
      Left            =   225
      TabIndex        =   5
      Top             =   1545
      Width           =   4305
   End
   Begin VB.Image imgNotes 
      Height          =   720
      Left            =   195
      Picture         =   "frmFinanceSuperviseParaSet.frx":06EA
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblTittle 
      Caption         =   "   备用金缺省领用金额表示收费员在上岗时缺省领用的备用金."
      Height          =   600
      Left            =   1080
      TabIndex        =   4
      Top             =   435
      Width           =   4575
   End
End
Attribute VB_Name = "frmFinanceSuperviseParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As String, mstrPrivs As String, mblnOK As Boolean
Public Function ShowMe(ByVal frmMain As Form, _
    ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置相关参数的入口
    '返回:参数设置成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-12 14:33:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOK = False: mlngModule = lngModule: mstrPrivs = strPrivs
    Me.Show 1, frmMain
    ShowMe = mblnOK
End Function
Private Sub LoadPara()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载参数
    '编制:刘兴洪
    '日期:2013-09-12 15:26:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtDrawMoney.Text = Val(zlDatabase.GetPara("缺省领用备用金额", glngSys, mlngModule, 1000, Array(txtDrawMoney, lblDrawMoney), InStr(1, mstrPrivs, ";参数设置;") > 0))
End Sub
Private Sub SavePara()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存参数
    '编制:刘兴洪
    '日期:2013-09-12 15:28:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call zlDatabase.SetPara("缺省领用备用金额", Val(txtDrawMoney.Text), glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Call SavePara
    Unload Me
    mblnOK = True
End Sub
Private Sub cmdPrintSetup_Click(Index As Integer)
    If Index = 0 Then
        Call ReportPrintSet(gcnOracle, glngSys, "zl" & Int(glngSys / 100) & "_BILL_1500", Me)
    Else
        Call ReportPrintSet(gcnOracle, glngSys, "zl" & Int(glngSys / 100) & "_BILL_1500_1", Me)
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Call LoadPara
End Sub

