VERSION 5.00
Begin VB.Form frmStuffPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmStuffPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   165
      TabIndex        =   2
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5625
      TabIndex        =   0
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6870
      TabIndex        =   1
      Top             =   4725
      Width           =   1100
   End
End
Attribute VB_Name = "frmStuffPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnActive As Boolean
Private mintTabIndex As Integer
Private mstrPrivs As String
Private mblnHavePriv As Boolean
Private Const mlngModule = 1711

Public Sub ShowMe(ByVal strPrivs As String, ByVal frmMain As Object)
    '----------------------------------------------------------------------------------
    '功能:参数设置入口
    '参数:mstrPrivs -权限串
    '     frmMain-调用父窗口
    '返回:
    '编制:刘兴宏
    '日期:2007/12/24
    '----------------------------------------------------------------------------------
    mstrPrivs = strPrivs
    Me.Show 1, frmMain
End Sub

Private Sub cmdCancel_Click()
    gblnIncomeItem = False
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdOk_Click()
    Dim intSave As Integer
    Dim strReg As String
    
    If SaveSet = False Then Exit Sub
    gblnIncomeItem = True
    Unload Me
End Sub
Private Function SaveSet() As Boolean
    
End Function
Private Sub Form_Activate()
    If Not mblnActive Then Unload Me: Exit Sub
    
End Sub

