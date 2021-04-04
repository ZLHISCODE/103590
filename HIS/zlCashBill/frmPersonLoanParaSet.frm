VERSION 5.00
Begin VB.Form frmPersonLoanParaSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSplit 
      Height          =   135
      Left            =   -30
      TabIndex        =   3
      Top             =   2760
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3660
      TabIndex        =   2
      Top             =   3090
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2475
      TabIndex        =   1
      Top             =   3090
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "借款票据打印设置"
      Height          =   360
      Left            =   2880
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2235
      Width           =   1875
   End
End
Attribute VB_Name = "frmPersonLoanParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mblnHavePriv As Boolean, mstrPrivs As String, mblnSelect As Boolean
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/12/19
    '------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    SaveSet = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Function
Private Sub cmdOK_Click()
    If SaveSet = False Then Exit Sub
    mblnSelect = True
    Unload Me
End Sub

Public Function 设置参数(ByVal frmParent As Object, ByVal lngModuel As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对参数进行设置
    '入参:frmParent-调用的窗体
    '     lngModuel-调用的模块号
    '     strPrivs-权限串
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-10 12:15:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModuel: mstrPrivs = strPrivs
    mblnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    mblnSelect = False
    Me.Show vbModal, frmParent
    设置参数 = mblnSelect
End Function



Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_1502"
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub
