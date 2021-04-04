VERSION 5.00
Begin VB.Form frmSquareAffirmParaSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "结算凭条打印设置(&2)"
      Height          =   360
      Index           =   1
      Left            =   2250
      TabIndex        =   13
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Frame fraClinicBalance 
      Caption         =   "诊间门诊结帐后结帐凭条打印方式"
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   105
      TabIndex        =   8
      Top             =   2340
      Width           =   4650
      Begin VB.OptionButton optClinicBalancePrint 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   525
         TabIndex        =   9
         Top             =   450
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optClinicBalancePrint 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   3030
         TabIndex        =   11
         Top             =   450
         Width           =   1380
      End
      Begin VB.OptionButton optClinicBalancePrint 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   1620
         TabIndex        =   10
         Top             =   450
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   4860
      TabIndex        =   14
      Top             =   -450
      Width           =   30
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5070
      TabIndex        =   15
      Top             =   255
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5070
      TabIndex        =   16
      Top             =   750
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "票据打印设置(&1)"
      Height          =   360
      Index           =   0
      Left            =   105
      TabIndex        =   12
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Frame fraRecored 
      Caption         =   "诊间记帐审核后记帐票据打印方式"
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   105
      TabIndex        =   4
      Top             =   1305
      Width           =   4650
      Begin VB.OptionButton optRecordPrint 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   1620
         TabIndex        =   6
         Top             =   450
         Width           =   1020
      End
      Begin VB.OptionButton optRecordPrint 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   3030
         TabIndex        =   7
         Top             =   450
         Width           =   1380
      End
      Begin VB.OptionButton optRecordPrint 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   525
         TabIndex        =   5
         Top             =   450
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.Frame fraCharge 
      Caption         =   "诊间消费确定后收费票据打印方式"
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   105
      TabIndex        =   0
      Top             =   270
      Width           =   4650
      Begin VB.OptionButton optChargePrint 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   465
         TabIndex        =   1
         Top             =   375
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optChargePrint 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   3030
         TabIndex        =   3
         Top             =   375
         Width           =   1380
      End
      Begin VB.OptionButton optChargePrint 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   1620
         TabIndex        =   2
         Top             =   375
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmSquareAffirmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String, mblnOk As Boolean
Private Const mlngModul = 1151
Public Function SetPara(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费的相关参数设置入口
    '返回:点击确定,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-11 00:16:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOk = False
    Me.Show 1, frmMain
    SetPara = mblnOk
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function SavePara() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存参数
    '编制:刘兴洪
    '日期:2011-08-10 23:37:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "收费打印方式", IIf(optChargePrint(0).value, 0, IIf(optChargePrint(1).value, 1, 2)), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "审核打印方式", IIf(optRecordPrint(0).value, 0, IIf(optRecordPrint(1).value, 1, 2)), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "门诊结帐凭条打印方式", IIf(optClinicBalancePrint(0).value, 0, IIf(optClinicBalancePrint(1).value, 1, 2)), glngSys, mlngModul, blnHavePrivs
    
    SavePara = True
End Function
 Private Sub cmdOK_Click()
    If SavePara = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数
    '编制:刘兴洪
    '日期:2011-08-10 23:48:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim blnHavePrivs As Boolean, strValue As String
    Dim j As Long
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    i = Val(zlDatabase.GetPara("收费打印方式", glngSys, mlngModul, , Array(optChargePrint(0), optChargePrint(1), optChargePrint(2)), blnHavePrivs))
    i = IIf(i < 0, 0, i): i = IIf(i > 2, 2, i)
    optChargePrint(i).value = True
    
    i = Val(zlDatabase.GetPara("审核打印方式", glngSys, mlngModul, , Array(optRecordPrint(0), optRecordPrint(1), optRecordPrint(2)), blnHavePrivs))
    i = IIf(i < 0, 0, i): i = IIf(i > 2, 2, i)
    optRecordPrint(i).value = True
    
    i = Val(zlDatabase.GetPara("门诊结帐凭条打印方式", glngSys, mlngModul, , Array(optClinicBalancePrint(0), optClinicBalancePrint(1), optClinicBalancePrint(2)), blnHavePrivs))
    i = IIf(i < 0, 0, i): i = IIf(i > 2, 2, i)
    optClinicBalancePrint(i).value = True
End Sub

Private Sub cmdPrintSet_Click(Index As Integer)
    Select Case Index
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_5", Me)
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mstrPrivs = ";" & GetPrivFunc(glngSys, mlngModul) & ";"
    Call InitPara
End Sub

