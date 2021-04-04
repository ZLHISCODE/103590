VERSION 5.00
Begin VB.Form frmRequestPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmRequestPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk申领核查 
      Caption         =   "申领需要核查后才能移库"
      Height          =   375
      Left            =   1020
      TabIndex        =   10
      Top             =   1440
      Width           =   3105
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "单据打印设置(&S)"
      Height          =   350
      Left            =   1020
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2385
      Width           =   2400
   End
   Begin VB.Frame fra 
      Height          =   120
      Index           =   1
      Left            =   -30
      TabIndex        =   9
      Top             =   2910
      Width           =   5790
   End
   Begin VB.Frame fra 
      Height          =   120
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   645
      Width           =   5580
   End
   Begin VB.ComboBox Cbo指定单位 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1050
      Width           =   2415
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   30
      TabIndex        =   6
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2940
      TabIndex        =   4
      Top             =   3150
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4275
      TabIndex        =   5
      Top             =   3150
      Width           =   1100
   End
   Begin VB.CheckBox chkSavePrint 
      Caption         =   "存盘打印"
      Height          =   375
      Left            =   1020
      TabIndex        =   2
      Top             =   1875
      Width           =   3105
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   45
      Picture         =   "frmRequestPara.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "如果选择存盘打印，则在单据中，单据存盘后自动打印，否则不打印。"
      Height          =   615
      Left            =   630
      TabIndex        =   8
      Top             =   225
      Width           =   3180
   End
   Begin VB.Label lbl材料单位 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "材料单位"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   1110
      Width           =   720
   End
End
Attribute VB_Name = "frmRequestPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Private mstrPrivs As String
Private mlngModule As Long
Private mblnHavePriv As Boolean '是否有参数设置权限

Private Sub Cbo指定单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chkSavePrint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/12/24
    '------------------------------------------------------------------------------------------
 
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
    Call zlDatabase.SetPara("存盘打印", IIf(chkSavePrint.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("卫材单位", Cbo指定单位.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("申领需要核查后才能移库", IIf(chk申领核查.Value = 1, 1, 0), glngSys, mlngModule)
    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function
Private Sub cmdOk_Click()
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub
Private Sub initPara()
    '-----------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数设置
    '返回:
    '编制:刘兴宏
    '修改:2007/12/24
    '-----------------------------------------------------------------------------------------------------------------------------------------
    Dim strReg As String
    mblnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    chkSavePrint.Value = IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule, "0", Array(chkSavePrint), mblnHavePriv)) = 1, 1, 0)
    Me.CmdPrintSet.Enabled = InStr(1, mstrPrivs, ";单据打印;") <> 0
    chk申领核查.Value = IIf(Val(zlDatabase.GetPara("申领需要核查后才能移库", glngSys, mlngModule, "0", Array(chk申领核查), mblnHavePriv)) = 1, 1, 0)
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0", Array(Cbo指定单位, lbl材料单位), mblnHavePriv))
    With Cbo指定单位
        .Clear
        .AddItem "散装单位"
        .AddItem "包装单位"
        .ListIndex = Val(strReg)
    End With
End Sub
Public Sub 设置参数(ByVal lngModule As Long, frmMain As Form, Optional ByVal strFunction As String = "", Optional strPrivs As String = "")
    '-----------------------------------------------------------------------------------------------------------------------------------------
    '功能:进入参数设置界面
    '返回:
    '编制:刘兴宏
    '修改:2007/12/24
    '-----------------------------------------------------------------------------------------------------------------------------------------
    mstrFunction = strFunction: mlngModule = lngModule:    mstrPrivs = IIf(strPrivs = "", gstrPrivs, strPrivs)
    Call initPara
    frmRequestPara.Show vbModal, frmMain
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_" & glngModul
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

