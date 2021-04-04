VERSION 5.00
Begin VB.Form frmPayExitParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frmPayExitParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6000
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra设备定义 
      Caption         =   " 智能卡及其他设备定义 "
      Height          =   1000
      Left            =   3360
      TabIndex        =   16
      Top             =   3720
      Width           =   2535
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame fra 
      Caption         =   " 其他控制 "
      Height          =   1000
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   3135
      Begin VB.CheckBox chkDetailPage 
         Caption         =   "保持上一次窗体关闭时的页签"
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   2745
      End
      Begin VB.CheckBox chkSendByNo 
         Caption         =   "按单据号发料"
         Height          =   420
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2130
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " 业务类型 "
      Height          =   1272
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   3090
      Begin VB.ComboBox cbo收费处方 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   2280
      End
      Begin VB.CheckBox chk业务 
         Caption         =   "收费单(&S)"
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   1150
      End
      Begin VB.CheckBox chk业务 
         Caption         =   "记帐单(&J)"
         Height          =   285
         Index           =   1
         Left            =   1850
         TabIndex        =   1
         Top             =   240
         Width           =   1150
      End
      Begin VB.CheckBox chk业务 
         Caption         =   "记帐表(&B)"
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   1150
      End
      Begin VB.Label lbl收费处方 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "收费处方"
         Height          =   420
         Left            =   120
         TabIndex        =   20
         Top             =   825
         Width           =   465
      End
      Begin VB.Label lbl单据类型 
         Caption         =   "单据类型"
         Height          =   420
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   4920
      Width           =   8775
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   705
      Width           =   8775
   End
   Begin VB.Frame fra 
      Caption         =   " 打印及票据设置 "
      Height          =   1305
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   5775
      Begin VB.OptionButton opt打印方式 
         Caption         =   "不打印(&N)"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton opt打印方式 
         Caption         =   "自动打印(&A)"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton opt打印方式 
         Caption         =   "提示打印(&M)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "票据打印设置"
         Height          =   360
         Left            =   3120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   690
         Width           =   1875
      End
      Begin VB.ComboBox cbo票据设置 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2070
      End
      Begin VB.Label lbl票据 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "票据(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   " 缺省单位 "
      Height          =   1270
      Index           =   0
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Width           =   2364
      Begin VB.OptionButton opt单位 
         Caption         =   "包装单位(&2)"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton opt单位 
         Caption         =   "散装单位(&1)"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   8
      Top             =   5160
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4680
      TabIndex        =   7
      Top             =   5175
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   6
      Top             =   5175
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmPayExitParaSet.frx":030A
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "根据下面选项目,设置相关的打印、发料单位和相关票据的设置"
      Height          =   390
      Index           =   0
      Left            =   735
      TabIndex        =   9
      Top             =   390
      Width           =   5205
   End
End
Attribute VB_Name = "frmPayExitParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mblnExit As Boolean
Private mlngModule As Long
Private mstrPrivs As String
Private mblnHavePriv As Boolean

Private Sub cbo票据设置_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chk打印_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub



Private Sub chk单位_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab

End Sub



 
 


Private Sub chk业务_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub
Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1723)
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
    Dim str业务类型 As String
    Dim n As Integer
    
    str业务类型 = IIf(chk业务(0).Value = 1, "24", "0")
    str业务类型 = str业务类型 & IIf(chk业务(1).Value = 1, ",25", ",0")
    str业务类型 = str业务类型 & IIf(chk业务(2).Value = 1, ",26", ",0")
    
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
   
    Call zlDatabase.SetPara("发料打印提醒方式", IIf(opt打印方式(0).Value = True, 0, IIf(opt打印方式(1).Value = True, 1, 2)), glngSys, mlngModule)
    Call zlDatabase.SetPara("查询业务类型", str业务类型, glngSys, mlngModule)
    Call zlDatabase.SetPara("卫材单位", IIf(opt单位(1).Value = True, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("按单据号发料", chkSendByNo.Value, glngSys, mlngModule)
    Call zlDatabase.SetPara("收费处方显示方式", cbo收费处方.ListIndex, glngSys, mlngModule)
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\卫材发放管理", "保持上一次窗体关闭时的页签", Me.chkDetailPage.Value)

    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdOk_Click()
    If SaveSet = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    
    If cbo票据设置.ListIndex < 0 Then
        ShowMsgBox "请设置好票据!"
        cbo票据设置.SetFocus
    End If
    Select Case cbo票据设置.ListIndex
    Case 0
        '单据打印
        strBill = "ZL1_BILL_1723"
    Case 1
        '清单打印
        strBill = "ZL1_BILL_1723_1"
    Case 2
        '处方退料通知单
        strBill = "ZL1_BILL_1723_2"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim i As Long
    Dim strArr As Variant
    Dim str病区发料 As String
    Dim BlnSelect As Boolean
    Dim n As Integer
    Dim int收费处方 As Integer
    
    mblnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    
    With cbo收费处方
        .Clear
        .AddItem "1-显示所有的处方"
        .AddItem "2-仅显示已收费处方"
        .AddItem "3-仅显示未收费处方"
        .ListIndex = 0
    End With
    
    With cbo票据设置
        .Clear
        .AddItem "1-卫材处方单"
        .AddItem "2-打印已发料清单"
        .AddItem "3-退料通知单打印"
        .ListIndex = 0
    End With
  
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0", Array(opt单位(0), opt单位(1)), mblnHavePriv))
    If Val(strReg) >= 0 And Val(strReg) <= 1 Then
        opt单位(Val(strReg)).Value = True
    Else
        opt单位(0).Value = True
    End If
      
    strReg = Trim(zlDatabase.GetPara("发料打印提醒方式", glngSys, mlngModule, "0", Array(opt打印方式(0), opt打印方式(1), opt打印方式(2)), mblnHavePriv))
    
    If Val(strReg) >= 0 And Val(strReg) <= 2 Then
        opt打印方式(Val(strReg)).Value = True
    Else
        opt打印方式(0).Value = True
    End If
 
    strReg = Trim(zlDatabase.GetPara("查询业务类型", glngSys, mlngModule, "", Array(lbl单据类型, chk业务(0), chk业务(1), chk业务(2), Frame3), mblnHavePriv))
    If strReg = "" Then strReg = "24,25,26"
    strArr = Split(strReg & "," & "," & ",", ",")
    For i = 0 To UBound(strArr)
        If i > 2 Then Exit For
        chk业务(i).Value = IIf(Val(strArr(i)) > 0, 1, 0)
    Next
    
    chkSendByNo.Value = IIf(Val(zlDatabase.GetPara("按单据号发料", glngSys, mlngModule, , Array(chkSendByNo), mblnHavePriv)) = 1, 1, 0)
    
    int收费处方 = Val(zlDatabase.GetPara("收费处方显示方式", glngSys, mlngModule, 0, Array(lbl收费处方, cbo收费处方), mblnHavePriv))
    If int收费处方 >= 0 And int收费处方 <= 2 Then
        cbo收费处方.ListIndex = int收费处方
    Else
        cbo收费处方.ListIndex = 0
    End If
    
    '注册表参数
    Me.chkDetailPage.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "卫材发放管理", "保持上一次窗体关闭时的页签", 0))
End Sub
 
Public Function ShowSetPara(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:设置参数入口
    '参数:
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴宏
    '修改:2007/12/24
    '-----------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    '本地参数设置
     Me.Show 1, frmMain
    ShowSetPara = mblnOk
End Function
