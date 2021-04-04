VERSION 5.00
Begin VB.Form frmEMCStationSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   Icon            =   "frmEMCStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk上次诊断 
      Caption         =   "按科室提取病人上次诊断"
      Height          =   240
      Left            =   4485
      TabIndex        =   51
      Top             =   1270
      Width           =   2300
   End
   Begin VB.CheckBox chk缺省药房 
      Caption         =   "下达医嘱时强制缺省药房"
      Height          =   240
      Left            =   4485
      TabIndex        =   50
      Top             =   1595
      Width           =   2580
   End
   Begin VB.Frame fra急诊诊疗单打印 
      Caption         =   "医嘱发送后,诊疗单据"
      Height          =   1875
      Index           =   0
      Left            =   7320
      TabIndex        =   47
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton opt急诊诊疗单打印 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   55
         Top             =   1320
         Width           =   1080
      End
      Begin VB.OptionButton opt急诊诊疗单打印 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   870
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton opt急诊诊疗单打印 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   420
         Width           =   1560
      End
   End
   Begin VB.Frame fra急诊指引单打印 
      Caption         =   "医嘱发送后,指引单"
      Height          =   1635
      Left            =   7320
      TabIndex        =   43
      Top             =   2280
      Width           =   2295
      Begin VB.OptionButton opt急诊指引单打印 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   1200
      End
      Begin VB.OptionButton opt急诊指引单打印 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   810
         Width           =   1560
      End
      Begin VB.OptionButton opt急诊指引单打印 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   420
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.CheckBox chkYYBR 
      Caption         =   "候诊列表中显示预约病人"
      Height          =   240
      Left            =   150
      TabIndex        =   42
      Top             =   2640
      Width           =   2340
   End
   Begin VB.CheckBox chkCanPay 
      Caption         =   "诊间支付允许使用预交款"
      Height          =   250
      Left            =   4485
      TabIndex        =   41
      Top             =   600
      Width           =   2310
   End
   Begin VB.CheckBox chkAutoClose 
      Caption         =   "医嘱发送后自动关闭窗口"
      Height          =   195
      Left            =   4485
      TabIndex        =   40
      Top             =   1920
      Width           =   2745
   End
   Begin VB.CheckBox chkAutoFinish 
      Caption         =   "接诊病人时自动将上一个病人处理为完成就诊或需回诊"
      Height          =   195
      Left            =   150
      TabIndex        =   37
      Top             =   3600
      Width           =   4665
   End
   Begin VB.Frame fraEPR 
      Caption         =   "提醒设置"
      Height          =   1410
      Left            =   135
      TabIndex        =   24
      Top             =   4000
      Width           =   9480
      Begin VB.CheckBox chk危急值 
         Caption         =   "危急值弹窗提醒"
         Height          =   240
         Left            =   7725
         TabIndex        =   52
         Top             =   277
         Width           =   1620
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "输血反应"
         Height          =   195
         Index           =   6
         Left            =   7725
         TabIndex        =   49
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "用血审核"
         Height          =   195
         Index           =   5
         Left            =   6630
         TabIndex        =   48
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "备血完成"
         Height          =   195
         Index           =   4
         Left            =   5535
         TabIndex        =   36
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "启用语音提示"
         Height          =   195
         Left            =   4350
         TabIndex        =   39
         Top             =   300
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "语音设置(&S)"
         Height          =   350
         Left            =   5850
         TabIndex        =   38
         Top             =   240
         Width           =   1365
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "传染病"
         Height          =   195
         Index           =   3
         Left            =   4605
         TabIndex        =   35
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "处方审查"
         Height          =   195
         Index           =   2
         Left            =   3500
         TabIndex        =   34
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "医嘱安排"
         Height          =   195
         Index           =   1
         Left            =   2330
         TabIndex        =   33
         Top             =   975
         Width           =   1035
      End
      Begin VB.Frame fraNotifyEPRDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   690
         TabIndex        =   30
         Top             =   840
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPR 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   675
         TabIndex        =   27
         Top             =   470
         Width           =   300
      End
      Begin VB.TextBox txtNotifyEPR 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   690
         MaxLength       =   3
         TabIndex        =   26
         Text            =   "10"
         Top             =   270
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "危急值"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   25
         Top             =   975
         Width           =   1035
      End
      Begin VB.CheckBox chkNotifyEPR 
         Caption         =   "每    分钟自动刷新提醒区域中的内容"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   300
         Width           =   3450
      End
      Begin VB.TextBox txtNotifyEPRDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   705
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "1"
         Top             =   630
         Width           =   300
      End
      Begin VB.Label lblNotifyEPRDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "将    天内产生的消息显示在提醒区域"
         Height          =   180
         Left            =   480
         TabIndex        =   32
         Top             =   660
         Width           =   3060
      End
      Begin VB.Label lblArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提醒内容:"
         Height          =   180
         Left            =   465
         TabIndex        =   29
         Top             =   975
         Width           =   810
      End
   End
   Begin VB.Frame fraReceive 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   105
      TabIndex        =   20
      Top             =   3200
      Width           =   4560
      Begin VB.OptionButton optAdd 
         Caption         =   "新增医嘱"
         Enabled         =   0   'False
         Height          =   260
         Index           =   0
         Left            =   1290
         TabIndex        =   23
         Top             =   60
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.CheckBox chkAutoAdd 
         Caption         =   "接诊病人后"
         Height          =   195
         Left            =   45
         TabIndex        =   22
         Top             =   90
         Width           =   2055
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "新增病历"
         Enabled         =   0   'False
         Height          =   260
         Index           =   1
         Left            =   2520
         TabIndex        =   21
         Top             =   60
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdPBPSet 
      Caption         =   "支付票据打印设置"
      Height          =   300
      Left            =   4485
      TabIndex        =   19
      Top             =   210
      Width           =   1620
   End
   Begin VB.CheckBox chkStaKB 
      Caption         =   "启用屏幕键盘"
      Height          =   250
      Left            =   4485
      TabIndex        =   18
      Top             =   935
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   1410
      TabIndex        =   17
      Top             =   2490
      Width           =   465
   End
   Begin VB.TextBox txtQueuePatis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   1365
      MaxLength       =   3
      TabIndex        =   16
      Text            =   "3"
      ToolTipText     =   "表示本次医生最多能呼叫多少个病人来就诊,超过后，就不能再次呼叫;此参数需要配合分诊台模块的排队叫号模式为医生主动呼叫有效"
      Top             =   2280
      Width           =   465
   End
   Begin VB.CheckBox chk自动接诊 
      Caption         =   "查找到候诊病人之后自动接诊"
      Height          =   255
      Left            =   150
      TabIndex        =   10
      Top             =   2950
      Width           =   3135
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   350
      Left            =   135
      TabIndex        =   11
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Caption         =   " 就诊参数 "
      Height          =   2055
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4155
      Begin VB.CommandButton cmdYS 
         Caption         =   "…"
         Height          =   255
         Left            =   3645
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   1635
         Width           =   255
      End
      Begin VB.TextBox txt接诊医生 
         Height          =   300
         Left            =   1020
         TabIndex        =   8
         Top             =   1605
         Width           =   2910
      End
      Begin VB.ComboBox cbo科室 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   2910
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   255
         Left            =   3645
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   690
         Width           =   255
      End
      Begin VB.ComboBox cbo范围 
         ForeColor       =   &H80000012&
         Height          =   300
         ItemData        =   "frmEMCStationSetup.frx":000C
         Left            =   1020
         List            =   "frmEMCStationSetup.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "接诊的病人范围"
         Top             =   1005
         Width           =   2910
      End
      Begin VB.TextBox txt诊室 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1020
         MaxLength       =   20
         TabIndex        =   3
         Top             =   660
         Width           =   2910
      End
      Begin VB.Label lblEditDept 
         AutoSize        =   -1  'True
         Caption         =   "接诊科室"
         Height          =   180
         Left            =   255
         TabIndex        =   0
         Top             =   360
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   45
         X2              =   4090
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label lbl医生 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "接诊医生"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   1665
         Width           =   720
      End
      Begin VB.Label lbl诊室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医生诊室"
         Height          =   180
         Left            =   255
         TabIndex        =   2
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lbl范围 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "接诊范围"
         Height          =   180
         Left            =   225
         TabIndex        =   5
         Top             =   1065
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   45
         X2              =   4090
         Y1              =   1460
         Y2              =   1460
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8565
      TabIndex        =   13
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   12
      Top             =   5760
      Width           =   1100
   End
   Begin VB.Label lblQueuePatis 
      AutoSize        =   -1  'True
      Caption         =   "医生最多能呼叫      人"
      Height          =   180
      Left            =   135
      TabIndex        =   15
      ToolTipText     =   "表示本次医生最多能呼叫多少个病人来就诊,超过后，就不能再次呼叫;此参数需要配合分诊台模块的排队叫号模式为医生主动呼叫有效"
      Top             =   2300
      Width           =   1980
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   -240
      X2              =   10455
      Y1              =   5535
      Y2              =   5535
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   10320
      Y1              =   5565
      Y2              =   5565
   End
End
Attribute VB_Name = "frmEMCStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Private mstrLike As String
Private mbln诊间支付 As Boolean

Private Enum Enum_chkWarn
    chkD危急值 = 0
    chkD医嘱安排 = 1
    chkD处方审查 = 2
    chkD传染病 = 3
    chkD备血完成 = 4
    chkD用血审核 = 5
    chkD输血反应 = 6
End Enum

Private Sub chkAutoAdd_Click()
    If chkAutoAdd.Value = 1 Then
        optAdd(0).Enabled = True
        optAdd(1).Enabled = True
    Else
        optAdd(0).Enabled = False
        optAdd(1).Enabled = False
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call ZLCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdOK_Click()
    Dim blnHavePara As Boolean  '是否有参数设置权限
    Dim i As Integer
    Dim strTmp As String
    
    
    If txt诊室.Text = "" Then
        MsgBox "请设置医生的诊室。", vbInformation, gstrSysName
        txt诊室.SetFocus: Exit Sub
    End If
    If txt接诊医生.Text = "" Then
        MsgBox "请接诊医生。", vbInformation, gstrSysName
        txt接诊医生.SetFocus: Exit Sub
    End If
    If cbo科室.ListIndex < 0 Then
        MsgBox "接诊科室必须选择,请检查", vbInformation + vbOKOnly, gstrSysName
        cbo科室.SetFocus
        Exit Sub
    End If
    
    If chkNotifyEPR.Value = 1 And Val(txtNotifyEPR.Text) = 0 Then
        If txtNotifyEPR.Text = "" Then
            MsgBox "请设置消息提醒的自动刷新间隔。", vbInformation, gstrSysName
        Else
            MsgBox "消息提醒的自动刷新间隔至少应为1分钟。", vbInformation, gstrSysName
        End If
        txtNotifyEPR.SetFocus: Exit Sub
    End If
    
    If Val(txtNotifyEPRDay.Text) = 0 Then
        If txtNotifyEPRDay.Text = "" Then
            MsgBox "请设置要提醒消息的完成天数。", vbInformation, gstrSysName
        Else
            MsgBox "要提醒的消息完成天数至少应为1天。", vbInformation, gstrSysName
        End If
        txtNotifyEPRDay.SetFocus: Exit Sub
    End If
        
    blnHavePara = InStr(1, ";" & mstrPrivs & ";", ";参数设置;") > 0
    
    Call zlDatabase.SetPara("本地诊室", Me.txt诊室.Text, glngSys, p急诊医生站, blnHavePara)
    Call zlDatabase.SetPara("接诊范围", Me.cbo范围.ItemData(Me.cbo范围.ListIndex), glngSys, p急诊医生站, blnHavePara)
    Call zlDatabase.SetPara("接诊医生", Me.txt接诊医生.Text, glngSys, p急诊医生站, blnHavePara)
    
    '刘兴洪:应用于排队叫号的呼叫人次:需要配合分诊台模块的排队叫号模式为１并且有排队呼叫对象=2时有效
    If txtQueuePatis.Enabled Then
        Call zlDatabase.SetPara("医生就诊人数", Val(Me.txtQueuePatis.Text), glngSys, p急诊医生站, blnHavePara)
    End If
    '接诊科室
    Call zlDatabase.SetPara("接诊科室", cbo科室.ItemData(cbo科室.ListIndex), glngSys, p急诊医生站, blnHavePara)
    
    '发送完成后关闭医嘱窗体
    Call zlDatabase.SetPara("发送完成后关闭医嘱窗体", chkAutoClose.Value, glngSys, p门诊医嘱下达, blnHavePara)
    
    '找到病人后自动接诊
    Call zlDatabase.SetPara("找到病人后自动接诊", chk自动接诊.Value, glngSys, p急诊医生站, blnHavePara)
    
    '接诊后自动进行
    If optAdd(1).Value And optAdd(1).Enabled Then
        Call zlDatabase.SetPara("接诊后自动进行", 2, glngSys, p急诊医生站, blnHavePara)
    Else
        Call zlDatabase.SetPara("接诊后自动进行", chkAutoAdd.Value, glngSys, p急诊医生站, blnHavePara)
    End If

    '启用屏幕键盘
    Call zlDatabase.SetPara("启用屏幕键盘", chkStaKB.Value, glngSys, p急诊医生站, blnHavePara)
    
    Call zlDatabase.SetPara("自动刷新病历审阅间隔", IIf(chkNotifyEPR.Value = 1, Val(txtNotifyEPR.Text), ""), glngSys, p急诊医生站, blnHavePara)
    Call zlDatabase.SetPara("自动刷新病历审阅天数", Val(txtNotifyEPRDay.Text), glngSys, p急诊医生站, blnHavePara)
    strTmp = ""
    For i = chkD危急值 To chkD输血反应
        strTmp = strTmp & chkWarn(i).Value
    Next
    Call zlDatabase.SetPara("自动刷新内容", strTmp, glngSys, p急诊医生站, blnHavePara)
    Call zlDatabase.SetPara("接诊时自动处理完成就诊", chkAutoFinish.Value, glngSys, p急诊医生站, blnHavePara)
    Call zlDatabase.SetPara("启用语音提示", chkSound.Value, glngSys, p急诊医生站, blnHavePara)
    
    Call zlDatabase.SetPara("诊间支付允许使用预交款", chkCanPay.Value, glngSys, p门诊医嘱下达, blnHavePara)
    
    Call zlDatabase.SetPara("显示预约病人", chkYYBR.Value, glngSys, p急诊医生站, blnHavePara)
    
    Call zlDatabase.SetPara("门诊医嘱下达强制缺省药房", chk缺省药房.Value, glngSys, p门诊医嘱下达, blnHavePara)
    
    Call zlDatabase.SetPara("急诊危急值弹窗提醒", chk危急值.Value, glngSys, p急诊医生站, blnHavePara)
    
    Call zlDatabase.SetPara("上次诊断按科室提取", chk上次诊断.Value, glngSys, p门诊医嘱下达, blnHavePara)
    
    '医生站是否打印诊疗单据
    Call zlDatabase.SetPara("门诊发送单据打印", IIf(opt急诊诊疗单打印(0).Value = True, 0, IIf(opt急诊诊疗单打印(1).Value = True, 1, 2)), glngSys, p门诊医嘱下达, blnHavePara)
    '医生站是否打印指引单
    Call zlDatabase.SetPara("指引单打印方式", IIf(opt急诊指引单打印(0).Value = True, 0, IIf(opt急诊指引单打印(1).Value = True, 1, 2)), glngSys, p门诊医嘱下达, blnHavePara)

    gblnOK = True
    Unload Me
End Sub

Private Sub chkNotifyEPR_Click()
    txtNotifyEPR.Enabled = chkNotifyEPR.Value = 1
    If Visible And txtNotifyEPR.Enabled Then txtNotifyEPR.SetFocus
End Sub


Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 0)
End Sub

Private Sub txtNotifyEPR_GotFocus()
    Call zlcontrol.TxtSelAll(txtNotifyEPR)
End Sub

Private Sub txtNotifyEPR_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPRDay_GotFocus()
    Call zlcontrol.TxtSelAll(txtNotifyEPRDay)
End Sub

Private Sub txtNotifyEPRDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cmdPBPSet_Click()
    If InitObjPublicExpense Then
        Call gobjPublicExpense.zlCliniqueRoomPayPrintSet(Me)
    End If
End Sub

Private Sub cmdSel_Click()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim vRect As RECT, blnCancel As Boolean
    
    
    If txt诊室.Tag <> txt诊室 Then Exit Sub '由txt诊室的Validate事件处理
    On Error GoTo errH
    If gbln挂号按排 Then
        strSql = "Select Distinct a.Id, a.名称, a.简码" & vbNewLine & _
            " From 门诊诊室 A, 门诊诊室适用科室 B, 部门人员 C, 上机人员表 D,临床部门 F" & vbNewLine & _
            " Where a.Id = b.诊室id And b.科室id = c.部门id And c.人员id = d.人员id" & vbNewLine & _
            "       And d.用户名 = User And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null) And C.部门ID = F.部门ID And F.工作性质 = '20'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Else
        Set rsTmp = GetRs挂号安排诊室列表(1, "", 1, 0, "", p急诊医生站)
    End If
    
    '如果没有查找到数据，则读取出所有的门诊诊室供选择
    If rsTmp.EOF Then
        strSql = "Select a.Id, a.名称, a.简码 From 门诊诊室 A Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    End If
    
    vRect = zlcontrol.GetControlRect(txt诊室.hwnd)
    Set rsTmp = zlDatabase.ShowRecSelect(Me, rsTmp, 0, "门诊诊室", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt诊室.Height, blnCancel, False, True)
        
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            txt诊室.Tag = rsTmp("名称"): txt诊室 = txt诊室.Tag
            If cbo范围.Enabled And cbo范围.Visible Then cbo范围.SetFocus
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdYS_Click()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim blnCanle As Boolean
    If txt接诊医生.Tag <> txt接诊医生 Then Exit Sub '由txt医生的Validate事件处理
    On Error GoTo errH
    strSql = "Select Distinct A.编号 as ID,A.姓名 as 名称,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D,临床部门 E" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And B.部门ID=D.部门ID" & _
        " And C.人员性质||''='医生' And D.服务对象 IN(1,3) And D.工作性质||''='临床'" & _
        " And B.部门ID In (Select 部门ID From 部门人员 Where 人员ID=[1])" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) And B.部门ID = E.部门ID And E.工作性质 = '20'" & _
        " Order by A.简码"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "接诊医生", False, "", "", False, False, False, 0, 0, txt接诊医生.Height, blnCanle, False, True, UserInfo.ID)
    If blnCanle Then Exit Sub
    If Not rsTmp Is Nothing Then txt接诊医生.Tag = rsTmp("名称"): txt接诊医生 = txt接诊医生.Tag: Me.cmdOK.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strPar As String
    Dim blnSetup As Boolean
    Dim i As Long
    Dim intType As Integer
    Dim strNotify As String
    Dim str诊室 As String
    
    blnSetup = InStr(1, ";" & mstrPrivs & ";", ";参数设置;") > 0
    gblnOK = False
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    On Error Resume Next
    str诊室 = zlDatabase.GetPara("本地诊室", glngSys, p急诊医生站, "", Array(lbl诊室, txt诊室, cmdSel), blnSetup)
    On Error GoTo 0
    
    On Error GoTo errH
    mbln诊间支付 = Val(zlDatabase.GetPara("急诊医嘱发送后启用诊间支付", glngSys, p门诊医嘱下达)) = 1
    cmdPBPSet.Enabled = mbln诊间支付
    '读取病人缺省科室范围
    strPar = zlDatabase.GetPara("接诊科室", glngSys, p急诊医生站, "", Array(lblEditDept, cbo科室), blnSetup)
    
    strSql = "Select Distinct B.ID,B.编码,B.名称,A.缺省" & _
        " From 部门人员 A,部门表 B,部门性质说明 C,临床部门 D" & _
        " Where A.部门ID=B.ID And B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null) And D.工作性质 = '20' And D.部门ID = A.部门ID" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And A.人员ID=[1]" & _
        " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        cbo科室.AddItem rsTmp!名称
        cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
        If rsTmp!ID = Val(strPar) Then
            cbo科室.ListIndex = cbo科室.NewIndex
        ElseIf Nvl(rsTmp!缺省, 0) = 1 And cbo科室.ListIndex = -1 Then
            cbo科室.ListIndex = cbo科室.NewIndex
        End If
        rsTmp.MoveNext
    Next
    Me.cbo范围.ListIndex = Val(zlDatabase.GetPara("接诊范围", glngSys, p急诊医生站, "2", Array(lbl范围, cbo范围), blnSetup)) - 1
    
    strSql = "Select 1 From 门诊诊室 E where e.名称=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str诊室)
    If Not rsTmp.EOF Then
        txt诊室.Text = str诊室
        txt诊室.Tag = str诊室
    End If
    
    '可以选择对其它医生就诊的病人进行就诊
    If InStr(mstrPrivs, "续诊病人") > 0 And InStr(mstrPrivs, "允许设置接诊医生") > 0 Then
        '可以选择本科室下的医生
        cmdYS.Enabled = True
        txt接诊医生.Enabled = True
    Else
        cmdYS.Enabled = False
        txt接诊医生.Enabled = False
    End If
    txt接诊医生.Tag = zlDatabase.GetPara("接诊医生", glngSys, p急诊医生站, UserInfo.姓名, Array(lbl医生, txt接诊医生, cmdYS), blnSetup)
    txt接诊医生.Text = txt接诊医生.Tag
    
   
    '刘兴洪:应用于排队叫号的呼叫人次:需要配合分诊台模块的排队叫号模式为１并且有排队呼叫站点=1时有效
    txtQueuePatis.Text = Val(zlDatabase.GetPara("医生就诊人数", glngSys, p急诊医生站, 3, Array(lblQueuePatis, txtQueuePatis), blnSetup))
    If txtQueuePatis.Enabled Then
        If Val(zlDatabase.GetPara("排队叫号模式", glngSys, p门诊分诊管理)) = 1 And Val(zlDatabase.GetPara("排队呼叫站点", glngSys, p门诊分诊管理, "0")) = 1 Then
            txtQueuePatis.Enabled = True
        Else
            txtQueuePatis.Enabled = False
        End If
    End If
    
    '发送完成后关闭医嘱窗体
    chkAutoClose.Value = Val(zlDatabase.GetPara("发送完成后关闭医嘱窗体", glngSys, p门诊医嘱下达, , Array(chkAutoClose), blnSetup))
    
    '找到病人后自动接诊
    chk自动接诊.Value = Val(zlDatabase.GetPara("找到病人后自动接诊", glngSys, p急诊医生站, , Array(chk自动接诊), blnSetup))
    
    '诊间支付允许使用预交款
    chkCanPay.Value = Val(zlDatabase.GetPara("诊间支付允许使用预交款", glngSys, p门诊医嘱下达, , Array(chkCanPay), blnSetup))
    
    '接诊后自动进行
    strPar = Val(zlDatabase.GetPara("接诊后自动进行", glngSys, p急诊医生站, , Array(chkAutoAdd, optAdd(0), optAdd(1)), blnSetup))
    
    If strPar = 2 Then
        chkAutoAdd.Value = 1
        optAdd(1).Value = True
    Else
        chkAutoAdd.Value = strPar
    End If
    
    '启用屏幕键盘
    chkStaKB.Value = Val(zlDatabase.GetPara("启用屏幕键盘", glngSys, p急诊医生站, , Array(chkStaKB), blnSetup))
    
    '消息提醒刷新
    strPar = zlDatabase.GetPara("自动刷新病历审阅间隔", glngSys, p急诊医生站, , Array(chkNotifyEPR), blnSetup, intType)
    If Val(strPar) > 0 Then
        chkNotifyEPR.Value = 1: txtNotifyEPR.Text = Val(strPar)
    End If
 
    If (intType = 3 Or intType = 15) And Not blnSetup Then
        txtNotifyEPR.Enabled = False
    End If
    
    strPar = zlDatabase.GetPara("自动刷新病历审阅天数", glngSys, p急诊医生站, 1, Array(lblNotifyEPRDay, txtNotifyEPRDay), blnSetup)
    txtNotifyEPRDay.Text = IIf(0 = Val(strPar), 1, Val(strPar))
        
    strNotify = zlDatabase.GetPara("自动刷新内容", glngSys, p急诊医生站, , Array(chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), lblArea), blnSetup)
    chkWarn(chkD危急值).Value = Val(Mid(strNotify, 1, 1))
    chkWarn(chkD医嘱安排).Value = Val(Mid(strNotify, 2, 1))
    chkWarn(chkD处方审查).Value = Val(Mid(strNotify, 3, 1))
    chkWarn(chkD传染病).Value = Val(Mid(strNotify, 4, 1))
    chkWarn(chkD备血完成).Value = Val(Mid(strNotify, 5, 1))
    chkWarn(chkD用血审核).Value = Val(Mid(strNotify, 6, 1))
    chkWarn(chkD输血反应).Value = Val(Mid(strNotify, 7, 1))
    chkWarn(chkD备血完成).Visible = gbln血库系统
    chkWarn(chkD用血审核).Visible = gbln血库系统
    chkWarn(chkD输血反应).Visible = gbln血库系统
    If InStr(mstrPrivs, "参数设置") = 0 Then
        chkWarn(chkD危急值).Enabled = False
        chkWarn(chkD医嘱安排).Enabled = False
        chkWarn(chkD处方审查).Enabled = False
        chkWarn(chkD传染病).Enabled = False
        chkWarn(chkD备血完成).Enabled = False
        chkWarn(chkD用血审核).Enabled = False
        chkWarn(chkD输血反应).Enabled = False
    End If
    chkAutoFinish.Value = Val(zlDatabase.GetPara("接诊时自动处理完成就诊", glngSys, p急诊医生站, , Array(chkAutoFinish), blnSetup))
    chkSound.Value = Val(zlDatabase.GetPara("启用语音提示", glngSys, p急诊医生站, , Array(chkSound, cmdSoundSet), blnSetup))
    
    '显示预约病人
    chkYYBR.Value = Val(zlDatabase.GetPara("显示预约病人", glngSys, p急诊医生站, "1", Array(chkYYBR), blnSetup))
    
    '急诊医嘱下达强制缺省药房
    chk缺省药房.Value = Val(zlDatabase.GetPara("门诊医嘱下达强制缺省药房", glngSys, p门诊医嘱下达, "1", Array(chk缺省药房), blnSetup))
    
    '急诊危急值弹窗提醒
    chk危急值.Value = Val(zlDatabase.GetPara("急诊危急值弹窗提醒", glngSys, p急诊医生站, "1", Array(chk危急值), blnSetup))
    
    chk上次诊断.Value = Val(zlDatabase.GetPara("上次诊断按科室提取", glngSys, p门诊医嘱下达, , Array(chk上次诊断), blnSetup))
    
    '医生站是否打印诊疗单据
    strPar = Val(zlDatabase.GetPara("门诊发送单据打印", glngSys, p门诊医嘱下达, , Array(opt急诊诊疗单打印(0), opt急诊诊疗单打印(1), opt急诊诊疗单打印(2)), blnSetup))
    opt急诊诊疗单打印(Val(strPar)) = True
    '医生站是否打印指引单
    strPar = Val(zlDatabase.GetPara("指引单打印方式", glngSys, p门诊医嘱下达, , Array(opt急诊指引单打印(0), opt急诊指引单打印(1), opt急诊指引单打印(2)), blnSetup))
    opt急诊指引单打印(Val(strPar)) = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
End Sub

Private Sub txt接诊医生_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim vRect As RECT, blnCancel As Boolean

    If txt接诊医生.Tag = txt接诊医生 Then Exit Sub

    strSql = "Select Distinct A.编号 as ID,A.姓名 as 名称,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D,临床部门 E" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And B.部门ID=D.部门ID" & _
        " And C.人员性质||''='医生' And D.服务对象 IN(1,3) And D.工作性质||''='临床' And B.部门ID = E.部门ID And E.工作性质 = '20'" & _
        " And B.部门ID In(Select 部门ID From 部门人员 Where 人员ID=[1])" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And (Upper(A.编号) Like [2] Or Upper(A.简码) Like [3] Or A.姓名 Like [3])" & _
        " Order by A.简码"
        
    vRect = zlcontrol.GetControlRect(txt接诊医生.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "接诊医生", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt接诊医生.Height, blnCancel, False, True, UserInfo.ID, UCase(txt接诊医生.Text) & "%", mstrLike & UCase(txt接诊医生.Text) & "%")
    If Not rsTmp Is Nothing Then
        txt接诊医生.Tag = rsTmp("名称")
        txt接诊医生 = txt接诊医生.Tag
    Else
        txt接诊医生.Tag = ""
        txt接诊医生 = ""
        Cancel = blnCancel
    End If
End Sub

Private Sub txt接诊医生_GotFocus()
    Call zlcontrol.TxtSelAll(txt接诊医生)
End Sub

Private Sub txt接诊医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt接诊医生 = "" Then txt接诊医生.Tag = "1"
        ZLCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt诊室_GotFocus()
    Call zlcontrol.TxtSelAll(txt诊室)
End Sub

Private Sub txt诊室_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt诊室 = "" Then txt诊室.Tag = "1"
        ZLCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt诊室_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If txt诊室.Tag = txt诊室 Then Exit Sub
    On Error GoTo errH
    If gbln挂号按排 Then
        strSql = "Select Distinct a.Id, a.名称, a.简码" & vbNewLine & _
            " From 门诊诊室 A, 门诊诊室适用科室 B, 部门人员 C, 上机人员表 D,临床部门 F" & vbNewLine & _
            " Where a.Id = b.诊室id And b.科室id = c.部门id And c.人员id = d.人员id And C.部门ID = F.部门ID And F.工作性质 = '20'" & vbNewLine & _
            " And (Upper(a.编码) Like [1] Or Upper(a.简码) Like [2] Or Upper(a.名称) Like [2])" & _
            "       And d.用户名 = User And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)"
            
            '如果没有查找到数据，则读取出所有的门诊诊室供选择
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(txt诊室.Text) & "%", mstrLike & UCase(txt诊室.Text) & "%")
    Else
        Set rsTmp = GetRs挂号安排诊室列表(1, UCase(txt诊室.Text), 1, 0, "", p急诊医生站)
    End If
        

    If rsTmp.EOF Then
        strSql = "Select a.Id, a.名称, a.简码 From 门诊诊室 A Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)" & _
            " And (Upper(a.编码) Like [1] Or Upper(a.简码) Like [2] Or Upper(a.名称) Like [2])"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(txt诊室.Text) & "%", mstrLike & UCase(txt诊室.Text) & "%")
    End If
        
    vRect = zlcontrol.GetControlRect(txt诊室.hwnd)
    Set rsTmp = zlDatabase.ShowRecSelect(Me, rsTmp, 0, "门诊诊室", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt诊室.Height, blnCancel, False, True)
    If Not rsTmp Is Nothing Then
        txt诊室.Tag = rsTmp("名称")
        txt诊室 = txt诊室.Tag
    Else
        txt诊室.Tag = ""
        txt诊室 = ""
        Cancel = blnCancel
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




