VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutStationSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmOutStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk危急值 
      Caption         =   "危急值弹窗提醒"
      Height          =   240
      Left            =   4365
      TabIndex        =   57
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Frame fra门诊诊疗单打印 
      Caption         =   "门诊发送后,诊疗单据"
      Height          =   680
      Index           =   0
      Left            =   120
      TabIndex        =   51
      Top             =   5280
      Width           =   4695
      Begin VB.OptionButton opt门诊诊疗单打印 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   300
         Width           =   1560
      End
      Begin VB.OptionButton opt门诊诊疗单打印 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   53
         Top             =   300
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton opt门诊诊疗单打印 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   2
         Left            =   3360
         TabIndex        =   52
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame fra门诊指引单打印 
      Caption         =   "门诊发送后,指引单"
      Height          =   680
      Left            =   5280
      TabIndex        =   47
      Top             =   5280
      Width           =   4695
      Begin VB.OptionButton opt门诊指引单打印 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   2
         Left            =   3360
         TabIndex        =   50
         Top             =   300
         Width           =   1200
      End
      Begin VB.OptionButton opt门诊指引单打印 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   49
         Top             =   300
         Width           =   1560
      End
      Begin VB.OptionButton opt门诊指引单打印 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   300
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.CheckBox chkYYBR 
      Caption         =   "候诊列表中显示预约病人"
      Height          =   240
      Left            =   4365
      TabIndex        =   46
      Top             =   1800
      Width           =   2340
   End
   Begin VB.CheckBox chkCanPay 
      Caption         =   "诊间支付允许使用预交款"
      Height          =   250
      Left            =   4365
      TabIndex        =   45
      Top             =   600
      Width           =   2310
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "全清"
      Height          =   300
      Index           =   1
      Left            =   9540
      TabIndex        =   44
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "全选"
      Height          =   300
      Index           =   0
      Left            =   8880
      TabIndex        =   43
      Top             =   120
      Width           =   600
   End
   Begin VB.CheckBox chkAutoClose 
      Caption         =   "发送完成后自动关闭医嘱窗体"
      Height          =   195
      Left            =   135
      TabIndex        =   40
      Top             =   3480
      Width           =   2745
   End
   Begin VB.CheckBox chkAutoFinish 
      Caption         =   "接诊病人时自动处理上一个病人完成就诊或需回诊"
      Height          =   195
      Left            =   135
      TabIndex        =   37
      Top             =   3135
      Width           =   6105
   End
   Begin VB.Frame fraEPR 
      Caption         =   "提醒设置"
      Height          =   1410
      Left            =   135
      TabIndex        =   24
      Top             =   3765
      Width           =   6480
      Begin VB.CheckBox chkWarn 
         Caption         =   "输血反应"
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   56
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "用血审核"
         Height          =   195
         Index           =   5
         Left            =   4320
         TabIndex        =   55
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "备血完成"
         Height          =   195
         Index           =   4
         Left            =   3255
         TabIndex        =   36
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "启用语音提示"
         Height          =   195
         Left            =   4125
         TabIndex        =   39
         Top             =   285
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "语音设置(&S)"
         Height          =   350
         Left            =   4125
         TabIndex        =   38
         Top             =   555
         Width           =   1410
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "传染病"
         Height          =   195
         Index           =   3
         Left            =   2330
         TabIndex        =   35
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "处方审查"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   34
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "医嘱安排"
         Height          =   195
         Index           =   1
         Left            =   2330
         TabIndex        =   33
         Top             =   855
         Width           =   1035
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
         Top             =   510
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPRDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   690
         TabIndex        =   30
         Top             =   720
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
         Top             =   450
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
         Top             =   255
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "危急值"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   25
         Top             =   855
         Width           =   1035
      End
      Begin VB.CheckBox chkNotifyEPR 
         Caption         =   "每    分钟自动刷新提醒区域中的内容"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   270
         Width           =   3450
      End
      Begin VB.Label lblNotifyEPRDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "将    天内产生的消息显示在提醒区域"
         Height          =   180
         Left            =   480
         TabIndex        =   32
         Top             =   540
         Width           =   3060
      End
      Begin VB.Label lblArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提醒内容:"
         Height          =   180
         Left            =   465
         TabIndex        =   29
         Top             =   855
         Width           =   810
      End
   End
   Begin VB.Frame fraReceive 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   105
      TabIndex        =   20
      Top             =   2490
      Width           =   6360
      Begin VB.OptionButton optAdd 
         Caption         =   "新增医嘱,切换到病历时新增病历"
         Enabled         =   0   'False
         Height          =   260
         Index           =   0
         Left            =   330
         TabIndex        =   23
         Top             =   300
         Value           =   -1  'True
         Width           =   2940
      End
      Begin VB.CheckBox chkAutoAdd 
         Caption         =   "病人接诊后自动进行"
         Height          =   195
         Left            =   45
         TabIndex        =   22
         Top             =   90
         Width           =   2055
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "新增病历,切换到医嘱时新增医嘱"
         Enabled         =   0   'False
         Height          =   260
         Index           =   1
         Left            =   3390
         TabIndex        =   21
         Top             =   300
         Width           =   2940
      End
   End
   Begin VB.CommandButton cmdPBPSet 
      Caption         =   "支付票据打印设置"
      Height          =   300
      Left            =   4365
      TabIndex        =   19
      Top             =   210
      Width           =   1620
   End
   Begin VB.CheckBox chkStaKB 
      Caption         =   "启用屏幕键盘"
      Height          =   250
      Left            =   4365
      TabIndex        =   18
      Top             =   930
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
      Top             =   2430
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
      Top             =   2265
      Width           =   465
   End
   Begin VB.CheckBox chk自动接诊 
      Caption         =   "查找到候诊病人之后自动接诊"
      Height          =   500
      Left            =   4365
      TabIndex        =   10
      Top             =   1230
      Width           =   2070
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   350
      Left            =   135
      TabIndex        =   11
      Top             =   6120
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Caption         =   " 就诊参数 "
      Height          =   1935
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
         Top             =   1515
         Width           =   255
      End
      Begin VB.TextBox txt接诊医生 
         Height          =   300
         Left            =   1020
         TabIndex        =   8
         Top             =   1485
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
         ItemData        =   "frmOutStationSetup.frx":000C
         Left            =   1020
         List            =   "frmOutStationSetup.frx":0019
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
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Label lbl医生 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "接诊医生"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   1545
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
         Y1              =   1395
         Y2              =   1395
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8925
      TabIndex        =   13
      Top             =   6120
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7680
      TabIndex        =   12
      Top             =   6120
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwEPRList 
      Height          =   4680
      Left            =   6720
      TabIndex        =   41
      Top             =   480
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   8255
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ColHdrIcons     =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "报表"
         Object.Width           =   6615
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "类型"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   7320
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutStationSetup.frx":0037
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutStationSetup.frx":05D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutStationSetup.frx":0B6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutStationSetup.frx":1105
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutStationSetup.frx":169F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutStationSetup.frx":1C39
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutStationSetup.frx":21D3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblEPRList 
      AutoSize        =   -1  'True
      Caption         =   "门诊病历缺省页签"
      Height          =   180
      Left            =   6720
      TabIndex        =   42
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label lblQueuePatis 
      AutoSize        =   -1  'True
      Caption         =   "医生最多能呼叫      人"
      Height          =   180
      Left            =   135
      TabIndex        =   15
      ToolTipText     =   "表示本次医生最多能呼叫多少个病人来就诊,超过后，就不能再次呼叫;此参数需要配合分诊台模块的排队叫号模式为医生主动呼叫有效"
      Top             =   2265
      Width           =   1980
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   -240
      X2              =   10455
      Y1              =   6020
      Y2              =   6020
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   10320
      Y1              =   6040
      Y2              =   6040
   End
End
Attribute VB_Name = "frmOutStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Private mstrLike As String
Private mobjSquareCard As Object
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
    Call zlCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdOK_Click()
    Dim str病人接诊控制 As String '问题号:57566
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
    
    Call zlDatabase.SetPara("本地诊室", Me.txt诊室.Text, glngSys, p门诊医生站, blnHavePara)
    Call zlDatabase.SetPara("接诊范围", Me.cbo范围.ItemData(Me.cbo范围.ListIndex), glngSys, p门诊医生站, blnHavePara)
    Call zlDatabase.SetPara("接诊医生", Me.txt接诊医生.Text, glngSys, p门诊医生站, blnHavePara)
    
    '刘兴洪:应用于排队叫号的呼叫人次:需要配合分诊台模块的排队叫号模式为１并且有排队呼叫对象=2时有效
    If txtQueuePatis.Enabled Then
        Call zlDatabase.SetPara("医生就诊人数", Val(Me.txtQueuePatis.Text), glngSys, p门诊医生站, blnHavePara)
    End If
    '接诊科室
    Call zlDatabase.SetPara("接诊科室", cbo科室.ItemData(cbo科室.ListIndex), glngSys, p门诊医生站, blnHavePara)
    
    '发送完成后关闭医嘱窗体
    Call zlDatabase.SetPara("发送完成后关闭医嘱窗体", chkAutoClose.Value, glngSys, p门诊医嘱下达, blnHavePara)
    
    '找到病人后自动接诊
    Call zlDatabase.SetPara("找到病人后自动接诊", chk自动接诊.Value, glngSys, p门诊医生站, blnHavePara)
    
    '接诊后自动进行
    If optAdd(1).Value And optAdd(1).Enabled Then
        Call zlDatabase.SetPara("接诊后自动进行", 2, glngSys, p门诊医生站, blnHavePara)
    Else
        Call zlDatabase.SetPara("接诊后自动进行", chkAutoAdd.Value, glngSys, p门诊医生站, blnHavePara)
    End If

    '启用屏幕键盘
    Call zlDatabase.SetPara("启用屏幕键盘", chkStaKB.Value, glngSys, p门诊医生站, blnHavePara)
    
    Call zlDatabase.SetPara("自动刷新病历审阅间隔", IIf(chkNotifyEPR.Value = 1, Val(txtNotifyEPR.Text), ""), glngSys, p门诊医生站, blnHavePara)
    Call zlDatabase.SetPara("自动刷新病历审阅天数", Val(txtNotifyEPRDay.Text), glngSys, p门诊医生站, blnHavePara)
    strTmp = ""
    For i = chkD危急值 To chkD输血反应
        strTmp = strTmp & chkWarn(i).Value
    Next
    Call zlDatabase.SetPara("自动刷新内容", strTmp, glngSys, p门诊医生站, blnHavePara)
    Call zlDatabase.SetPara("接诊时自动处理完成就诊", chkAutoFinish.Value, glngSys, p门诊医生站, blnHavePara)
    Call zlDatabase.SetPara("启用语音提示", chkSound.Value, glngSys, p门诊医生站, blnHavePara)
    strTmp = ""
    For i = 1 To lvwEPRList.ListItems.Count
        If lvwEPRList.ListItems(i).Checked Then
            strTmp = strTmp & "," & Mid(lvwEPRList.ListItems(i).Key, 2)
        End If
    Next
    strTmp = Mid(strTmp, 2)
    Call zlDatabase.SetPara("门诊病历缺省页签", strTmp, glngSys, p门诊医嘱下达, blnHavePara)
    
    Call zlDatabase.SetPara("诊间支付允许使用预交款", chkCanPay.Value, glngSys, p门诊医嘱下达, blnHavePara)
    
    Call zlDatabase.SetPara("显示预约病人", chkYYBR.Value, glngSys, p门诊医生站, blnHavePara)
    
    Call zlDatabase.SetPara("门诊危急值弹窗提醒", chk危急值.Value, glngSys, p门诊医生站, blnHavePara)
    
    '医生站是否打印诊疗单据
    Call zlDatabase.SetPara("门诊发送单据打印", IIf(opt门诊诊疗单打印(0).Value = True, 0, IIf(opt门诊诊疗单打印(1).Value = True, 1, 2)), glngSys, p门诊医嘱下达, blnHavePara)
    '医生站是否打印指引单
    Call zlDatabase.SetPara("指引单打印方式", IIf(opt门诊指引单打印(0).Value = True, 0, IIf(opt门诊指引单打印(1).Value = True, 1, 2)), glngSys, p门诊医嘱下达, blnHavePara)

    gblnOK = True
    Unload Me
End Sub

Private Sub chkNotifyEPR_Click()
    txtNotifyEPR.Enabled = chkNotifyEPR.Value = 1
    If Visible And txtNotifyEPR.Enabled Then txtNotifyEPR.SetFocus
End Sub

Private Sub cmdS_Click(Index As Integer)
    Dim i As Long
    For i = 1 To lvwEPRList.ListItems.Count
        lvwEPRList.ListItems(i).Checked = Index = 0
    Next
End Sub

Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 0)
End Sub

Private Sub lvwEPRList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
    lvwEPRList.ToolTipText = Item.Tag
End Sub

Private Sub lvwEPRList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lvwEPRList.ToolTipText = Item.Tag
End Sub

Private Sub txtNotifyEPR_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPR)
End Sub

Private Sub txtNotifyEPR_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPRDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPRDay)
End Sub

Private Sub txtNotifyEPRDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cmdPBPSet_Click()
    
    On Error Resume Next
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, p门诊医嘱下达, glngSys, gstrDBUser, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
            err.Clear: Exit Sub
        End If
    End If
    Call mobjSquareCard.zlCliniqueRoomPayPrintSet(Me)
    err.Clear: On Error GoTo 0
End Sub

Private Sub cmdSel_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If txt诊室.Tag <> txt诊室 Then Exit Sub '由txt诊室的Validate事件处理
    
    If gbln挂号按排 Then
        strSQL = "Select Distinct a.Id, a.名称, a.简码" & vbNewLine & _
            " From 门诊诊室 A, 门诊诊室适用科室 B, 部门人员 C, 上机人员表 D" & vbNewLine & _
            " Where a.Id = b.诊室id And b.科室id = c.部门id And c.人员id = d.人员id" & vbNewLine & _
            "       And d.用户名 = User And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)"
    Else
        strSQL = "Select Distinct e.编码 As ID,e.名称,e.简码" & vbNewLine & _
               "From 门诊诊室 E, 挂号安排诊室 D, 挂号安排 C, 部门人员 A, 上机人员表 B" & vbNewLine & _
               "Where a.人员id = b.人员id And b.用户名 = User And c.科室id = a.部门id And c.Id = d.号表id And e.名称 = d.门诊诊室 " & _
               " And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null)"
    End If
    '如果没有查找到数据，则读取出所有的门诊诊室供选择
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        strSQL = "Select a.Id, a.名称, a.简码 From 门诊诊室 A Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)"
    End If
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "门诊诊室", , , , , , , txt诊室.Left, txt诊室.Top, txt诊室.Height, , , True)
    If Not rsTmp Is Nothing Then
        txt诊室.Tag = rsTmp("名称"): txt诊室 = txt诊室.Tag
        If cbo范围.Enabled And cbo范围.Visible Then cbo范围.SetFocus
    End If
End Sub

Private Sub cmdYS_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnCanle As Boolean
    If txt接诊医生.Tag <> txt接诊医生 Then Exit Sub '由txt医生的Validate事件处理
            
    strSQL = "Select Distinct A.编号 as ID,A.姓名 as 名称,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And B.部门ID=D.部门ID" & _
        " And C.人员性质||''='医生' And D.服务对象 IN(1,3) And D.工作性质||''='临床'" & _
        " And B.部门ID In (Select 部门ID From 部门人员 Where 人员ID=[1])" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.简码"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "接诊医生", False, "", "", False, False, False, 0, 0, txt接诊医生.Height, blnCanle, False, True, UserInfo.ID)
    If blnCanle Then Exit Sub
    If Not rsTmp Is Nothing Then txt接诊医生.Tag = rsTmp("名称"): txt接诊医生 = txt接诊医生.Tag: Me.cmdOK.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String
    Dim blnSetup As Boolean
    Dim i As Long
    Dim str病人接诊控制 As String  '问题号:57566
    Dim intType As Integer
    Dim strNotify As String
    Dim str诊室 As String
    
    blnSetup = InStr(1, ";" & mstrPrivs & ";", ";参数设置;") > 0
    gblnOK = False
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    On Error Resume Next
    str诊室 = zlDatabase.GetPara("本地诊室", glngSys, p门诊医生站, "", Array(lbl诊室, txt诊室, cmdSel), blnSetup)
    On Error GoTo 0
    
    On Error GoTo errH
    mbln诊间支付 = Val(zlDatabase.GetPara("门诊医嘱发送后启用诊间支付", glngSys, p门诊医嘱下达)) = 1
    cmdPBPSet.Enabled = mbln诊间支付
    '读取病人缺省科室范围
    strPar = zlDatabase.GetPara("接诊科室", glngSys, p门诊医生站, "", Array(lblEditDept, cbo科室), blnSetup)
    
    strSQL = "Select Distinct B.ID,B.编码,B.名称,A.缺省" & _
        " From 部门人员 A,部门表 B,部门性质说明 C" & _
        " Where A.部门ID=B.ID And B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And A.人员ID=[1]" & _
        " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        cbo科室.AddItem rsTmp!名称
        cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
        If rsTmp!ID = Val(strPar) Then
            cbo科室.ListIndex = cbo科室.NewIndex
        ElseIf NVL(rsTmp!缺省, 0) = 1 And cbo科室.ListIndex = -1 Then
            cbo科室.ListIndex = cbo科室.NewIndex
        End If
        rsTmp.MoveNext
    Next
    Me.cbo范围.ListIndex = Val(zlDatabase.GetPara("接诊范围", glngSys, p门诊医生站, "2", Array(lbl范围, cbo范围), blnSetup)) - 1
    
    strSQL = "Select 1 From 门诊诊室 E where e.名称=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str诊室)
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
    txt接诊医生.Tag = zlDatabase.GetPara("接诊医生", glngSys, p门诊医生站, UserInfo.姓名, Array(lbl医生, txt接诊医生, cmdYS), blnSetup)
    txt接诊医生.Text = txt接诊医生.Tag
    
   
    '刘兴洪:应用于排队叫号的呼叫人次:需要配合分诊台模块的排队叫号模式为１并且有排队呼叫站点=1时有效
    txtQueuePatis.Text = Val(zlDatabase.GetPara("医生就诊人数", glngSys, p门诊医生站, 3, Array(lblQueuePatis, txtQueuePatis), blnSetup))
    If txtQueuePatis.Enabled Then
        txtQueuePatis.Enabled = CheckDoctorPatisIsValid
    End If
    
    '发送完成后关闭医嘱窗体
    chkAutoClose.Value = Val(zlDatabase.GetPara("发送完成后关闭医嘱窗体", glngSys, p门诊医嘱下达, , Array(chkAutoClose), blnSetup))
    
    '找到病人后自动接诊
    chk自动接诊.Value = Val(zlDatabase.GetPara("找到病人后自动接诊", glngSys, p门诊医生站, , Array(chk自动接诊), blnSetup))
    
    '诊间支付允许使用预交款
    chkCanPay.Value = Val(zlDatabase.GetPara("诊间支付允许使用预交款", glngSys, p门诊医嘱下达, , Array(chkCanPay), blnSetup))
    
    '接诊后自动进行
    strPar = Val(zlDatabase.GetPara("接诊后自动进行", glngSys, p门诊医生站, , Array(chkAutoAdd, optAdd(0), optAdd(1)), blnSetup))
    
    If strPar = 2 Then
        chkAutoAdd.Value = 1
        optAdd(1).Value = True
    Else
        chkAutoAdd.Value = strPar
    End If
    
    '启用屏幕键盘
    chkStaKB.Value = Val(zlDatabase.GetPara("启用屏幕键盘", glngSys, p门诊医生站, , Array(chkStaKB), blnSetup))
    
    '消息提醒刷新
    strPar = zlDatabase.GetPara("自动刷新病历审阅间隔", glngSys, p门诊医生站, , Array(chkNotifyEPR), blnSetup, intType)
    If Val(strPar) > 0 Then
        chkNotifyEPR.Value = 1: txtNotifyEPR.Text = Val(strPar)
    End If
 
    If (intType = 3 Or intType = 15) And Not blnSetup Then
        txtNotifyEPR.Enabled = False
    End If
    
    strPar = zlDatabase.GetPara("自动刷新病历审阅天数", glngSys, p门诊医生站, 1, Array(lblNotifyEPRDay, txtNotifyEPRDay), blnSetup)
    txtNotifyEPRDay.Text = IIf(0 = Val(strPar), 1, Val(strPar))
        
    strNotify = zlDatabase.GetPara("自动刷新内容", glngSys, p门诊医生站, , Array(chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), lblArea), blnSetup)
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
    chkAutoFinish.Value = Val(zlDatabase.GetPara("接诊时自动处理完成就诊", glngSys, p门诊医生站, , Array(chkAutoFinish), blnSetup))
    chkSound.Value = Val(zlDatabase.GetPara("启用语音提示", glngSys, p门诊医生站, , Array(chkSound, cmdSoundSet), blnSetup))
    strPar = zlDatabase.GetPara("门诊病历缺省页签", glngSys, p门诊医嘱下达)
    Call Load缺省病历(strPar)
    
    '显示预约病人
    chkYYBR.Value = Val(zlDatabase.GetPara("显示预约病人", glngSys, p门诊医生站, "1", Array(chkYYBR), blnSetup))
    
    '门诊危急值弹窗提醒
    chk危急值.Value = Val(zlDatabase.GetPara("门诊危急值弹窗提醒", glngSys, p门诊医生站, "1", Array(chk危急值), blnSetup))
    
    '医生站是否打印诊疗单据
    strPar = Val(zlDatabase.GetPara("门诊发送单据打印", glngSys, p门诊医嘱下达, , Array(opt门诊诊疗单打印(0), opt门诊诊疗单打印(1), opt门诊诊疗单打印(2)), blnSetup))
    opt门诊诊疗单打印(Val(strPar)) = True
    '医生站是否打印指引单
    strPar = Val(zlDatabase.GetPara("指引单打印方式", glngSys, p门诊医嘱下达, , Array(opt门诊指引单打印(0), opt门诊指引单打印(1), opt门诊指引单打印(2)), blnSetup))
    opt门诊指引单打印(Val(strPar)) = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    Set mobjSquareCard = Nothing
End Sub

Private Sub txt接诊医生_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, blnCancel As Boolean

    If txt接诊医生.Tag = txt接诊医生 Then Exit Sub

    strSQL = "Select Distinct A.编号 as ID,A.姓名 as 名称,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And B.部门ID=D.部门ID" & _
        " And C.人员性质||''='医生' And D.服务对象 IN(1,3) And D.工作性质||''='临床'" & _
        " And B.部门ID In(Select 部门ID From 部门人员 Where 人员ID=[1])" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And (Upper(A.编号) Like [2] Or Upper(A.简码) Like [3] Or Upper(A.姓名) Like [3])" & _
        " Order by A.简码"
        
    vRect = zlControl.GetControlRect(txt接诊医生.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "接诊医生", False, "", "", False, False, True, _
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
    Call zlControl.TxtSelAll(txt接诊医生)
End Sub

Private Sub txt接诊医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt接诊医生 = "" Then txt接诊医生.Tag = "1"
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt诊室_GotFocus()
    Call zlControl.TxtSelAll(txt诊室)
End Sub

Private Sub txt诊室_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt诊室 = "" Then txt诊室.Tag = "1"
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt诊室_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If txt诊室.Tag = txt诊室 Then Exit Sub
    
    If gbln挂号按排 Then
        strSQL = "Select Distinct a.Id, a.名称, a.简码" & vbNewLine & _
            " From 门诊诊室 A, 门诊诊室适用科室 B, 部门人员 C, 上机人员表 D" & vbNewLine & _
            " Where a.Id = b.诊室id And b.科室id = c.部门id And c.人员id = d.人员id" & vbNewLine & _
            " And (Upper(a.编码) Like [1] Or Upper(a.简码) Like [2] Or Upper(a.名称) Like [2])" & _
            "       And d.用户名 = User And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)"
    Else
        strSQL = "Select Distinct e.编码 As ID,e.名称,e.简码" & vbNewLine & _
                "From 门诊诊室 E, 挂号安排诊室 D, 挂号安排 C, 部门人员 A, 上机人员表 B" & vbNewLine & _
                "Where a.人员id = b.人员id And b.用户名 = User And c.科室id = a.部门id And c.Id = d.号表id And e.名称 = d.门诊诊室 " & _
                " And (Upper(E.编码) Like [1] Or Upper(E.简码) Like [2] Or Upper(E.名称) Like [2])" & _
                " And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) "
    End If
        
    '如果没有查找到数据，则读取出所有的门诊诊室供选择
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(txt诊室.Text) & "%", mstrLike & UCase(txt诊室.Text) & "%")
    If rsTmp.EOF Then
        strSQL = "Select a.Id, a.名称, a.简码 From 门诊诊室 A Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)" & _
            " And (Upper(a.编码) Like [1] Or Upper(a.简码) Like [2] Or Upper(a.名称) Like [2])"
    End If
        
    vRect = zlControl.GetControlRect(txt诊室.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "门诊诊室", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt诊室.Height, blnCancel, False, True, UCase(txt诊室.Text) & "%", mstrLike & UCase(txt诊室.Text) & "%")
    If Not rsTmp Is Nothing Then
        txt诊室.Tag = rsTmp("名称")
        txt诊室 = txt诊室.Tag
    Else
        txt诊室.Tag = ""
        txt诊室 = ""
        Cancel = blnCancel
    End If
End Sub

Private Sub Load缺省病历(ByVal strPar As String)
'功能：加载老版病历缺省清单
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem
    Dim str种类名 As String
    Dim lngIcon As Long
    Dim objTmp As Object
    
    On Error GoTo errH
    
    strSQL = "Select B.ID" & _
        " From 部门人员 A,部门表 B,部门性质说明 C" & _
        " Where A.部门ID=B.ID And B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And A.人员ID=[1]"
        
    strSQL = "Select F.ID,F.名称,f.种类 From 病历文件列表 F,病历应用科室 A,(" & strSQL & ") b " & _
     " Where F.ID=A.文件id(+) And f.种类 In (1,5,6) And f.保留 <> 4  And (f.通用 = 1 Or f.通用 = 2 and A.科室id =b.id)" & _
     " Group By F.ID,F.名称,f.种类,f.编号 Order By f.种类,f.编号"
     
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    Do While Not rsTmp.EOF
        If Val(rsTmp!种类 & "") = 5 Then
            lngIcon = 6
            str种类名 = "疾病证明报告"
        ElseIf Val(rsTmp!种类 & "") = 6 Then
            lngIcon = 7
            str种类名 = "知情文件"
        Else
            lngIcon = 2
            str种类名 = "门诊病历"
        End If
        Set objItem = lvwEPRList.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称, , lngIcon)
        objItem.Tag = rsTmp!名称 & "【" & str种类名 & "】"
        objItem.SubItems(1) = str种类名
        If InStr("," & strPar & ",", "," & Val(rsTmp!ID) & ",") > 0 Then
            objItem.Checked = True
        End If
        rsTmp.MoveNext
    Loop
    '新病历
    Set rsTmp = Nothing
    On Error Resume Next
    Set objTmp = CreateObject("zl9EmrInterface.ClsEmrInterface")
    If Not objTmp Is Nothing Then
        strSQL = "Select Rawtohex(ID) as ID,Title as 名称 From Antetype_List Where Kind in ('01','04','05') and nvl(disable,0)=0 Order By Code"
        Call gobjEmr.OpenSQLRecordset(strSQL, "", rsTmp)
    End If
    err.Clear: On Error GoTo 0
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            lngIcon = 1
            str种类名 = "新版病历"
            Set objItem = lvwEPRList.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称, , lngIcon)
            objItem.Tag = rsTmp!名称 & "【" & str种类名 & "】"
            objItem.SubItems(1) = str种类名
            If InStr("," & strPar & ",", "," & rsTmp!ID & ",") > 0 Then
                objItem.Checked = True
            End If
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
