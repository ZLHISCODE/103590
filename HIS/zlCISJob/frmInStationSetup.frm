VERSION 5.00
Begin VB.Form frmInStationSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmInStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraAdvice 
      Caption         =   "提醒设置 "
      Height          =   2340
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4800
      Begin VB.CheckBox chkWarn 
         Caption         =   "血袋回收"
         Height          =   195
         Index           =   12
         Left            =   465
         TabIndex        =   55
         Top             =   1710
         Width           =   1020
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "备血完成"
         Height          =   195
         Index           =   11
         Left            =   3750
         TabIndex        =   53
         Top             =   1455
         Width           =   1020
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "标本拒收"
         Height          =   195
         Index           =   10
         Left            =   2740
         TabIndex        =   52
         Top             =   1455
         Width           =   1040
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "取血通知"
         Height          =   195
         Index           =   9
         Left            =   1740
         TabIndex        =   50
         Top             =   1455
         Width           =   1040
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RIS预约准备"
         Height          =   195
         Index           =   8
         Left            =   465
         TabIndex        =   43
         Top             =   1455
         Width           =   1365
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RIS预约"
         Height          =   195
         Index           =   7
         Left            =   3615
         TabIndex        =   42
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkSoundHS 
         Caption         =   "启用语音提示"
         Height          =   195
         Left            =   300
         TabIndex        =   40
         Top             =   1980
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundHSSet 
         Caption         =   "语音设置(&S)"
         Height          =   350
         Left            =   1830
         TabIndex        =   39
         Top             =   1890
         Width           =   1410
      End
      Begin VB.TextBox txtNotifyAdvice 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "10"
         Top             =   315
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdvice 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   17
         Top             =   495
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdviceDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   16
         Top             =   765
         Width           =   300
      End
      Begin VB.TextBox txtNotifyAdviceDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "1"
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "新开"
         Height          =   195
         Index           =   0
         Left            =   1125
         TabIndex        =   14
         Top             =   915
         Width           =   675
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "新停"
         Height          =   195
         Index           =   1
         Left            =   1875
         TabIndex        =   13
         Top             =   915
         Width           =   675
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "新废"
         Height          =   195
         Index           =   2
         Left            =   2700
         TabIndex        =   12
         Top             =   915
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "安排"
         Height          =   195
         Index           =   3
         Left            =   3495
         TabIndex        =   11
         Top             =   915
         Width           =   675
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "危急值"
         Height          =   195
         Index           =   4
         Left            =   465
         TabIndex        =   10
         Top             =   1185
         Width           =   870
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "输液拒绝"
         Height          =   195
         Index           =   5
         Left            =   1395
         TabIndex        =   9
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "销帐申请"
         Height          =   195
         Index           =   6
         Left            =   2535
         TabIndex        =   8
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkNotifyAdvice 
         Caption         =   "每    分钟自动刷新医嘱提醒区域中的内容"
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   330
         Width           =   3900
      End
      Begin VB.Label lbl提醒内容 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提醒内容:"
         Height          =   180
         Left            =   300
         TabIndex        =   21
         Top             =   915
         Width           =   810
      End
      Begin VB.Label lblNotifyAdviceDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "将    天内处理的医嘱病人显示在提醒区域"
         Height          =   180
         Left            =   570
         TabIndex        =   20
         Top             =   600
         Width           =   3420
      End
   End
   Begin VB.Frame fraEPR 
      Caption         =   "提醒设置"
      Height          =   2340
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   4800
      Begin VB.CheckBox chk危急值 
         Caption         =   "危急值弹窗提醒"
         Height          =   240
         Left            =   120
         TabIndex        =   56
         Top             =   1975
         Width           =   1590
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "输血反应"
         Height          =   195
         Index           =   26
         Left            =   3720
         TabIndex        =   54
         Top             =   1635
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "用血审核"
         Height          =   195
         Index           =   25
         Left            =   2565
         TabIndex        =   51
         Top             =   1635
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "校对疑问"
         Height          =   195
         Index           =   24
         Left            =   1440
         TabIndex        =   45
         Top             =   1635
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "备血完成"
         Height          =   195
         Index           =   23
         Left            =   3720
         TabIndex        =   44
         Top             =   1380
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "病历质控"
         Height          =   195
         Index           =   22
         Left            =   2565
         TabIndex        =   41
         Top             =   1380
         Width           =   1035
      End
      Begin VB.CheckBox chkSoundYS 
         Caption         =   "启用语音提示"
         Height          =   195
         Left            =   1800
         TabIndex        =   38
         Top             =   1998
         Width           =   1455
      End
      Begin VB.CommandButton cmdSoundYSSet 
         Caption         =   "语音设置(&S)"
         Height          =   350
         Left            =   3240
         TabIndex        =   37
         Top             =   1920
         Width           =   1410
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "传染病"
         Height          =   195
         Index           =   21
         Left            =   1440
         TabIndex        =   36
         Top             =   1380
         Width           =   855
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "处方审查"
         Height          =   195
         Index           =   20
         Left            =   3720
         TabIndex        =   35
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "医嘱审核"
         Height          =   195
         Index           =   19
         Left            =   2565
         TabIndex        =   34
         Top             =   1125
         Width           =   1035
      End
      Begin VB.TextBox txtNotifyEPR 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "10"
         Top             =   270
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPR 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   29
         Top             =   510
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPRDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   28
         Top             =   780
         Width           =   300
      End
      Begin VB.TextBox txtNotifyEPRDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "1"
         Top             =   600
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "病历审阅"
         Height          =   195
         Index           =   15
         Left            =   1440
         TabIndex        =   26
         Top             =   885
         Width           =   1065
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "医嘱安排"
         Height          =   195
         Index           =   16
         Left            =   2565
         TabIndex        =   25
         Top             =   885
         Width           =   1020
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "危急值"
         Height          =   195
         Index           =   17
         Left            =   3720
         TabIndex        =   24
         Top             =   885
         Width           =   885
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "报告撤消"
         Height          =   195
         Index           =   18
         Left            =   1440
         TabIndex        =   23
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkNotifyEPR 
         Caption         =   "每    分钟自动刷新提醒区域中的内容"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   280
         Width           =   3900
      End
      Begin VB.Label Label提醒区域 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提醒内容:"
         Height          =   180
         Left            =   600
         TabIndex        =   33
         Top             =   880
         Width           =   810
      End
      Begin VB.Label lblNotifyEPRDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "将    天内完成的内容显示在提醒区域"
         Height          =   180
         Left            =   615
         TabIndex        =   32
         Top             =   615
         Width           =   3060
      End
   End
   Begin VB.Frame fra住院诊疗单打印 
      Caption         =   "医嘱发送后,诊疗单据"
      Height          =   630
      Left            =   120
      TabIndex        =   46
      Top             =   2520
      Width           =   4815
      Begin VB.OptionButton opt住院诊疗单打印 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   49
         Top             =   300
         Width           =   840
      End
      Begin VB.OptionButton opt住院诊疗单打印 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   48
         Top             =   300
         Width           =   1440
      End
      Begin VB.OptionButton opt住院诊疗单打印 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   2
         Left            =   2880
         TabIndex        =   47
         Top             =   300
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.Frame fraBaby 
      Caption         =   "医嘱处理缺省范围(含提醒)"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   4815
      Begin VB.OptionButton optBaby 
         Caption         =   "病人医嘱"
         Height          =   180
         Index           =   1
         Left            =   1440
         TabIndex        =   6
         Top             =   285
         Width           =   1200
      End
      Begin VB.OptionButton optBaby 
         Caption         =   "全部医嘱"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optBaby 
         Caption         =   "婴儿医嘱"
         Height          =   180
         Index           =   2
         Left            =   2880
         TabIndex        =   4
         Top             =   285
         Width           =   1440
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   3915
      Width           =   5055
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3720
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2535
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   8040
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   8040
         Y1              =   15
         Y2              =   0
      End
   End
End
Attribute VB_Name = "frmInStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbln护士站 As Boolean
Public mstrPrivs As String
Private mlngModual As Long

Private Enum Enum_chkWarn
    '护士站提醒参数
    chkN新开 = 0
    chkN新停 = 1
    chkN新废 = 2
    chkN安排 = 3
    chkN危急值 = 4
    chkN输液拒绝 = 5
    chkN销帐申请 = 6
    chkNRIS预约 = 7
    chkNRIS预约准备 = 8
    chk取血通知 = 9
    chk标本拒收 = 10
    chk备血完成 = 11
    chk血袋回收 = 12
    
    
    '医生站提醒参数
    chkD病历审阅 = 15
    chkD医嘱安排 = 16
    chkD危急值 = 17
    chkD报告撤消 = 18
    chkD医嘱审核 = 19
    chkD处方审查 = 20
    chkD传染病 = 21
    chkD病历质控 = 22
    chkD备血完成 = 23
    chkD校对疑问 = 24
    chkD用血审核 = 25
    chkD输血反应 = 26
End Enum

Public Sub ShowMe()
    '由新版住院护士工作站调用，显示标注按钮
    Me.Show vbModal
End Sub


Private Sub chkNotifyAdvice_Click()
    txtNotifyAdvice.Enabled = chkNotifyAdvice.Value = 1
    If Visible And txtNotifyAdvice.Enabled Then txtNotifyAdvice.SetFocus
End Sub

Private Sub chkNotifyEPR_Click()
    txtNotifyEPR.Enabled = chkNotifyEPR.Value = 1
    If Visible And txtNotifyEPR.Enabled Then txtNotifyEPR.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date
    Dim strTmp As String
    Dim i As Integer
    Dim blnSetup As Boolean
    
    If mbln护士站 Then
        If chkNotifyAdvice.Value = 1 And Val(txtNotifyAdvice.Text) = 0 Then
            If txtNotifyAdvice.Text = "" Then
                MsgBox "请设置医嘱提醒的自动刷新间隔。", vbInformation, gstrSysName
            Else
                MsgBox "医嘱提醒的自动刷新间隔至少应为1分钟。", vbInformation, gstrSysName
            End If
            txtNotifyAdvice.SetFocus: Exit Sub
        End If
        If Val(txtNotifyAdviceDay.Text) = 0 Then
            If txtNotifyAdviceDay.Text = "" Then
                MsgBox "请设置要提醒的医嘱天数。", vbInformation, gstrSysName
            Else
                MsgBox "要提醒的医嘱天数至少应为1天。", vbInformation, gstrSysName
            End If
            txtNotifyAdviceDay.SetFocus: Exit Sub
        End If
    Else
        If chkNotifyEPR.Value = 1 And Val(txtNotifyEPR.Text) = 0 Then
            If txtNotifyEPR.Text = "" Then
                MsgBox "请设置病历审阅提醒的自动刷新间隔。", vbInformation, gstrSysName
            Else
                MsgBox "病历审阅提醒的自动刷新间隔至少应为1分钟。", vbInformation, gstrSysName
            End If
            txtNotifyEPR.SetFocus: Exit Sub
        End If
        
        If Val(txtNotifyEPRDay.Text) = 0 Then
            If txtNotifyEPRDay.Text = "" Then
                MsgBox "请设置要提醒审阅的病历完成天数。", vbInformation, gstrSysName
            Else
                MsgBox "要提醒审阅的病历完成天数至少应为1天。", vbInformation, gstrSysName
            End If
            txtNotifyEPRDay.SetFocus: Exit Sub
        End If
    End If
    

    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0

    '自动刷新医嘱提醒
    If mbln护士站 Then
        Call zlDatabase.SetPara("自动刷新医嘱间隔", IIf(chkNotifyAdvice.Value = 1, Val(txtNotifyAdvice.Text), ""), glngSys, p住院护士站, blnSetup)
        Call zlDatabase.SetPara("自动刷新医嘱天数", Val(txtNotifyAdviceDay.Text), glngSys, p住院护士站, blnSetup)
        strTmp = ""
        For i = chkN新开 To chk血袋回收
            strTmp = strTmp & chkWarn(i).Value
        Next
        Call zlDatabase.SetPara("自动刷新医嘱类型", strTmp, glngSys, p住院护士站, blnSetup)
        Call zlDatabase.SetPara("医嘱处理范围", IIf(optBaby(0).Value, 0, IIf(optBaby(1).Value, 1, 2)), glngSys, p住院医嘱发送, blnSetup)
        Call zlDatabase.SetPara("启用语音提示", chkSoundHS.Value, glngSys, p住院护士站, blnSetup)
    Else
        Call zlDatabase.SetPara("自动刷新病历审阅间隔", IIf(chkNotifyEPR.Value = 1, Val(txtNotifyEPR.Text), ""), glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("自动刷新病历审阅天数", Val(txtNotifyEPRDay.Text), glngSys, p住院医生站, blnSetup)
        strTmp = ""
        For i = chkD病历审阅 To chkD输血反应
            strTmp = strTmp & chkWarn(i).Value
        Next
        Call zlDatabase.SetPara("自动刷新内容", strTmp, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("启用语音提示", chkSoundYS.Value, glngSys, p住院医生站, blnSetup)
    End If
    
    Call zlDatabase.SetPara("住院危急值弹窗提醒", chk危急值.Value, glngSys, p住院医生站, blnSetup)
    
    Call zlDatabase.SetPara("住院发送单据打印", IIf(opt住院诊疗单打印(0).Value, 0, IIf(opt住院诊疗单打印(1).Value, 1, 2)), glngSys, p住院医嘱发送, blnSetup)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdSoundYSSet_Click()
'医生
    Call frmMsgCallSetup.ShowMe(Me, 1)
End Sub

Private Sub cmdSoundHSSet_Click()
'护士
    Call frmMsgCallSetup.ShowMe(Me, 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String, i As Long
    Dim curDate As Date, intDay As Integer
    Dim intType As Integer
    Dim strNotify As String
    
    gblnOK = False
    mlngModual = IIf(mbln护士站, p住院护士站, p住院医生站)
    If mbln护士站 Then
        fraAdvice.Visible = True
        fraBaby.Visible = True
        fraEPR.Visible = False
    Else
        fraAdvice.Visible = False
        fraBaby.Visible = False
        fraEPR.Visible = True
        i = fraBaby.Height + 60
    End If
    Me.Height = Me.Height - i
            
    opt住院诊疗单打印(Val(zlDatabase.GetPara("住院发送单据打印", glngSys, p住院医嘱发送, "0", Array(opt住院诊疗单打印(0), opt住院诊疗单打印(1), opt住院诊疗单打印(2)), InStr(mstrPrivs, "参数设置") > 0))).Value = True
    chkWarn(chk取血通知).Visible = gbln血库系统
    chkWarn(chk备血完成).Visible = gbln血库系统
    chkWarn(chk血袋回收).Visible = gbln血库系统
    
    '危急值弹窗提醒
    chk危急值.Value = Val(zlDatabase.GetPara("住院危急值弹窗提醒", glngSys, p住院医生站, "1", Array(chk危急值), intType))

    '自动刷新医嘱提醒
    If mbln护士站 Then
        strPar = zlDatabase.GetPara("自动刷新医嘱间隔", glngSys, mlngModual, , Array(chkNotifyAdvice), InStr(mstrPrivs, "参数设置") > 0, intType)
        If Val(strPar) > 0 Then
            chkNotifyAdvice.Value = 1: txtNotifyAdvice.Text = Val(strPar)
        End If
        '前面事件中会自动可用，因此后面强制设置
        If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0 Then
            txtNotifyAdvice.Enabled = False
        End If
        
        strPar = zlDatabase.GetPara("自动刷新医嘱天数", glngSys, mlngModual, 1, Array(lblNotifyAdviceDay, txtNotifyAdviceDay), InStr(mstrPrivs, "参数设置") > 0)
        txtNotifyAdviceDay.Text = Val(strPar)
        
        strPar = zlDatabase.GetPara("自动刷新医嘱类型", glngSys, mlngModual, "000000000000", Array(lbl提醒内容, chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), chkWarn(6), chkWarn(7), chkWarn(8), chkWarn(9), chkWarn(10), chkWarn(11), chkWarn(12)), InStr(mstrPrivs, "参数设置") > 0)
        For i = 1 To Len(strPar)
            chkWarn(i - 1).Value = IIf(Val(Mid(strPar, i, 1)) = 1, 1, 0)
        Next
    
        optBaby(Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0", Array(optBaby(0), optBaby(1), optBaby(2)), InStr(mstrPrivs, "参数设置") > 0))).Value = True
        
        chkSoundHS.Value = Val(zlDatabase.GetPara("启用语音提示", glngSys, mlngModual, "1", Array(chkSoundHS, cmdSoundHSSet), InStr(mstrPrivs, "参数设置") > 0))
        
    Else
        strPar = zlDatabase.GetPara("自动刷新病历审阅间隔", glngSys, mlngModual, , Array(chkNotifyEPR), InStr(mstrPrivs, "参数设置") > 0, intType)
        If Val(strPar) > 0 Then
            chkNotifyEPR.Value = 1: txtNotifyEPR.Text = Val(strPar)
        End If
        '前面事件中会自动可用，因此后面强制设置
        If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0 Then
            txtNotifyEPR.Enabled = False
        End If
        
        strPar = zlDatabase.GetPara("自动刷新病历审阅天数", glngSys, mlngModual, 1, Array(lblNotifyEPRDay, txtNotifyEPRDay), InStr(mstrPrivs, "参数设置") > 0)
        txtNotifyEPRDay.Text = Val(strPar)
       
        strNotify = zlDatabase.GetPara("自动刷新内容", glngSys, p住院医生站, , Array(chkWarn(15), chkWarn(16), chkWarn(17), chkWarn(18), chkWarn(19), chkWarn(20), chkWarn(21), chkWarn(22), chkWarn(23), chkWarn(24), chkWarn(25), Label提醒区域), InStr(mstrPrivs, "参数设置") > 0)
        chkWarn(chkD病历审阅).Value = Val(Mid(strNotify, 1, 1))
        chkWarn(chkD医嘱安排).Value = Val(Mid(strNotify, 2, 1))
        chkWarn(chkD危急值).Value = Val(Mid(strNotify, 3, 1))
        chkWarn(chkD报告撤消).Value = Val(Mid(strNotify, 4, 1))
        chkWarn(chkD医嘱审核).Value = Val(Mid(strNotify, 5, 1))
        chkWarn(chkD处方审查).Value = Val(Mid(strNotify, 6, 1))
        chkWarn(chkD传染病).Value = Val(Mid(strNotify, 7, 1))
        chkWarn(chkD病历质控).Value = Val(Mid(strNotify, 8, 1))
        chkWarn(chkD备血完成).Value = Val(Mid(strNotify, 9, 1))
        chkWarn(chkD备血完成).Visible = gbln血库系统
        chkWarn(chkD校对疑问).Value = Val(Mid(strNotify, 10, 1))
        chkWarn(chkD用血审核).Value = Val(Mid(strNotify, 11, 1))
        chkWarn(chkD用血审核).Visible = gbln血库系统
        chkWarn(chkD输血反应).Value = Val(Mid(strNotify, 12, 1))
        chkWarn(chkD输血反应).Visible = gbln血库系统
        
        If InStr(mstrPrivs, "参数设置") = 0 Then
            chkWarn(chkD病历审阅).Enabled = False
            chkWarn(chkD医嘱安排).Enabled = False
            chkWarn(chkD危急值).Enabled = False
            chkWarn(chkD报告撤消).Enabled = False
            chkWarn(chkD医嘱审核).Enabled = False
            chkWarn(chkD处方审查).Enabled = False
            chkWarn(chkD传染病).Enabled = False
            chkWarn(chkD病历质控).Enabled = False
            chkWarn(chkD备血完成).Enabled = False
            chkWarn(chkD校对疑问).Enabled = False
            chkWarn(chkD用血审核).Enabled = False
            chkWarn(chkD输血反应).Enabled = False
        End If
        
        chkSoundYS.Value = Val(zlDatabase.GetPara("启用语音提示", glngSys, mlngModual, "1", Array(chkSoundYS, cmdSoundYSSet), InStr(mstrPrivs, "参数设置") > 0))
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln护士站 = False
End Sub

Private Sub txtNotifyAdvice_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdvice)
End Sub

Private Sub txtNotifyAdvice_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPR_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPR)
End Sub

Private Sub txtNotifyEPR_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdviceDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdviceDay)
End Sub

Private Sub txtNotifyAdviceDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPRDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPRDay)
End Sub

Private Sub txtNotifyEPRDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

