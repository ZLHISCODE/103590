VERSION 5.00
Begin VB.Form frmTechnicSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "frmTechnicSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraAutoTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   350
      TabIndex        =   27
      Top             =   550
      Width           =   465
   End
   Begin VB.TextBox txtAutoTime 
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
      Left            =   320
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   465
   End
   Begin VB.Frame fraNotify 
      Caption         =   "提醒设置"
      Height          =   1230
      Left            =   105
      TabIndex        =   13
      Top             =   3870
      Width           =   6270
      Begin VB.CheckBox chkWarn 
         Caption         =   "血袋回收"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   28
         Top             =   885
         Width           =   1020
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "待安排"
         Height          =   195
         Index           =   1
         Left            =   2500
         TabIndex        =   14
         Top             =   885
         Width           =   900
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "语音设置(&S)"
         Height          =   350
         Left            =   4575
         TabIndex        =   25
         Top             =   630
         Width           =   1410
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "启用语音提示"
         Height          =   195
         Left            =   4575
         TabIndex        =   24
         Top             =   360
         Width           =   1470
      End
      Begin VB.Frame fraLinM 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   23
         Top             =   525
         Width           =   300
      End
      Begin VB.TextBox txtMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "10"
         Top             =   360
         Width           =   300
      End
      Begin VB.CheckBox chkNotify 
         Caption         =   "每    分钟自动刷新提醒区域中的内容"
         Height          =   195
         Left            =   345
         TabIndex        =   20
         Top             =   345
         Width           =   3390
      End
      Begin VB.Frame fraNotifyEPR 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   18
         Top             =   510
         Width           =   300
      End
      Begin VB.Frame fraLinD 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   17
         Top             =   780
         Width           =   300
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "1"
         Top             =   600
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "销帐申请"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   15
         Top             =   885
         Width           =   1065
      End
      Begin VB.Label lblNotifyArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提醒内容:"
         Height          =   180
         Left            =   600
         TabIndex        =   22
         Top             =   880
         Width           =   810
      End
      Begin VB.Label lblNotifyDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "将    天内完成的内容显示在提醒区域"
         Height          =   180
         Left            =   615
         TabIndex        =   21
         Top             =   615
         Width           =   3060
      End
   End
   Begin VB.ListBox lst治疗类别 
      Columns         =   3
      ForeColor       =   &H80000012&
      Height          =   1110
      IMEMode         =   3  'DISABLE
      Left            =   2760
      Style           =   1  'Checkbox
      TabIndex        =   5
      ToolTipText     =   "按Ctrl+A全选，按Ctrl+C全清"
      Top             =   2610
      Width           =   3615
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   6525
      TabIndex        =   9
      Top             =   5115
      Width           =   6525
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5310
         TabIndex        =   7
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4215
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1500
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   7080
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   7080
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.ListBox lst诊疗类别 
      Columns         =   3
      ForeColor       =   &H80000012&
      Height          =   1740
      IMEMode         =   3  'DISABLE
      Left            =   2760
      Style           =   1  'Checkbox
      TabIndex        =   4
      ToolTipText     =   "按Ctrl+A全选，按Ctrl+C全清"
      Top             =   450
      Width           =   3615
   End
   Begin VB.CheckBox chkExeLog 
      Caption         =   "严格要求记录执行的情况"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   675
      Width           =   2280
   End
   Begin VB.CheckBox chkRoom 
      Caption         =   "只显示指定的执行间范围"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2280
   End
   Begin VB.Frame fraRoom 
      Caption         =   " 执行间 "
      Height          =   2520
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2445
      Begin VB.ListBox lstRoom 
         Enabled         =   0   'False
         Height          =   2160
         ItemData        =   "frmTechnicSetup.frx":000C
         Left            =   120
         List            =   "frmTechnicSetup.frx":000E
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label lblAutoTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "每      秒自动刷新病人清单"
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   375
      Width           =   2340
   End
   Begin VB.Label lbl治疗类别 
      Caption         =   "治疗类别"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   2355
      Width           =   1215
   End
   Begin VB.Label lbl诊疗类别 
      Caption         =   "单据过滤类别"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmTechnicSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mstrPrivs As String
Public mlng科室ID As Long 'IN:当前执行科室ID
Public mblnOK As Boolean

Private Sub chkRoom_Click()
    lstRoom.Enabled = chkRoom.Value = 1 And lstRoom.Tag = ""
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call ZLCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long, k As Long, bln治疗 As Boolean
    Dim blnSetup As Boolean
    
    '执行间范围
    strPar = ""
    If chkRoom.Value = 1 Then
        For i = 0 To lstRoom.ListCount - 1
            If lstRoom.Selected(i) Then
                strPar = strPar & "|" & lstRoom.List(i)
            End If
        Next
        If strPar = "" Then
            MsgBox "请至少选择一个执行间。", vbInformation, gstrSysName
            lstRoom.SetFocus: Exit Sub
        End If
    End If
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    Call zlDatabase.SetPara("执行间范围", Replace(Mid(strPar, 2), "'", "''"), glngSys, p医技工作站, blnSetup)
        
 
    '严格要求记录执行的情况
    Call zlDatabase.SetPara("记录执行情况", chkExeLog.Value, glngSys, p医技工作站, blnSetup)
    
    '诊疗类别
    k = 0
    strPar = ""
    For i = 0 To lst诊疗类别.ListCount - 1
        If lst诊疗类别.Selected(i) Then
            strPar = strPar & Chr(lst诊疗类别.ItemData(i))
            If Chr(lst诊疗类别.ItemData(i)) = "E" Then bln治疗 = True
            k = k + 1
        End If
    Next
    If strPar = "" Then
        MsgBox "请至少选择一种要执行的诊疗类别。", vbInformation, gstrSysName
        lst诊疗类别.SetFocus: Exit Sub
    End If
    If k = lst诊疗类别.ListCount Then strPar = ""
    Call zlDatabase.SetPara("诊疗类别", strPar, glngSys, p医技工作站, blnSetup)
    
    '治疗类别
    If bln治疗 Then
        k = 0
        strPar = ""
        For i = 0 To lst治疗类别.ListCount - 1
            If lst治疗类别.Selected(i) Then
                strPar = strPar & "," & lst治疗类别.ItemData(i)
                k = k + 1
            End If
        Next
        If strPar = "" Then
            MsgBox "请至少选择一种要执行的治疗类别。", vbInformation, gstrSysName
            lst治疗类别.SetFocus: Exit Sub
        Else
            strPar = Mid(strPar, 2)
        End If
        If k = lst治疗类别.ListCount Then strPar = ""
        Call zlDatabase.SetPara("治疗类别", strPar, glngSys, p医技工作站, blnSetup)
    End If
    
    Call zlDatabase.SetPara("自动刷新医嘱间隔", IIf(chkNotify.Value = 1, Val(txtMin.Text), ""), glngSys, p医技工作站, blnSetup)
    Call zlDatabase.SetPara("自动刷新医嘱天数", Val(txtDay.Text), glngSys, p医技工作站, blnSetup)
    Call zlDatabase.SetPara("自动刷新医嘱类型", "" & chkWarn(0).Value & chkWarn(1).Value & chkWarn(2).Value, glngSys, p医技工作站, blnSetup)
    Call zlDatabase.SetPara("启用语音提示", chkSound.Value, glngSys, p医技工作站, blnSetup)
    Call zlDatabase.SetPara("医技刷新间隔", Val(txtAutoTime.Text), glngSys, p医技工作站, blnSetup)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 3)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        If KeyCode = vbKeyA Then
            SelAll诊疗类别 (True)
        ElseIf KeyCode = vbKeyC Then
            SelAll诊疗类别 (False)
        End If
    End If
End Sub

Private Sub SelAll诊疗类别(ByVal blnSel As Boolean)
    Dim i As Long
    
    For i = 0 To lst诊疗类别.ListCount - 1
        lst诊疗类别.Selected(i) = blnSel
    Next
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strPar As String
    Dim blnSetup As Boolean, arrtmp As Variant, i As Long, bln治疗 As Boolean
    Dim intType As Integer
    
    mblnOK = False
    
    blnSetup = InStr(mstrPrivs, "参数设置") > 0
    
    '严格要求记录执行的情况
    chkExeLog.Value = Val(zlDatabase.GetPara("记录执行情况", glngSys, p医技工作站, "0", Array(chkExeLog), blnSetup))
 
    '执行房间
    strPar = zlDatabase.GetPara("执行间范围", glngSys, p医技工作站, "", Array(chkRoom, fraRoom, lstRoom), blnSetup)
    If Not chkRoom.Enabled Then lstRoom.Tag = "1" '固定标记为不可用
    chkRoom.Value = IIf(strPar = "", 0, 1)
    strSql = "Select 执行间 From 医技执行房间 Where 科室ID=[1]"
    On Error GoTo Errh
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng科室ID)
    Do While Not rsTmp.EOF
        lstRoom.AddItem rsTmp!执行间
        If InStr("|" & strPar & "|", "|" & rsTmp!执行间 & "|") > 0 Then
            lstRoom.Selected(lstRoom.NewIndex) = True
        End If
        rsTmp.MoveNext
    Loop
    If lstRoom.ListCount > 0 Then
        lstRoom.TopIndex = 0
        lstRoom.ListIndex = 0
    ElseIf blnSetup Then
        chkRoom.Value = 0
        chkRoom.Enabled = False
    End If
    
    
    '诊疗类别
    strPar = zlDatabase.GetPara("诊疗类别", glngSys, p医技工作站, , Array(lst诊疗类别), blnSetup)
        
    strSql = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('5','6','7','8','9') Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    With lst诊疗类别
        Do While Not rsTmp.EOF
            .AddItem rsTmp!编码 & "-" & rsTmp!名称
            .ItemData(.NewIndex) = Asc(rsTmp!编码)
            
            If strPar <> "" Then
                If InStr(strPar, rsTmp!编码) > 0 Then
                    .Selected(.NewIndex) = True
                    If rsTmp!编码 = "E" Then bln治疗 = True
                End If
            Else
                .Selected(.NewIndex) = True
                If Not bln治疗 Then bln治疗 = True
            End If
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    
    strPar = "0-普通;1-过敏试验;2-给药方法;3-中药煎法;4-中药用法;5-特殊治疗;6-采集方法;7-配血方法;8-输血途径"
    arrtmp = Split(strPar, ";")
    
    strPar = zlDatabase.GetPara("治疗类别", glngSys, p医技工作站, , Array(lst治疗类别), blnSetup)
    If strPar <> "" Then
        strPar = "," & strPar & ","
    End If
    With lst治疗类别
        For i = 0 To UBound(arrtmp)
            .AddItem arrtmp(i)
            .ItemData(.NewIndex) = Val(arrtmp(i))
            
            If strPar <> "" Then
                If InStr(strPar, "," & Val(arrtmp(i)) & ",") > 0 Then
                    .Selected(.NewIndex) = True
                End If
            Else
                .Selected(.NewIndex) = True
            End If
        Next
    End With
    lst治疗类别.Enabled = bln治疗
    
    strPar = zlDatabase.GetPara("自动刷新医嘱间隔", glngSys, p医技工作站, , Array(chkNotify), InStr(mstrPrivs, "参数设置") > 0, intType)
    If Val(strPar) > 0 Then chkNotify.Value = 1: txtMin.Text = Val(strPar)
    '前面事件中会自动可用，因此后面强制设置
    If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0 Then
        txtMin.Enabled = False
    End If
    
    strPar = zlDatabase.GetPara("自动刷新医嘱天数", glngSys, p医技工作站, 1, Array(lblNotifyDay, txtDay), InStr(mstrPrivs, "参数设置") > 0)
    txtDay.Text = Val(strPar)
    
    strPar = zlDatabase.GetPara("医技刷新间隔", glngSys, p医技工作站, 1, Array(lblAutoTime, txtAutoTime), InStr(mstrPrivs, "参数设置") > 0)
    txtAutoTime.Text = Val(strPar)
    
    strPar = zlDatabase.GetPara("自动刷新医嘱类型", glngSys, p医技工作站, "000", Array(lblNotifyArea, chkWarn(0), chkWarn(1), chkWarn(2)), InStr(mstrPrivs, "参数设置") > 0)
    chkWarn(2).Visible = gbln血库系统
    For i = 1 To chkWarn.Count
        chkWarn(i - 1).Value = IIf(Val(Mid(strPar, i, 1)) = 1, 1, 0)
    Next
    chkSound.Value = Val(zlDatabase.GetPara("启用语音提示", glngSys, p医技工作站, "1", Array(chkSound, cmdSoundSet), blnSetup))
    Exit Sub
Errh:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkNotify_Click()
    txtMin.Enabled = chkNotify.Value = 1
    If Visible And txtMin.Enabled Then txtMin.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng科室ID = 0
    mstrPrivs = ""
End Sub

Private Sub lst诊疗类别_ItemCheck(Item As Integer)
    If Chr(lst诊疗类别.ItemData(Item)) = "E" Then
        lst治疗类别.Enabled = lst诊疗类别.Selected(Item)
    End If
End Sub


'限制输入只能为数字
Private Sub txtAutoTime_KeyPress(KeyAscii As Integer)
    If Not KeyAscii = 8 Then
        If InStr("1234567890", Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
