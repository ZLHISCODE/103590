VERSION 5.00
Begin VB.Form frmSetup 
   Caption         =   "参数设置"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   5580
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkOrderStyle 
      Caption         =   "使用数据原始顺序排序"
      Height          =   255
      Left            =   2880
      TabIndex        =   48
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "排队分组显示类型"
      Height          =   975
      Left            =   240
      TabIndex        =   45
      Top             =   4800
      Width           =   5175
      Begin VB.OptionButton optGroupType 
         Caption         =   "按诊室分组"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   49
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optGroupType 
         Caption         =   "按医生姓名分组"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   47
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optGroupType 
         Caption         =   "按队列名称分组"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame framCalledColumn 
      Caption         =   "已呼叫列设置"
      Height          =   1095
      Left            =   240
      TabIndex        =   37
      Top             =   7080
      Width           =   5175
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "医生姓名"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   44
         Tag             =   "医生姓名"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "诊室"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   43
         Tag             =   "诊室"
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "号码"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Tag             =   "号码"
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "患者姓名"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   41
         Tag             =   "患者姓名"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "呼叫人"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   40
         Tag             =   "呼叫医生"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "呼叫时间"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   39
         Tag             =   "呼叫时间"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "回诊序号"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Tag             =   "回诊序号"
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.ComboBox cbxComeback 
      Height          =   300
      ItemData        =   "frmSetup.frx":06EA
      Left            =   960
      List            =   "frmSetup.frx":06F4
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   8235
      Width           =   975
   End
   Begin VB.Frame framColumn 
      Caption         =   "排队列设置"
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   5880
      Width           =   5175
      Begin VB.CheckBox chkColumn 
         Caption         =   "回诊序号"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   27
         Tag             =   "回诊序号"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "排队时间"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   25
         Tag             =   "排队时间"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "排队状态"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   24
         Tag             =   "排队状态"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "医生姓名"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Tag             =   "医生姓名"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "诊室"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   22
         Tag             =   "诊室"
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "优先"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   21
         Tag             =   "优先"
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "患者姓名"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   20
         Tag             =   "患者姓名"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "号码"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Tag             =   "号码"
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.CheckBox chkUseDisplay 
      Caption         =   "显示排队队列"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "叫号方式设置"
      Height          =   3375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   5175
      Begin VB.OptionButton optCallWay 
         Caption         =   "启用远端语音"
         Height          =   450
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox chkUseSound 
         Caption         =   "启用语音呼叫"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optCallWay 
         Caption         =   "启用本地语音"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Frame frm语音广播设置 
         Height          =   1935
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txtLoopQueryTime 
            Height          =   270
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   34
            Text            =   "30"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtSpeed 
            Height          =   270
            Left            =   1320
            TabIndex        =   32
            Text            =   "6"
            Top             =   1200
            Width           =   495
         End
         Begin VB.ComboBox cboSoundType 
            Height          =   300
            ItemData        =   "frmSetup.frx":0706
            Left            =   2760
            List            =   "frmSetup.frx":0710
            TabIndex        =   31
            Text            =   "cboSoundType"
            Top             =   340
            Width           =   1815
         End
         Begin VB.TextBox txtPlayCount 
            Height          =   270
            Left            =   3720
            TabIndex        =   16
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txt广播时间长度 
            Height          =   270
            Left            =   1800
            TabIndex        =   13
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "数据轮询间隔时间为"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1605
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "秒"
            Height          =   255
            Left            =   2480
            TabIndex        =   35
            Top             =   1605
            Width           =   255
         End
         Begin VB.Label Label6 
            Caption         =   "(语速范围在-10到10之间) "
            Height          =   255
            Left            =   1800
            TabIndex        =   33
            Top             =   1230
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "语音类型："
            Height          =   255
            Left            =   1920
            TabIndex        =   30
            Top             =   380
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "每段语音广播长度为        秒 播放次数为        次"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   855
            Width           =   4455
         End
         Begin VB.Label Label3 
            Caption         =   "语音广播语速："
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1215
            Width           =   1755
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   4815
         Begin VB.ComboBox cboWorkStation 
            Height          =   300
            Left            =   1320
            TabIndex        =   26
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label labRemoteComputerName 
            Caption         =   "远端站点名："
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   400
            Width           =   1215
         End
      End
   End
   Begin VB.Frame frm显示设备设置 
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cbo显示硬件类别 
         Height          =   300
         ItemData        =   "frmSetup.frx":0728
         Left            =   240
         List            =   "frmSetup.frx":072A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton cmd显示设备设置 
         Caption         =   "设备设置"
         Height          =   300
         Left            =   3840
         TabIndex        =   8
         Top             =   600
         Width           =   1100
      End
      Begin VB.Label Label2 
         Caption         =   "显示设备类别："
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   8640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   8640
      Width           =   1100
   End
   Begin VB.Label labCallBack 
      Caption         =   "回诊病人           排队"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   8280
      Width           =   2175
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrReg As String


Private Sub cbo语速_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub chkUseDisplay_Click()
    If chkUseDisplay.Value = 0 Then
        frm显示设备设置.Enabled = False
        
        cbo显示硬件类别.BackColor = frm显示设备设置.BackColor
    Else
        frm显示设备设置.Enabled = True
        
        cbo显示硬件类别.BackColor = &H80000005
        
        
    End If
End Sub

Private Sub cmdCancel_Click()
    '关闭窗口
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '保存参数设置
    Dim strRet As String
    Dim strColumnInf As String
    Dim strCalledColumnInf As String
    
    Dim i As Integer
    
    mstrReg = "公共全局\排队叫号"
    
    If Val(txtLoopQueryTime.Text) > 65 Then
        MsgBox "数据轮询间隔时间不能大于65秒，请重新设置。", vbOKOnly, Me.Caption
        
        txtLoopQueryTime.SetFocus
        Call zlControl.TxtSelAll(txtLoopQueryTime)
        
        Exit Sub
    End If
    
    'SaveSetting "ZLSOFT", strReg, "语音广播时间长度", Val(txt广播时间长度.Text)
    Call zlDatabase.SetPara("语音广播时间长度", Val(txt广播时间长度.Text), glngSys, glngModul)
    'SaveSetting "ZLSOFT", strReg, "语音广播语速", Val(cbo语速.Text)
    Call zlDatabase.SetPara("语音广播语速", Val(txtSpeed.Text), glngSys, glngModul)
    'SaveSetting "ZLSOFT", strReg, "显示排队队列", chkUseDisplay.Value
    Call zlDatabase.SetPara("显示排队队列", chkUseDisplay.Value, glngSys, glngModul)
    'SaveSetting "ZLSOFT", strReg, "启用语音呼叫", chkUseSound.Value
    Call zlDatabase.SetPara("启用语音呼叫", chkUseSound.Value, glngSys, glngModul)
    
    'SaveSetting "ZLSOFT", strReg, "远端站点名称", txtRemoteComputerName.Text
    Call zlDatabase.SetPara("远端呼叫站点", cboWorkStation.Text, glngSys, glngModul)
    'SaveSetting "ZLSOFT", strReg, "语音播放次数", txtPlayCount.Text
    Call zlDatabase.SetPara("语音播放次数", Val(txtPlayCount.Text), glngSys, glngModul)
    
    '语音类型
    Call zlDatabase.SetPara("语音类型", cboSoundType.Text, glngSys, glngModul)
    '轮询时间
    Call zlDatabase.SetPara("轮询时间", Val(txtLoopQueryTime.Text), glngSys, glngModul)
    
    strColumnInf = ""
    For i = 0 To 7
        If chkColumn(i).Value = vbChecked Then
            If Trim(strColumnInf) <> "" Then strColumnInf = strColumnInf & ","
            strColumnInf = strColumnInf & chkColumn(i).Tag
        End If
    Next i
    
    Call zlDatabase.SetPara("数据显示列", strColumnInf, glngSys, glngModul)
    
    
    
    strCalledColumnInf = ""
    For i = 0 To 6
        If chkCalledColumn(i).Value = vbChecked Then
            If Trim(strCalledColumnInf) <> "" Then strCalledColumnInf = strCalledColumnInf & ","
            strCalledColumnInf = strCalledColumnInf & chkCalledColumn(i).Tag
        End If
    Next i
    
    Call zlDatabase.SetPara("呼叫数据显示列", strCalledColumnInf, glngSys, glngModul)
    
    
    '保存叫号方式
    If optCallWay(0).Value Then
        'SaveSetting "ZLSOFT", strReg, "叫号方式", 1
         Call zlDatabase.SetPara("叫号方式", 1, glngSys, glngModul)
    Else
        'SaveSetting "ZLSOFT", strReg, "叫号方式", 0
        Call zlDatabase.SetPara("叫号方式", 0, glngSys, glngModul)
    End If
    
    
    '保存显示设备
    If cbo显示硬件类别.ListIndex <> -1 Then
        'SaveSetting "ZLSOFT", strReg, "显示设备类别", cbo显示硬件类别.ItemData(cbo显示硬件类别.ListIndex)
        Call zlDatabase.SetPara("显示设备类别", cbo显示硬件类别.ItemData(cbo显示硬件类别.ListIndex), glngSys, glngModul)
    End If
    
    Call zlDatabase.SetPara("回诊病人是否优先", cbxComeback.ListIndex, glngSys, glngModul)
    
    For i = 0 To optGroupType.Count - 1
        If optGroupType(i).Value Then
            Call zlDatabase.SetPara("排队分组类型", i, glngSys, glngModul)
            Exit For
        End If
    Next
    
    Call zlDatabase.SetPara("使用数据原始顺序排序", chkOrderStyle.Value, glngSys, glngModul)
    '关闭窗口
    Unload Me
End Sub

Private Sub cmd显示设备设置_Click()
    If pobjLEDShow Is Nothing Then
        Call frmQueueStation.InitLED(plngLEDModal)
    End If
        
    If Not pobjLEDShow Is Nothing Then
        Call pobjLEDShow.zlSetup(Me)
    End If
End Sub

Private Sub ReadWorkStationInf()
'*****************************************************
'读取站点信息
'*****************************************************

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select 工作站 from zlClients where 禁止使用<>1 order by 工作站"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "读取站点信息")
    
    If rsTemp.EOF Then Exit Sub
    
    While Not rsTemp.EOF
        Call cboWorkStation.AddItem(rsTemp("工作站"))
        rsTemp.MoveNext
    Wend
    
End Sub

Private Sub ReadLocalPara()
    Dim lng广播语速 As Long
    Dim lngLEDModal As Long
    Dim strColumnInf As String
    Dim strCalledColumnInf As String
    
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
   '读取叫号方式
        
    'optCallWay(0).Value = Val(GetSetting("ZLSOFT", strReg, "叫号方式", 0))
    optCallWay(0).Value = Val(zlDatabase.GetPara("叫号方式", glngSys, glngModul, "0"))
    optCallWay(1).Value = Not optCallWay(0).Value
    
    'txtRemoteComputerName.Text = GetSetting("ZLSOFT", strReg, "远端站点名称", "")
    cboWorkStation.Text = zlDatabase.GetPara("远端呼叫站点", glngSys, glngModul, "")
    cboWorkStation.Enabled = optCallWay(0).Value
    
    
    txtLoopQueryTime.Text = Val(zlDatabase.GetPara("轮询时间", glngSys, glngModul, "30"))
    'txt广播时间长度.Text = Val(GetSetting("ZLSOFT", strReg, "语音广播时间长度", 15))
    txt广播时间长度.Text = Val(zlDatabase.GetPara("语音广播时间长度", glngSys, glngModul, "15"))
    'txtPlayCount.Text = Val(GetSetting("ZLSOFT", strReg, "语音播放次数", 3))
    txtPlayCount.Text = Val(zlDatabase.GetPara("语音播放次数", glngSys, glngModul, "3"))
    'lng广播语速 = Val(GetSetting("ZLSOFT", strReg, "语音广播语速", 60))
    lng广播语速 = Val(zlDatabase.GetPara("语音广播语速", glngSys, glngModul, "60"))
    
    cboSoundType.Text = zlDatabase.GetPara("语音类型", glngSys, glngModul, "系统默认")
    
    strColumnInf = zlDatabase.GetPara("数据显示列", glngSys, glngModul, ",号码,患者姓名,排队状态,")
    strColumnInf = Replace(strColumnInf, "，", ",")
    strColumnInf = "," & strColumnInf & ","
    
    For i = 0 To 7
        chkColumn(i).Value = Int(IIf(InStr(1, strColumnInf, "," & chkColumn(i).Tag & ",") > 0, vbChecked, vbUnchecked))
    Next i
    
    
    
    strCalledColumnInf = zlDatabase.GetPara("呼叫数据显示列", glngSys, glngModul, ",号码,患者姓名,")
    strCalledColumnInf = Replace(strCalledColumnInf, "，", ",")
    strCalledColumnInf = "," & strCalledColumnInf & ","
    
    For i = 0 To 6
        chkCalledColumn(i).Value = Int(IIf(InStr(1, strCalledColumnInf, "," & chkCalledColumn(i).Tag & ",") > 0, vbChecked, vbUnchecked))
    Next i
    
    If optCallWay(0).Value Then
        txt广播时间长度.BackColor = Me.BackColor
        txtPlayCount.BackColor = Me.BackColor
        txtSpeed.BackColor = Me.BackColor
        
        cboWorkStation.BackColor = &H80000005
        
        frm语音广播设置.Enabled = False
    Else
        txt广播时间长度.BackColor = &H80000005
        txtPlayCount.BackColor = &H80000005
        txtSpeed.BackColor = &H80000005
        
        cboWorkStation.BackColor = Me.BackColor
        
        frm语音广播设置.Enabled = True
    End If
    
    
    If lng广播语速 <= 10 And lng广播语速 >= -10 Then
        txtSpeed.Text = lng广播语速
    Else
        txtSpeed.Text = 0
    End If
    
    'chkUseSound.Value = GetSetting("ZLSOFT", strReg, "启用语音呼叫", 1)
    chkUseSound.Value = zlDatabase.GetPara("启用语音呼叫", glngSys, glngModul, "1")
    
    'chkUseDisplay.Value = GetSetting("ZLSOFT", strReg, "显示排队队列", 1)
    chkUseDisplay.Value = zlDatabase.GetPara("显示排队队列", glngSys, glngModul, "1")
    If chkUseDisplay.Value = 0 Then
        cbo显示硬件类别.BackColor = frm显示设备设置.BackColor
    End If
    
    '填写显示设备类别
    'lngLEDModal = GetSetting("ZLSOFT", strReg, "显示设备类别", 101)
    lngLEDModal = zlDatabase.GetPara("显示设备类别", glngSys, glngModul, "101")
    
    cbo显示硬件类别.Clear
    
    strSql = "Select 部件类型,部件名,Nvl(启用,0) AS 启用,说明 From 排队LED显示部件  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取该LED显示接口的注册信息")
    
    While rsTemp.EOF = False
        cbo显示硬件类别.AddItem Nvl(rsTemp!说明)
        cbo显示硬件类别.ItemData(cbo显示硬件类别.ListCount - 1) = Nvl(rsTemp!部件类型, 0)
        If lngLEDModal = Nvl(rsTemp!部件类型, 0) Then
            cbo显示硬件类别.ListIndex = cbo显示硬件类别.ListCount - 1
        End If
        rsTemp.MoveNext
    Wend
    
    If cbo显示硬件类别.ListCount > 0 And cbo显示硬件类别.ListIndex = -1 Then
        cbo显示硬件类别.ListIndex = 0
    End If
    
    cbxComeback.ListIndex = zlDatabase.GetPara("回诊病人是否优先", glngSys, glngModul, "1", Array(labCallBack, cbxComeback), True)
    
    optGroupType(Val(zlDatabase.GetPara("排队分组类型", glngSys, glngModul, "0"))).Value = True
    
    chkOrderStyle.Value = zlDatabase.GetPara("使用数据原始顺序排序", glngSys, glngModul, "0")
End Sub

Private Sub Form_Load()
    Call ReadWorkStationInf
    
    Call ReadLocalPara
End Sub


Private Sub optCallWay_Click(Index As Integer)
    cboWorkStation.Enabled = optCallWay(0).Value
    
    If optCallWay(0).Value Then
        frm语音广播设置.Enabled = False
        
        txt广播时间长度.BackColor = Me.BackColor
        txtPlayCount.BackColor = Me.BackColor
        txtSpeed.BackColor = Me.BackColor
        
        cboWorkStation.BackColor = &H80000005
    Else
        frm语音广播设置.Enabled = True
        
        txt广播时间长度.BackColor = &H80000005
        txtPlayCount.BackColor = &H80000005
        txtSpeed.BackColor = &H80000005
        
        cboWorkStation.BackColor = Me.BackColor
    End If
End Sub

Private Sub txt广播时间长度_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
