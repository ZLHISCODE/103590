VERSION 5.00
Begin VB.Form frmWork_QueueCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7575
   Icon            =   "frmWork_QueueCfg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chkQueueQuick 
      Caption         =   "自动弹出快捷呼叫窗口"
      Height          =   180
      Left            =   240
      TabIndex        =   36
      Top             =   3960
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkLockAfterCall 
      Caption         =   "呼叫后锁定采集"
      Height          =   180
      Left            =   3240
      TabIndex        =   35
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CheckBox chkShowMySelfCalled 
      Caption         =   "只显示自己呼叫的队列"
      Height          =   180
      Left            =   240
      TabIndex        =   34
      Top             =   3540
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "打印机设置(&P)"
      Height          =   375
      Left            =   1635
      Picture         =   "frmWork_QueueCfg.frx":1042
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1440
   End
   Begin VB.Frame framColumn 
      Caption         =   "排队列设置"
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   1005
      Width           =   7095
      Begin VB.CheckBox chkColumn 
         Caption         =   "医嘱内容"
         Height          =   255
         Index           =   6
         Left            =   1305
         TabIndex        =   28
         Tag             =   "医嘱内容"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "备注"
         Height          =   255
         Index           =   9
         Left            =   5505
         TabIndex        =   27
         Tag             =   "备注"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "排队号码"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Tag             =   "排队号码"
         Top             =   375
         Value           =   1  'Checked
         Width           =   1110
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "患者姓名"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   25
         Tag             =   "患者姓名"
         Top             =   375
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "性别"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   24
         Tag             =   "性别"
         Top             =   375
         Width           =   855
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "诊室(执行间)"
         Height          =   255
         Index           =   4
         Left            =   5505
         TabIndex        =   23
         Tag             =   "诊室"
         Top             =   375
         Width           =   1440
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "检查项目"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Tag             =   "检查项目"
         Top             =   720
         Width           =   1065
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "当前状态"
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   21
         Tag             =   "排队状态"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "排队时间"
         Height          =   255
         Index           =   8
         Left            =   4020
         TabIndex        =   20
         Tag             =   "排队时间"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "年龄"
         Height          =   255
         Index           =   3
         Left            =   4020
         TabIndex        =   19
         Tag             =   "年龄"
         Top             =   375
         Width           =   1095
      End
   End
   Begin VB.Frame framCalledColumn 
      Caption         =   "呼叫列设置"
      Height          =   1080
      Left            =   225
      TabIndex        =   10
      Top             =   2175
      Width           =   7110
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "备注"
         Height          =   255
         Index           =   10
         Left            =   5790
         TabIndex        =   32
         Tag             =   "备注"
         Top             =   705
         Width           =   705
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "当前状态"
         Height          =   255
         Index           =   9
         Left            =   4320
         TabIndex        =   31
         Tag             =   "排队状态"
         Top             =   705
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "医嘱内容"
         Height          =   255
         Index           =   6
         Left            =   105
         TabIndex        =   30
         Tag             =   "医嘱内容"
         Top             =   705
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "诊室(执行间)"
         Height          =   255
         Index           =   4
         Left            =   4290
         TabIndex        =   29
         Tag             =   "诊室"
         Top             =   360
         Width           =   1425
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "检查项目"
         Height          =   255
         Index           =   5
         Left            =   5805
         TabIndex        =   17
         Tag             =   "检查项目"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "呼叫时间"
         Height          =   255
         Index           =   8
         Left            =   2835
         TabIndex        =   16
         Tag             =   "呼叫时间"
         Top             =   705
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "呼叫人"
         Height          =   255
         Index           =   7
         Left            =   1575
         TabIndex        =   15
         Tag             =   "呼叫医生"
         Top             =   705
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "患者姓名"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1305
         TabIndex        =   14
         Tag             =   "患者姓名"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "排队号码"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Tag             =   "排队号码"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "性别"
         Height          =   255
         Index           =   2
         Left            =   2610
         TabIndex        =   12
         Tag             =   "性别"
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkCalledColumn 
         Caption         =   "年龄"
         Height          =   255
         Index           =   3
         Left            =   3435
         TabIndex        =   11
         Tag             =   "年龄"
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdVoiceCfg 
      Caption         =   "语音设置(&V)"
      Height          =   375
      Left            =   225
      Picture         =   "frmWork_QueueCfg.frx":118C
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1275
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   375
      Left            =   5115
      Picture         =   "frmWork_QueueCfg.frx":12D6
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   375
      Left            =   6210
      Picture         =   "frmWork_QueueCfg.frx":1420
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1100
   End
   Begin VB.ComboBox cbxTurnPage 
      Height          =   300
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3480
      Width           =   2580
   End
   Begin VB.Frame frmRoomCfg 
      Caption         =   "本机执行间设置"
      Height          =   765
      Left            =   270
      TabIndex        =   0
      Top             =   165
      Width           =   7050
      Begin VB.ComboBox cbxRoomName 
         Height          =   300
         Left            =   4635
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2325
      End
      Begin VB.ComboBox cbxDept 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   2310
      End
      Begin VB.Label Label2 
         Caption         =   "执行间名称："
         Height          =   195
         Left            =   3555
         TabIndex        =   2
         Top             =   345
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "所属科室："
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   345
         Width           =   900
      End
   End
   Begin VB.Label Label3 
      Caption         =   "接诊后跳转页面："
      Height          =   240
      Left            =   3225
      TabIndex        =   5
      Top             =   3540
      Width           =   1455
   End
End
Attribute VB_Name = "frmWork_QueueCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlngModule As Long
Private mobjQueue As Object
Private mstrPrivs As String
Private mblnLockAfterCall As Boolean
Private mblnQueueQucik As Boolean


Public Function ShowQueueConfig(objQueue As Object, _
                                ByVal lngModule As Long, _
                                ByVal strPrivs As String, _
                                Optional objOwner As Object = Nothing, _
                                Optional ByRef blnLockAfterCall As Boolean = False, _
                                Optional ByRef blnQueueQuick As Boolean = False) As Boolean
'显示pacs队列配置
    ShowQueueConfig = False
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    Set mobjQueue = objQueue
    
    Call LoadTurnPage
    Call LoadStudyDept
    Call ReadCfgParameter
    
    CheckAddHeight
    Me.Show 1, objOwner

    blnLockAfterCall = mblnLockAfterCall
    ShowQueueConfig = mblnOK
    blnQueueQuick = mblnQueueQucik
End Function


Private Sub ReadCfgParameter()
'读取配置参数
    Dim i As Long
    Dim strColumnInfo As String
    
    If mlngModule = 1291 Then
        chkLockAfterCall.value = zlDatabase.GetPara("呼叫后锁定采集", glngSys, mlngModule, "0")
        mblnLockAfterCall = chkLockAfterCall.value
    End If
    
    '读取排队队列信息定义
    strColumnInfo = zlDatabase.GetPara("排队队列信息定义", glngSys, mlngModule, "排队号码,患者姓名")
    
    For i = 0 To 9
        chkColumn(i).value = Int(IIf(InStr(1, "," & strColumnInfo & ",", "," & chkColumn(i).tag & ",") > 0, vbChecked, vbUnchecked))
    Next i
    
    '读取呼叫队列信息定义
    strColumnInfo = zlDatabase.GetPara("呼叫队列信息定义", glngSys, mlngModule, "排队号码,患者姓名")
    
    For i = 0 To 9
        chkCalledColumn(i).value = Int(IIf(InStr(1, "," & strColumnInfo & ",", "," & chkCalledColumn(i).tag & ",") > 0, vbChecked, vbUnchecked))
    Next i
    
    chkShowMySelfCalled.value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\排队叫号", "只显示自己呼叫的队列", "1"))
    
    chkQueueQuick.value = Val(zlDatabase.GetPara("自动弹出快捷呼叫窗口", glngSys, mlngModule, "1"))
    mblnQueueQucik = IIf(chkQueueQuick.value = 1, True, False)
End Sub



Private Sub cbxDept_Click()
On Error GoTo errHandle
    Call LoadExeRoom(cbxDept.ItemData(cbxDept.ListIndex))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkLockAfterCall_Click()
    mblnLockAfterCall = chkLockAfterCall.value
End Sub


Private Sub chkQueueQuick_Click()
On Error GoTo errHandle
    mblnQueueQucik = IIf(chkQueueQuick.value = 1, True, False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    
    Unload Me
End Sub


Private Sub cmdPrintSet_Click()
'设置打印机
On Error GoTo errHandle
    Call mobjQueue.QueueOper.PrintSet
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdSure_Click()
'保存参数
On Error GoTo errHandle
    Dim i As Long
    Dim strColumnInf As String
    
    zlDatabase.SetPara "本机执行间科室", cbxDept.ItemData(cbxDept.ListIndex), glngSys, mlngModule
    zlDatabase.SetPara "本机执行间名称", cbxRoomName.Text, glngSys, mlngModule
    zlDatabase.SetPara "接诊后跳转页面", cbxTurnPage.Text, glngSys, mlngModule
    If mlngModule = 1291 Then zlDatabase.SetPara "呼叫后锁定采集", chkLockAfterCall.value, glngSys, mlngModule
    
    '保存排队队列配置
    strColumnInf = ""
    For i = 0 To 9
        If chkColumn(i).value = vbChecked Or chkColumn(i).tag = "排队号码" Or chkColumn(i).tag = "患者姓名" Then
            If Trim(strColumnInf) <> "" Then strColumnInf = strColumnInf & ","
            strColumnInf = strColumnInf & chkColumn(i).tag
        End If
    Next i
    
    zlDatabase.SetPara "排队队列信息定义", strColumnInf, glngSys, mlngModule
    
    '保存呼叫队列配置
    strColumnInf = ""
    For i = 0 To 10
        If chkCalledColumn(i).value = vbChecked Or chkCalledColumn(i).tag = "排队号码" Or chkCalledColumn(i).tag = "患者姓名" Then
            If Trim(strColumnInf) <> "" Then strColumnInf = strColumnInf & ","
            strColumnInf = strColumnInf & chkCalledColumn(i).tag
        End If
    Next i
    
    zlDatabase.SetPara "呼叫队列信息定义", strColumnInf, glngSys, mlngModule
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\排队叫号", "只显示自己呼叫的队列", chkShowMySelfCalled.value
    
    zlDatabase.SetPara "自动弹出快捷呼叫窗口", chkQueueQuick.value, glngSys, mlngModule
    
    mblnOK = True
    
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdVoiceCfg_Click()
'打开语音配置窗口
On Error GoTo errHandle
    If mobjQueue Is Nothing Then Exit Sub
    
    Call mobjQueue.ShowVoiceConfig

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadTurnPage()
'载入接诊后的跳转页面
    Dim strTurnPage As String
    Dim strPages() As String
    Dim i As Long
    
    strTurnPage = zlDatabase.GetPara("接诊后跳转页面", glngSys, mlngModule, "")

    If mlngModule = 1290 Then
        strPages = Split("影像,报告,费用,医嘱,病历", ",")
    Else
        strPages = Split("采集,报告,费用,医嘱,病历", ",")
    End If
    
    cbxTurnPage.Clear
    
    Call cbxTurnPage.AddItem("")
    
    For i = 0 To UBound(strPages)
        If Trim(strPages(i)) <> "" Then
            cbxTurnPage.AddItem strPages(i)
        End If
        
        If Trim(strPages(i)) = strTurnPage Then
            cbxTurnPage.ListIndex = i + 1
        End If
    Next i
    
    If cbxTurnPage.ListIndex < 0 Then cbxTurnPage.ListIndex = 0
End Sub


Private Sub LoadStudyDept()
'载入检查科室
    Dim strSql As String
    Dim rsData As New ADODB.Recordset
    Dim str来源 As String
    Dim strCfgDept As String
    
    str来源 = "1,2,3"
    
    strCfgDept = zlDatabase.GetPara("本机执行间科室", glngSys, mlngModule)

    If CheckPopedom(mstrPrivs, "所有科室") Then
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')> 0 And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    Else
        
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=" & UserInfo.ID & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')>0  And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    End If

    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询执行间所属检查科室", CStr("," & str来源 & ","))
    

    Do Until rsData.EOF
        cbxDept.AddItem Nvl(rsData!名称)
        cbxDept.ItemData(cbxDept.ListCount - 1) = Val(Nvl(rsData!ID))
        
        If Nvl(rsData!ID) = strCfgDept Then
            cbxDept.ListIndex = cbxDept.ListCount - 1
        End If
        
        rsData.MoveNext
    Loop
        
    If cbxDept.ListCount > 0 And cbxDept.ListIndex < 0 Then cbxDept.ListIndex = 0
End Sub

Private Sub LoadExeRoom(ByVal lngDeptID As Long)
'载入执行间
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strCfgRoom As String
    
    strCfgRoom = zlDatabase.GetPara("本机执行间名称", glngSys, mlngModule)
    
    strSql = "select 执行间,设备名 from 医技执行房间 a, 影像设备目录 b Where a.检查设备=b.设备号(+) and 科室ID=[1] order by 执行间"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询科室执行间", lngDeptID)
    
    cbxRoomName.Clear
    If rsData.RecordCount <= 0 Then Exit Sub
    
    While Not rsData.EOF
        cbxRoomName.AddItem Nvl(rsData!执行间) & "-" & Nvl(rsData!设备名)
        
        If Nvl(rsData!执行间) & "-" & Nvl(rsData!设备名) = strCfgRoom Then
            cbxRoomName.ListIndex = cbxRoomName.ListCount - 1
        End If
        
        rsData.MoveNext
    Wend
    
    If cbxRoomName.ListCount > 0 And cbxRoomName.ListIndex <= 0 Then cbxRoomName.ListIndex = 0
End Sub

Private Sub CheckAddHeight()
    '判断是否需要增加窗体高度，104686相关，如果是采集工作站，需要增加高度以显示“呼叫后锁定采集”这个参数
    If mlngModule = 1291 Then
        chkLockAfterCall.Visible = True
    Else
        chkLockAfterCall.Visible = False
    End If
End Sub


