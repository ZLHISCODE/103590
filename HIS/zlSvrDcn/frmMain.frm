VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "数据变动通知服务"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9630
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer TimerState 
      Interval        =   1000
      Left            =   4320
      Top             =   600
   End
   Begin VB.Timer tmrDcn 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   600
   End
   Begin XtremeSuiteControls.TabControl tabMain 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
      _Version        =   589884
      _ExtentX        =   4683
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6090
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2223
            MinWidth        =   882
            Picture         =   "frmMain.frx":4D4A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10716
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "2019/10/17"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "13:14"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   5040
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1320
      Top             =   840
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Bindings        =   "frmMain.frx":5165
      Left            =   1800
      Top             =   840
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":5179
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CommandBarIDCond
    conMenu_Start = 1
    conMenu_Stop
    conMenu_Log_Clear
    conMenu_Clear
    conMenu_Setting
    conMenu_Exit
End Enum

Private Enum ChangeType
    DcnInsert = 1
    DcnUpdate = 2
    DcnDelete = 3
End Enum

Private mblnStartUp As Boolean

Private mblnDcn As Boolean  '是否开启了DCN
Private mblnOciConnected As Boolean
Private mblnCancel As Boolean

Private mrsNoticeSet As ADODB.Recordset '保存注册DCN的Notice信息
Private mfrmNoticeList As New frmNoticeList
Private mfrmNoticeLog As New frmNoticeLog

Private mlngCheckInterval As Long   'DCN存活时间更新间隔
Private mlngCheck As Long

Private Sub cbsMain_Resize()
    Dim lngTop As Long, lngBottom As Long
    Dim lngLeft As Long, lngRight As Long
    On Error Resume Next
    
    cbsMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    With tabMain
        .Left = 0
        .Top = lngTop + 10
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowFullAfterDelay = True
        .ShowTextBelowIcons = True
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    Set cbsMain.Icons = imgMain.Icons
    cbsMain.ActiveMenuBar.Visible = False
    
    '工具栏部分
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Start, "开启")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Stop, "停止")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = False
    
        Set objControl = .Add(xtpControlButton, conMenu_Log_Clear, "清空运行日志")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Clear, "本地日志清理")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Setting, "设置")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = False
        
        Set objControl = .Add(xtpControlButton, conMenu_Exit, "退出")
        objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
    End With
    
End Sub

Private Sub InitTab()
    '功能:初始化tab控件
    Dim objItem As TabControlItem
    
    With tabMain.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .Color = xtpTabColorOffice2003
    End With
    
    Set objItem = tabMain.InsertItem(1, "运行日志", mfrmNoticeLog.hwnd, 0)
    tabMain.InsertItem 2, "数据变动通知列表", mfrmNoticeList.hwnd, 0
    
    objItem.Selected = True
End Sub

Private Sub InitNoticeSet()
    Set mrsNoticeSet = GetNoticeList
    If mrsNoticeSet Is Nothing Then Exit Sub
    
    mfrmNoticeList.SetDataSource mrsNoticeSet
End Sub

Private Sub Form_Activate()
    If mblnStartUp Then Exit Sub
    mblnStartUp = True
    
    
    DoEvents    '防止注册时间较长  界面卡顿
    
    Call ChangeDcnState(1)
    Call UpdateCmdState
    
    Me.Caption = Me.Caption & " - [" & gstrServer & "]"
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset, strSql As String
    
    glngPort = 9999 '默认端口9999
    gintLog = 1 '默认保存本地日志
    gintInterval = 200  '默认刷新频率 200ms

    '格式:IP;端口;状态;会话号
    strSql = "SELECT 参数值 FROM zltools.zloptions WHERE 参数号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSql, Me.Caption, 27)
    If rs.RecordCount <> 0 Then
        If rs!参数值 & "" <> "" Then
            glngPort = Split(rs!参数值, ";")(1)
        End If
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSql, Me.Caption, 28)
    If rs.RecordCount <> 0 Then
        If rs!参数值 & "" <> "" Then
            gintLog = Val(rs!参数值)
        End If
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSql, Me.Caption, 29)
    If rs.RecordCount <> 0 Then
        If rs!参数值 & "" <> "" Then
            gintInterval = Val(rs!参数值)
        End If
    End If
    
    gstrIp = winSock.LocalIP
    glngSid = GetSid
    
    tmrDcn.Interval = gintInterval
    mlngCheckInterval = GetCheckInterval
    
    zlCommFun.SetWindowsInTaskBar Me.hwnd, True
    Call InitCommandBar
    Call InitTab
    
    Call InitNoticeSet
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mblnCancel = False
    
    If MsgBox("你是否真的要退出自动提醒服务？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
        Cancel = True
        mblnCancel = True
        Exit Sub
    End If

    On Error Resume Next
    Unload mfrmNoticeList: Set mfrmNoticeList = Nothing
    Unload mfrmNoticeLog: Set mfrmNoticeLog = Nothing
    
    If mblnDcn Then
        Call OCI_UnRigister
        Call UpdateDcnState2DB(0)
    End If
    If gcnOracle.State <> adStateClosed Then gcnOracle.Close
End Sub

Private Sub ChangeDcnState(ByVal intType As Integer)
    '功能:开启或关闭DCN
    'intType=1: 开启  intType=0:关闭
    Dim strTmp As String
    
    If intType = 0 And mblnDcn = False Then Exit Sub
    If intType = 1 And mblnDcn = True Then Exit Sub
    
    If UpdateDcnState2DB(intType) = False Then   '更新消息收发器的状态
        Exit Sub
    End If
    
    If intType = 0 Then
        strTmp = "正在关闭数据变动通知..."
        stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
        
        Call DcnStop
        tmrDcn.Enabled = False
        If winSock.State <> sckClosed Then
            winSock.Close
        End If
        
        strTmp = "数据变动通知已关闭。"
        stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
    Else
        strTmp = "正在开启数据变动通知注册..."
        stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
        mfrmNoticeLog.Refresh
        
        If winSock.State <> sckOpen Then
            winSock.LocalPort = glngPort
            winSock.Bind
        End If
        
        If DcnStart = False Then
            Call UpdateDcnState2DB(0)
            strTmp = "数据库连接失败，无法开启数据变动通知 。"
            stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
            Exit Sub
        End If
        Call DcnRigister
        Call UpdateDcnTime
        
        tmrDcn.Enabled = True
        strTmp = "数据变动通知已开启。"
        stbThis.Panels(2).Text = strTmp: mfrmNoticeLog.WriteLog strTmp, 1
    End If
    
    If mblnDcn = True Then
        mblnDcn = False
    Else
        mblnDcn = True
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_Start
            Call ChangeDcnState(1)
        Case conMenu_Stop
            Call ChangeDcnState(0)
        Case conMenu_Setting
            frmNoticeSet.ShowEdit
            Exit Sub
        Case conMenu_Clear
            frmLogClear.ShowClearLog
            Exit Sub
        Case conMenu_Exit
            Unload Me
            Exit Sub
        Case conMenu_Log_Clear
            Call mfrmNoticeLog.ClearLog
            Exit Sub
    End Select
    Call UpdateCmdState
End Sub

Private Sub DcnStop()

    If lpPrevWndProc <> 0 Then
        UnHook Me.hwnd
        lpPrevWndProc = 0
    End If
End Sub

Private Function DcnStart() As Boolean
    '验证zlNoticeLib是否可用,并登录数据库
    Dim strServiceName As String, strIp As String, strPort As String
    Dim i  As Integer
    
    On Error GoTo errH
    If mblnOciConnected Then    '如果已经建立了OCI连接,就不需要再连接
        If lpPrevWndProc = 0 Then Hook Me.hwnd  '绑定窗口函数
        DcnStart = True
        Exit Function
    End If
    
    GetServerInfo gstrServer, strServiceName, strIp, strPort
    
    mblnDcn = OCI_ConnCreate(strIp & ":" & strPort & "/" & strServiceName, gstrUserName, gstrUserPwd)    '登录OCI

    '如果登录失败,尝试使用注册表中的IP信息
    If mblnDcn = False Then
        strIp = GetSetting("ZLSOFT\公共模块", "zlSvrNotice", "IP", strIp)
        strPort = GetSetting("ZLSOFT\公共模块", "zlSvrNotice", "PORT", strPort)
        strServiceName = GetSetting("ZLSOFT\公共模块", "zlSvrNotice", "Server", strServiceName)
        
        If strPort <> "" Then
            mblnDcn = OCI_ConnCreate(strIp & ":" & strPort & "/" & strServiceName, gstrUserName, gstrUserPwd)    '登录OCI
        End If
    End If
    
    '如果注册表中没有Ip信息或者登录失败,就弹出确认框
    Do While mblnDcn = False And i < 3
        mfrmNoticeLog.WriteLog "数据变动通知组件连接数据库失败，请检查IP、端口、服务名等信息是否正确，同时检查Sqlnet.ora文件中不含有 NAMES.DIRECTORY_PATH 配置", 1
         If frmUserCheckLogin.GetSerInfo(strIp, strPort, strServiceName) = True Then
            i = i + 1
            mblnDcn = OCI_ConnCreate(strIp & ":" & strPort & "/" & strServiceName, gstrUserName, gstrUserPwd)    '登录OCI
         Else
            Exit Function
         End If
    Loop
    mblnOciConnected = mblnDcn
    DcnStart = mblnDcn
    Exit Function
errH:
    ErrCenter
End Function

Private Sub DcnRigister()
    '注册DCN服务
    Dim strSql As String, strTmp As String
    
    If mrsNoticeSet Is Nothing Then Exit Sub
    mrsNoticeSet.Filter = "Status =  1" '过滤掉已停用的配置
    If mrsNoticeSet.RecordCount = 0 Then Exit Sub
    
    mrsNoticeSet.Sort = "Tableowner,Tablename"
    
    Do While Not mrsNoticeSet.EOF
        If strTmp <> mrsNoticeSet!Tableowner & "." & mrsNoticeSet!Tablename Then
            strTmp = mrsNoticeSet!Tableowner & "." & mrsNoticeSet!Tablename
            strSql = "Select  * from " & strTmp
            OCI_Register Me.hwnd, strSql '注册信息,传入句柄
        End If
        mrsNoticeSet.MoveNext
    Loop
    If lpPrevWndProc = 0 Then Hook Me.hwnd  '绑定窗口函数
End Sub

Private Sub SendNotice()
    '功能:循环通知队列, 发送消息
    Dim strOwner As String, strTable As String, strRowid As String
    Dim arrTmp()  As String, lngNoticeCode As Long, intChangeType As Integer
    Dim strSql As String, rsData As New ADODB.Recordset, strCols As String
     
    On Error GoTo errH
    With gcolNotice
        If .Count = 0 Then Exit Sub

        '变动信息格式为:   变动类型-所有者.表名-Rowid
        arrTmp = Split(.Item(1), "-")
        intChangeType = arrTmp(0)
        strRowid = arrTmp(2)
        
        arrTmp = Split(arrTmp(1), ".")
        strOwner = arrTmp(0)
        strTable = arrTmp(1)
        
        '当一个事务变动条数大于>80时,返回的RowID为1
        If strRowid = "1" Then
            mfrmNoticeLog.WriteLog Now & "   表" & strTable & "数据变动数量过大，无法获取Rowid，不发送通知", 1
            .Remove 1
            Exit Sub
        End If
        
        If intChangeType = ChangeType.DcnDelete Then '如果变动类型是删除,就移除该条变动信息
            .Remove 1
            Exit Sub
        End If
        
        '根据Table和Owner进行过滤
        mrsNoticeSet.Filter = "Tableowner = '" & strOwner & "' And Tablename = '" & strTable & "' And Status = 1"
        
        If mrsNoticeSet.RecordCount = 0 Then
            .Remove 1
            Exit Sub
        End If
        
        '一条数据变动,可能涉及多条通知设定,循环通知设定,依次发送数据
        Do While Not mrsNoticeSet.EOF
        
            '获取变动数据
            If mrsNoticeSet!ReceiverTab & "" = "" Then
                strSql = "Select  " & IIf(mrsNoticeSet!ReceiverCols & "" = "", "1", mrsNoticeSet!ReceiverCols) & " From " & strOwner & "." & strTable & " A Where Rowid =  [1]  " & IIf(mrsNoticeSet!Filter & "" = "", "", " And " & mrsNoticeSet!Filter)
            Else
                If mrsNoticeSet!ReceiverCols & "" = "" Then
                    strCols = "1"
                Else
                    strCols = "B." & mrsNoticeSet!ReceiverCols
                    strCols = Replace(strCols, ",", ",B.")  '替换别名
                End If
            
                strSql = "Select " & strCols & " From " & strOwner & "." & strTable & " A , " & mrsNoticeSet!ReceiverTab & " B Where A.Rowid =[1] And " & mrsNoticeSet!ReceiverRelas & IIf(mrsNoticeSet!Filter & "" = "", "", " And " & mrsNoticeSet!Filter)
            End If
            
            Set rsData = zlDatabase.OpenSQLRecord(strSql, "获取变动数据", strRowid)
            If rsData.RecordCount > 0 Then
                Post2Client rsData, mrsNoticeSet!Noticekind, intChangeType, strRowid, mrsNoticeSet!NoticeCode, strOwner, strTable, _
                                mrsNoticeSet!ReceiverCols & "", mrsNoticeSet!SplitChar & "", _
                                mrsNoticeSet!ReceiverIP & "", mrsNoticeSet!ReceiverStaffKind & "", mrsNoticeSet!ReceiverDeptKind & ""
            End If
            mrsNoticeSet.MoveNext
        Loop
        
        .Remove 1   '执行完成后,移除该条变动信息
    End With
    
    Exit Sub
errH:
    '记录错误
    gcolNotice.Remove 1  '删除当前
    If 0 = 1 Then
        Resume
    End If
    mfrmNoticeLog.WriteLog "转发表" & strTable & "数据变动通知时发生错误 " & Err.Description, 1
End Sub

Private Sub Post2Client(ByVal rsData As ADODB.Recordset, intNoticeKind As Integer, intChangeType As Integer, _
                                strRowid As String, lngNoticeCode As Long, strOwner As String, strTable As String, _
                                strReceiverCol As String, strSplitChar As String, _
                                strReceiverIp As String, strReceiverStaffKind As String, strReceiverDeptKind As String)
    '功能:将变动消息发送到客户端
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim strField1 As String, strField2 As String
    Dim strIp  As String, lngPort As Long
    Dim strTmp As String, arrTmp() As String, i As Integer
    
    '提取ReceiverCol的值
    If strReceiverCol <> "" Then
        If InStr(1, strReceiverCol, ",") > 0 Then
            strField1 = rsData.Fields(0)
            strField2 = rsData.Fields(1)
        Else
            strField1 = rsData.Fields(0)
        End If
    End If
    
    If strSplitChar <> "" Then  '如果有分隔符,就将分隔符替换为逗号,以便使用f_str2list函数
        strField1 = Replace(strField1, strSplitChar, ",")
        strField2 = Replace(strField2, strSplitChar, ",")
    End If
    
    Select Case intNoticeKind
        Case 0 '通知所有客户端
            strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession Where 状态=1"
            
        Case 1 '指定部门，只发送至工作站当前部门
            If strSplitChar <> "" Then
                strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession Where 状态=1  And 当前部门ID " & _
                                " In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)"
            Else
                strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession Where 状态=1  And 当前部门ID = [1]"
            End If
            
        Case 2 '指定用户姓名
            If strSplitChar <> "" Then
                strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession Where 状态=1  And 人员姓名 " & _
                                " In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)"
            Else
                strSql = "Select IP,消息端口,工作站  From Zltools.Zlclientsession Where 状态=1  And 人员姓名= [1]"
            End If
            
        Case 3 '指定部门+位置
            If strField2 = "" Then   '若该行位置为空,则向所有部门发送
                strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession  Where 状态=1  And 当前部门ID= [1]"
            Else
                If InStr(1, strField2, ",") > 0 Then
                    strTmp = ""
                    arrTmp = Split(strField2, ",")
                    For i = 0 To UBound(arrTmp)
                        strTmp = strTmp & IIf(strTmp = "", "", "Or") & " Instr(',' || 当前位置 || ',' , '," & arrTmp(i) & ",' )>0 "
                    Next
                    strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession  Where 状态=1  And 当前部门ID= [1] " & _
                                "And ( " & strTmp & " )"
                Else
                    strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession  Where 状态=1  And 当前部门ID= [1] And instr(',' || 当前位置 || ',' , '," & strField2 & ",' )>0 "
                End If
            End If
            
        Case 4  '指定用户的用户名
            If strSplitChar <> "" Then
                strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession Where 状态=1  And 用户名 " & _
                                " In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)"
            Else
                strSql = "Select IP,消息端口,工作站  From Zltools.Zlclientsession Where 状态=1  And 用户名= [1]"
            End If
            
        Case 5  '指定IP、端口
            strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession  Where 状态=1  And IP = [1] And 消息端口 = [2] "
            strField1 = Split(strReceiverIp, ":")(0)    '按照IP\端口发送消息,直接取ReceiverIP中的值
            strField2 = Split(strReceiverIp, ":")(1)
            
        Case 6  '指定性质的工作站
            strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession  Where 状态=1 And 人员性质 = [1]"
            strField1 = strReceiverStaffKind    '按照人员性质发送消息,直接取ReceiverStaffKind中的值
            
        Case 7  '指定部门，关联检查工作站所属全部部门
            If strSplitChar = "" Then
                strSql = "Select Ip, 消息端口, 工作站 From Zltools.Zlclientsession Where 状态 = 1 And 当前部门id = [1]" & vbNewLine & _
                            "Union" & vbNewLine & _
                            "Select Ip, 消息端口, 工作站 From Zltools.Zlclientsession A, Zltools.Zlclientdepts B Where a.状态 = 1 And a.会话号 = b.会话号 And b.部门id = [1]"
            Else
                strSql = "Select Ip, 消息端口, 工作站" & vbNewLine & _
                            "From Zltools.Zlclientsession" & vbNewLine & _
                            "Where 状态 = 1 And 当前部门id In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)" & vbNewLine & _
                            "Union" & vbNewLine & _
                            "Select Ip, 消息端口, 工作站" & vbNewLine & _
                            "From Zltools.Zlclientsession A, Zltools.Zlclientdepts B" & vbNewLine & _
                            "Where a.状态 = 1 And a.会话号 = b.会话号 And b.部门id In (Select /*+ cardinality(a,10)*/ Column_Value From Table(f_Str2list([1])) A)"
            End If
            
        Case 8  '指定部门性质
            If strField1 = "" Then   '若该行位置为空,则向所有同性质部门发送
                strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession  Where 状态=1  And 部门性质 = [2]"
            Else
                If InStr(1, strField1, ",") > 0 Then    '一条数据发送给多个位置
                    strTmp = ""
                    arrTmp = Split(strField1, ",")
                    For i = 0 To UBound(arrTmp)
                        strTmp = strTmp & IIf(strTmp = "", "", "Or") & " Instr(',' || 当前位置 || ',' , '," & arrTmp(i) & ",' )>0 "
                    Next
                    strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession  Where 状态=1  And 部门性质 = [2] " & _
                                "And ( " & strTmp & " )"
                Else
                    strSql = "Select IP,消息端口,工作站 From Zltools.Zlclientsession  Where 状态=1  And 部门性质 = [2]  And instr(',' || 当前位置 || ',' , '," & strField1 & ",' )>0 "  '一台工作站同时设置多个位置
                End If
            End If
            strField2 = strReceiverDeptKind '按照部门性质发送消息,直接去ReceiverDeptKind中的值
    End Select
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取待发送工作站", strField1, strField2)
    If rsTmp.RecordCount <> 0 Then
        
        '记录日志类型
        gstrBuild.Clear
        gstrBuild.Append Now
        gstrBuild.Append "   表：": gstrBuild.Append strOwner
        gstrBuild.Append ".": gstrBuild.Append strTable
        gstrBuild.Append "，数据变动类型: ": gstrBuild.Append Decode(intChangeType, ChangeType.DcnInsert, "新增", ChangeType.DcnUpdate, "修改", "所有")
        mfrmNoticeLog.WriteLog gstrBuild.ToString, 2
        
        gstrBuild.Clear
        Do While Not rsTmp.EOF
            winSock.RemoteHost = rsTmp!IP
            winSock.RemotePort = rsTmp!消息端口
            winSock.SendData lngNoticeCode & "-" & intChangeType & "-" & strOwner & "-" & strTable & "-" & strRowid
            
            '拼接字符串,记录发送信息
            If gstrBuild.Length <> 0 Then gstrBuild.Append vbNewLine
            gstrBuild.Append "   工作站："
            gstrBuild.Append rsTmp!工作站: gstrBuild.Append "("
            gstrBuild.Append rsTmp!IP: gstrBuild.Append ":"
            gstrBuild.Append rsTmp!消息端口: gstrBuild.Append "）， 已发送"
            
            rsTmp.MoveNext
        Loop
        If gstrBuild.Length > 0 Then mfrmNoticeLog.WriteLog gstrBuild.ToString, 2
    End If
    
End Sub

Private Sub UpdateCmdState()
    '设置按钮的可用性
    Dim objControl As CommandBarControl
    
    '开启
    Set objControl = cbsMain.FindControl(, conMenu_Start)
    objControl.Enabled = Not mblnDcn
    '停止
    Set objControl = cbsMain.FindControl(, conMenu_Stop)
    objControl.Enabled = mblnDcn
    
End Sub

Private Sub TimerState_Timer()
        
    TimerState.Enabled = False
    
    mlngCheck = mlngCheck + 1
        
    If mlngCheck > mlngCheckInterval Then
        Call UpdateDcnTime
        mlngCheck = 0
    End If
    
    TimerState.Enabled = True
End Sub

Private Sub tmrDcn_Timer()
    tmrDcn.Enabled = False
    Call SendNotice
    tmrDcn.Enabled = True
End Sub
