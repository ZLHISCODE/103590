VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDrugStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, FrmMain As Object, strUsername As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '返回：
    '------------------------------------------------
    Dim frmTemp As Form
    Dim strTitle As String
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim strPrivs As String

    Set gcnOracle = cnMain
    Set gobjBrower = FrmMain
    
    On Error GoTo errHandle
    
    glngSys = lngSys
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("kind"), Default:="")
    
    gstrMatchMethod = GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0)
    
    gbytSimpleCodeTrans = Val(zlDataBase.GetPara("简码匹配方式切换", , , 1))
    
    gstrUnitName = GetUnitName()
    gstrDbUser = strUsername
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    glngModul = lngModul
    gstrprivs = GetPrivFunc(glngSys, lngModul)
    App.Title = IIf(glngSys \ 100 = 1, "药房事务", "零售药店管理")
    If Not CheckValid Then Exit Sub
    
    '读取药房部分要使用的系统参数值
    Call GetSysParms
    
    '取药品价格等级
    Call GetPriceClass
    
    '-------------------------------------------------
    'PASS接口初始化:只初始化一次
    If IsHavePrivs(gstrprivs, "合理用药监测") Then
        If InStr(",1341,1342,1345,", "," & lngModul & ",") > 0 And gobjPass Is Nothing Then
            On Error Resume Next
            Set gobjPass = CreateObject("zlpassinterface.clspass")
            Error.Clear: On Error GoTo 0
            
            If Not gobjPass Is Nothing Then
                If Not gobjPass.zlPassInit_YF(gcnOracle, glngSys, lngModul) Then
                    Set gobjPass = Nothing
                Else
                    gintPass = PassUseable
                End If
            End If
        End If
    End If
    '------------------------------------------------
    
    '1340    发药窗口管理
    '1341    药品处方发药
    '1342    药品部门发药
    '1343    药品申领管理
    '1344    协定药品入库
    '1345    PIVA管理
    
    Call setNOtExcetePrice '将到时间还没有生效的价格执行了
    
    Select Case lngModul
    Case 1332 '药品结存
        frmClosingAccount.Show , FrmMain
    Case 1340
        frm发药窗口.Show , FrmMain
    Case 1341
        '单据操作控制
        Call LoadBillControl
        
        frm药品处方发药New.Show , FrmMain
    Case 1342
        '单据操作控制
        Call LoadBillControl
        
        frm部门发药管理New.Show , FrmMain
        If frm部门发药管理New.mblnEnter = False Then Unload frm部门发药管理New
    Case 1343
        strTitle = "药品申领管理"
        For Each frmTemp In Forms
            If frmTemp.Caption = strTitle Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        frmRequestDrugList.ShowList FrmMain
    Case 1344
        
        strTitle = "协定药品入库"
        For Each frmTemp In Forms
            If frmTemp.Caption = strTitle Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmTemp = New frmMainList
        frmTemp.ShowList lngModul, strTitle, FrmMain
    Case 1345
        '检查部门设置
        
        gstrSQL = "Select Distinct A.ID, A.名称" & _
            " From 部门表 A, 部门性质说明 B " & _
            " Where A.ID = B.部门id And B.工作性质 = '配制中心' And " & _
            " B.部门id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房') " & _
            " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取配置中心")
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在部门管理中设置输液配置中心部门(必须具有药房和配制中心属性)！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查部门人员
        gstrSQL = "Select Distinct P.ID, P.名称" & _
            " From 部门表 P " & _
            " Where (P.站点 = '" & gstrNodeNo & "' Or P.站点 is Null) And P.ID In (Select Distinct A.部门id " & _
            " From 部门人员 A, 部门性质说明 B " & _
            " Where A.人员id = [1] And A.部门id = B.部门id And B.工作性质 = '配制中心' And " & _
            " B.部门id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房')) And " & _
            " (P.撤档时间 Is Null Or P.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取配置中心人员", glngUserId)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "你不是输液配制中心人员，不能使用本模块！", vbInformation, gstrSysName
            Exit Sub
        End If

        frmPIVAMain.Show , FrmMain
    Case 1347
        frm大处方审查.Show , FrmMain
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function AutoSetBatch(ByVal lng库房id As Long, ByVal str发送号 As String, cnMain As ADODB.Connection) As Boolean
    '自动设置配置中心批次
    'lng库房id：配置中心部门id
    'str发送号：医嘱发送号
'
End Function
Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, strUsername As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   lngMain:主窗体句柄
    '返回：
    '------------------------------------------------
    Dim frmTemp As Form
    Dim strTitle As String
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim strPrivs As String

    Set gcnOracle = cnMain
    
    glngSys = lngSys
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("kind"), Default:="")
    
    gstrMatchMethod = GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0)
    
    gbytSimpleCodeTrans = Val(zlDataBase.GetPara("简码匹配方式切换", , , 1))
    
    gstrUnitName = GetUnitName()
    gstrDbUser = strUsername
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    glngModul = lngModul
    gstrprivs = GetPrivFunc(glngSys, lngModul)
    App.Title = IIf(glngSys \ 100 = 1, "药房事务", "零售药店管理")
    If Not CheckValid Then Exit Sub
    
    '读取药房部分要使用的系统参数值
    Call GetSysParms
    
    '-------------------------------------------------
    'PASS接口初始化:只初始化一次
    If IsHavePrivs(gstrprivs, "合理用药监测") Then
        If InStr(",1341,1342,1345,", "," & lngModul & ",") > 0 And gobjPass Is Nothing Then
            On Error Resume Next
            Set gobjPass = CreateObject("zlpassinterface.clspass")
            Error.Clear: On Error GoTo 0
            
            If Not gobjPass Is Nothing Then
                If Not gobjPass.zlPassInit_YF(gcnOracle, glngSys, lngModul) Then
                    Set gobjPass = Nothing
                Else
                    gintPass = PassUseable
                End If
            End If
        End If
    End If
    '------------------------------------------------
    
    '1340    发药窗口管理
    '1341    药品处方发药
    '1342    药品部门发药
    '1343    药品申领管理
    '1344    协定药品入库
    '1345    PIVA管理
    
    Call setNOtExcetePrice '将到时间还没有生效的价格执行了
    
    Select Case lngModul
    Case 1332 '药品结存
        OS.ShowChildWindow frmClosingAccount.hWnd, lngMain
    Case 1340
        OS.ShowChildWindow frm发药窗口.hWnd, lngMain
    Case 1341
        '单据操作控制
        Call LoadBillControl
        
        OS.ShowChildWindow frm药品处方发药New.hWnd, lngMain
    Case 1342
        '单据操作控制
        Call LoadBillControl
        
        OS.ShowChildWindow frm部门发药管理New.hWnd, lngMain
        If frm部门发药管理New.mblnEnter = False Then Unload frm部门发药管理New
    Case 1343
        strTitle = "药品申领管理"
        For Each frmTemp In Forms
            If frmTemp.Caption = strTitle Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        frmRequestDrugList.ShowList lngMain
    Case 1344
        
        strTitle = "协定药品入库"
        For Each frmTemp In Forms
            If frmTemp.Caption = strTitle Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmTemp = New frmMainList
        frmTemp.ShowList lngModul, strTitle, lngMain
    Case 1345
        '检查部门设置
        
        gstrSQL = "Select Distinct A.ID, A.名称" & _
            " From 部门表 A, 部门性质说明 B " & _
            " Where A.ID = B.部门id And B.工作性质 = '配制中心' And " & _
            " B.部门id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房') " & _
            " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取配置中心")
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在部门管理中设置输液配置中心部门(必须具有药房和配制中心属性)！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查部门人员
        gstrSQL = "Select Distinct P.ID, P.名称" & _
            " From 部门表 P " & _
            " Where (P.站点 = '" & gstrNodeNo & "' Or P.站点 is Null) And P.ID In (Select Distinct A.部门id " & _
            " From 部门人员 A, 部门性质说明 B " & _
            " Where A.人员id = [1] And A.部门id = B.部门id And B.工作性质 = '配制中心' And " & _
            " B.部门id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房')) And " & _
            " (P.撤档时间 Is Null Or P.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取配置中心人员", glngUserId)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "你不是输液配制中心人员，不能使用本模块！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        OS.ShowChildWindow frmPIVAMain.hWnd, lngMain
    Case 1347
        OS.ShowChildWindow frm大处方审查.hWnd, lngMain
    End Select
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.count = 0)
End Function

'Private Sub Class_Terminate()
'   '在导航台退出时激活,此时注销PASS
'    If gintPass = 1 Then
'        Call PassQuit: gintPass = 0
'    End If
'End Sub
