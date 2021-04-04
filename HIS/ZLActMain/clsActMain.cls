VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private blnReady As Boolean

Public Function BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, ByVal lngMain As Long, _
    ByVal strDBUser As String, ByVal strPassword As String, ByVal strServer As String) As Boolean
'-----------------------------------------------------------------
'功能：根据主程序指定功能，调用执行相关程序
'参数：
'   lngSys,lngModul:需要执行功能的系统号及模块号
'   lngMain:主窗体句柄
'   strDbUser,strPassword,strServer:数据库连接用户名、密码(传入的是已转换过的数据库密码)、服务器
'返回：模块调用是否成功；中途任意一步退出则返回失败
'-----------------------------------------------------------------
    Dim strTitle As String                  '产品标题
    Dim strTag As String                    '旗舰版标志
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strClass As String, objClass As Object
    Dim intCount As Integer, blnLoad As Boolean
    Dim strVersion As String, strCompareVersion As String
    Dim ArrayVersion As Variant
    Dim objRis As Object
    Dim intAtom As Integer, strCommon As String
    Dim strModul As String, strFuncs As String
    
    On Error GoTo ErrMain
    '初始化Login对象
    If gclsLogin Is Nothing Then
        Set gclsLogin = New clsRelogin
    End If
    gstrCommand = CStr(Command())
    gclsLogin.DBUser = UCase(strDBUser)
    gclsLogin.InputPwd = strPassword
    gclsLogin.InputUser = strDBUser
    gclsLogin.ServerName = strServer
    gclsLogin.IsTransPwd = False
    
    Screen.MousePointer = 11
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    '判断是否已经有模态窗体打开，这种情况下不允许再打开模块
    strCommon = ExistModalWindows(lngMain)
    If strCommon <> "" Then
        Screen.MousePointer = 0
        'MsgBox """" & strCommon & """中已有独占的模态窗体打开，请先关闭再执行当前操作。", vbInformation, gstrSysName
        Exit Function
    End If
    glngMain = lngMain
    '结束异常进程
    'Call zlKillHISPID
    '登录相关处理：同一进程只登录一次
    If gcnOracle.State <> adStateOpen Then
        'BH置后
        Load frmTmr
        Call frmTmr.SetTimr(True)
        On Error GoTo ErrMain
        '创建注册部件(用于登录时获取连接对象)
        On Error Resume Next
        Set gobjRegister = CreateObject("zlRegister.clsRegister")
        If gobjRegister Is Nothing Then
            Err.Clear
            MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
            Call frmTmr.SetTimr(False)
            Exit Function
        End If
        On Error GoTo ErrMain
        '连接数据库(固定不转换密码，因为传入的是已转换过的数据库密码)
        Set gcnOracle = gobjRegister.GetConnection(strServer, strDBUser, strPassword, False)
        If gcnOracle.State = adStateClosed Then
            Call frmTmr.SetTimr(False)
            Exit Function
        End If
        '写入本次启动程序的EXE文件名
        Call SaveSetting("ZLSOFT", "公共全局", "执行文件", App.EXEName & ".exe")
        SaveSetting "ZLSOFT", "公共全局", "程序路径", App.Path & "\" & App.EXEName & ".exe"
    
        '是否允许使用本工作站
        If Not CheckAllowByTerminal() Then Call frmTmr.SetTimr(False): Exit Function
        zl9ComLib.SetDbUser gclsLogin.DBUser
        zl9ComLib.SetNodeNo gclsLogin.NodeNo
        '初始化公共部件(其中有注册验证)
        Call InitCommon(gcnOracle)
        On Error GoTo ErrMain
        If gcnOracle.State <> adStateOpen Then
            Call frmTmr.SetTimr(False)
            Exit Function
        End If
        '相关注册信息写入注册表
        Select Case zlRegInfo("授权性质")
        Case "1" '正式
            SaveSetting "ZLSOFT", "注册信息", "Kind", ""
        Case "2" '试用
            SaveSetting "ZLSOFT", "注册信息", "Kind", "试用"
        Case "3" '测试
            SaveSetting "ZLSOFT", "注册信息", "Kind", "测试"
        Case Else '不对
            Screen.MousePointer = 0
            MsgBox "授权性质不正确，程序被迫退出！", vbInformation, gstrSysName
            Call frmTmr.SetTimr(False)
            Exit Function
        End Select
    
        gstrSysName = zlRegInfo("产品简名") & "软件"
        SaveSetting "ZLSOFT", "注册信息", "提示", gstrSysName
        SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
        
        strTitle = zlRegInfo("产品标题")
        If strTitle <> "" Then
            If InStr(strTitle, "-") > 0 Then
                If Split(strTitle, "-")(1) = "Ultimate" Then
                    strTag = "旗舰版"
                ElseIf Split(strTitle, "-")(1) = "Professional" Then
                    strTag = "专业版"
                End If
            End If
        End If
        strTitle = Split(strTitle, "-")(0)
        SaveSetting "ZLSOFT", "注册信息", "产品全称", strTitle
        SaveSetting "ZLSOFT", "注册信息", "产品系列", strTag
        
        SaveSetting "ZLSOFT", "注册信息", "产品名称", zlRegInfo("产品简名")
        SaveSetting "ZLSOFT", "注册信息", "单位名称", zlRegInfo("单位名称", , -1)
        SaveSetting "ZLSOFT", "注册信息", "技术支持商", zlRegInfo("技术支持商", , -1)
        SaveSetting "ZLSOFT", "注册信息", "开发商", zlRegInfo("产品开发商", , -1)
        SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", zlRegInfo("支持商简名")
        SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", zlRegInfo("支持商MAIL")
        SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", zlRegInfo("支持商URL")
        
        SaveSetting "ZLSOFT", "注册信息", UCase("gstrAviPath"), App.Path & "\附加文件"
        
        '多帐套检测：
        '  不支持多帐套，不支持非共享安装
        '  指明了系统号，不存在多帐套的问题
        '###
        
        '指定导航台窗体，用于公共部件回调，可以不指定
        'Call SetMainForm(??)
        
        '初始化打印部件：涉及权限读取，要在授权验证之后
        Call IniPrintMode(gcnOracle, gclsLogin.DBUser)
        '初始新网RIS系统
        On Error Resume Next
        Set objRis = CreateObject("zl9XWInterface.clsHISInner")
        Err.Clear: On Error GoTo 0
        If Not objRis Is Nothing Then
            '固定不转换密码，因为传入的是已转换过的数据库密码
            Call objRis.SaveDBConnectInfo(UCase(strDBUser), strPassword, strServer, False)
        End If
    End If
    Call frmTmr.SetTimr(True)
    blnReady = True
    '----------------------------------------------------------------------------------
    '检查并获取模块对应的部件
    strSQL = "Select 部件,标题 From zlPrograms Where Nvl(系统,0) = [1] And 序号 = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.Title, lngSys, lngModul)
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "没有安装该模块！", vbExclamation, gstrSysName
        Call frmTmr.SetTimr(False)
        Exit Function
    End If
    If IsNull(rsTemp!部件) Then
        Screen.MousePointer = 0
        MsgBox "该模块的部件信息不正常！", vbExclamation, gstrSysName
        Call frmTmr.SetTimr(False)
        Exit Function
    End If
    strClass = rsTemp!部件
    strModul = rsTemp!标题
    
    '获取已经创建的部件
    blnLoad = False
    If Trim(Join(gstrClass)) <> "" Then
        For intCount = 0 To UBound(gstrClass)
            If gstrClass(intCount) = strClass Then
                Set objClass = gobjClass(intCount)
                blnLoad = True: Exit For
            End If
        Next
    End If
    '使用新病历部件
    If UCase(strClass) = UCase("zl9EmrInterface") And blnLoad = False Then
        If gclsLogin.mobjEmr Is Nothing Then
            MsgBox "病历组件创建失败！请检查并重新登录。", vbInformation, gstrSysName
            Call frmTmr.SetTimr(False)
            Exit Function
        ElseIf gclsLogin.mobjEmr.IsInited = False Then
            MsgBox "病历组件未能初始化," & gclsLogin.mobjEmr.GetError(), vbInformation, gstrSysName
            Call frmTmr.SetTimr(False)
            Exit Function
        End If
        If Trim(Join(gstrClass)) <> "" Then
            intCount = UBound(gstrClass)
        Else
            intCount = -1
        End If
        intCount = intCount + 1
        ReDim Preserve gstrClass(intCount)
        gstrClass(intCount) = strClass
        ReDim Preserve gobjClass(intCount)
        Set gobjClass(intCount) = gclsLogin.mobjEmr
        
        If Not gclsLogin.IsEMRProxy Then '使用代理用户登录，则不检查权限
            Dim strSpecify As String '片段，范文权限固定在调用前传递
            If Not gclsLogin.mobjEmr.HasInjectAuthorization(2201) Then
                strSpecify = GetPrivFunc(lngSys, 2201)
                Call gclsLogin.mobjEmr.InjectAuthorization(2201, strSpecify)
            End If
            If Not gclsLogin.mobjEmr.HasInjectAuthorization(2203) Then
                strSpecify = GetPrivFunc(lngSys, 2203)
                Call gclsLogin.mobjEmr.InjectAuthorization(2203, strSpecify)
            End If
        End If
        blnLoad = True
    End If
    '如果没有创建则新创建
    If blnLoad = False Then
        On Error Resume Next
        Set objClass = CreateObject(strClass & ".cls" & Mid(strClass, 4))
        If Err.Number <> 0 Then
            Screen.MousePointer = 0
            MsgBox "部件 " & strClass & " 不能正常创建，请检查安装是否正确！", vbExclamation, gstrSysName
            Err.Clear
            Call frmTmr.SetTimr(False)
            Exit Function
        End If
        On Error GoTo ErrMain
        
        '检查该部件的版本是否满足系统需求(主版本-3;次版本-3;附版本-3)
        If UCase(strClass) <> "ZL9REPORT" And UCase(strClass) <> "ZL9DOC" Then
            strSQL = "Select Nvl(主版本,1) As 主版本,Nvl(次版本,0) As 次版本,Nvl(附版本,0) As 附版本,名称" & _
                    " From zlComponent Where Upper(RTrim(部件))=[1] And 系统=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.Title, UCase(strClass), lngSys)
            If rsTemp.EOF Then
                Screen.MousePointer = 0
                MsgBox "系统表数据不完整，请与系统管理员联系！", vbExclamation, gstrSysName
                Call frmTmr.SetTimr(False)
                Exit Function
            End If
        
            With rsTemp
                '组装版本号为三位主版本、三位次版本及三位附版本
                strCompareVersion = String(3 - Len(!主版本), "0") & !主版本 & "." & _
                                    String(3 - Len(!次版本), "0") & !次版本 & "." & _
                                    String(3 - Len(!附版本), "0") & !附版本
                ArrayVersion = Split(objClass.Version, ".")
                strVersion = String(3 - Len(ArrayVersion(0)), "0") & ArrayVersion(0) & "." & _
                             String(3 - Len(ArrayVersion(1)), "0") & ArrayVersion(1) & "." & _
                             String(3 - Len(ArrayVersion(2)), "0") & ArrayVersion(2)
                If strVersion < strCompareVersion Then
                    Screen.MousePointer = 0
                    MsgBox !名称 & "(" & strClass & ")的版本已不能满足系统的需求，请与系统管理员联系！", vbExclamation, gstrSysName
                    Call frmTmr.SetTimr(False)
                    Exit Function
                End If
            End With
        End If
        
        '加入部件缓存
        If Trim(Join(gstrClass)) <> "" Then
            intCount = UBound(gstrClass)
        Else
            intCount = -1
        End If
        intCount = intCount + 1
        ReDim Preserve gstrClass(intCount)
        gstrClass(intCount) = strClass
        ReDim Preserve gobjClass(intCount)
        Set gobjClass(intCount) = objClass
    End If
    
    '执行部件中的模块功能
    If UCase(strClass) = "ZL9REPORT" Then
        '发布到导航台的自定义报表
        objClass.ReportOpen gcnOracle, lngSys, lngModul, Nothing
    ElseIf UCase(strClass) = UCase("zl9EmrInterface") Then
        On Error GoTo ErrMain
        strFuncs = GetPrivFunc(lngSys, lngModul)
        Call gclsLogin.mobjEmr.CodeMain(lngModul, strModul, lngMain, gclsLogin.EMRUser, gclsLogin.EMRPwd, strFuncs)
    Else
        '创建冲突的私有同义词
        Call CreateSynonyms(lngSys, lngModul)
        
        '为通讯原子赋值
        strCommon = Format(Now, "yyyyMMddHHmm")
        strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
        '加入通讯原子
        intAtom = GlobalAddAtom(strCommon)
        Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
        On Error Resume Next
        objClass.BHCodeMan lngSys, lngModul, gcnOracle, lngMain, gclsLogin.DBUser, gclsLogin
        '错误的参数个数和属性分配,有些BHComdeMan没有最后一个gclsLogin参数，会报该错误
        If Err.Number = 450 Or Err.Number = 450 - 2146828288 Then
            Err.Clear
            On Error GoTo ErrMain
            objClass.BHCodeMan lngSys, lngModul, gcnOracle, lngMain, gclsLogin.DBUser
        Else
            If Err.Number <> 0 Then
                GoTo ErrMain
            Else
                On Error GoTo ErrMain
            End If
        End If
        Call GlobalDeleteAtom(intAtom)
        '因医保部件只有CodeMan()才能获取系统号，在读取参数时必须知道系统号，特写入注册表，如果医保读不到默认为 100
        Call SaveSetting("ZLSOFT", "公共全局", "系统号", lngSys)
    End If
    Call frmTmr.SetTimr(False)
    Screen.MousePointer = 0
    BHCodeMan = True
    Exit Function
ErrMain:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call frmTmr.SetTimr(False)
End Function

Public Function CloseWindows() As Boolean
    Dim intCount As Integer, objClass As Object
    Dim frmThis As Form
    
    CloseWindows = True
    If Trim(Join(gstrClass)) <> "" Then
        Err = 0: On Error Resume Next
        For intCount = 0 To UBound(gstrClass)
            Set objClass = gobjClass(intCount)
            If Not objClass.CloseWindows() Then CloseWindows = False
        Next
    End If
    For Each frmThis In Forms
        Unload frmThis
    Next
    If Not gclsLogin.mobjEmr Is Nothing Then
        gclsLogin.mobjEmr.Dispose
        Set gclsLogin.mobjEmr = Nothing
    End If
    Set gclsLogin = Nothing
End Function

Public Function IsReady() As Boolean
    IsReady = blnReady
End Function

Private Sub Class_Terminate()
    Call Me.CloseWindows
End Sub

