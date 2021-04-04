Attribute VB_Name = "MdlBrower"
Option Explicit
'MDI必须
Public Type Menu_Type
    功能菜单 As Long
    窗口菜单 As Long
    其它功能菜单 As Long
    分隔菜单 As Long
End Type
Public 菜单基准 As Menu_Type
Public Enum 工具清单
    导航功能清单 = 10
    字典管理工具 = 11
    消息收发工具 = 12
    系统选项设置 = 13
    EXCEL报表工具 = 14
    本地参数管理 = 15
End Enum
'外挂功能
Public gobjPlugIn As Object

Public gobjRelogin As Object                   '重启类对象
Public FrmMainface As Form
Public gcnOracle As ADODB.Connection

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrUserFlag As String               '当前用户标志(两位表示)，第1位：是否DBA；第2位：系统所有者
Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码
Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrStation As String                '本工作站名称

Public gstrObj() As String
Public gobjCls() As Object
Public grsMenus As New ADODB.Recordset       '菜单记录集
Public gstrMenuSys As String                '菜单名称
Public gstrCommand As String                '命令行参数 陈东 2010-12-06
Private mlngSysPre As Long                  '上次调用私有同义词检查创建时的系统号
Private mlngWin32 As Long
Private mbln注销 As Boolean

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const Process_Query_Information = &H400
Private Const Still_Active = &H103
'---------------------------------------------------------------------------------------------------
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'关闭系统相关的变量及API函数
'----------------------------------------------------------------------------------------------------
Public Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type
Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
'The GetCurrentProcess function returns a pseudohandle for the current process.
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
'The OpenProcessToken function opens the access token associated with a process.
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'The LookupPrivilegeValue function retrieves the locally unique identifier (LUID) used on a specified system to locally represent the specified privilege name.
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
'The AdjustTokenPrivileges function enables or disables privileges in the specified access token. Enabling or disabling privileges in an access token requires TOKEN_ADJUST_PRIVILEGES access.
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Public Declare Function GetLastError Lib "kernel32" () As Long
'用于ExitWindowsEx
Private Const M_lng关闭计算机及电源 As Long = 8
Public Const EWX_FORCE = 4 '强行关闭程序并注销
'自定义
Public Const WINDOWS95 = 0
Public Const WINDOWSNT = 1

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'下列语句用于检测是否合法调用
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer

'Command命令相关
Public Const INFINITE As Long = &HFFFF&
Private Const SW_HIDE                           As Integer = 0 '隐藏窗口，激活另一个窗口
Private Const NORMAL_PRIORITY_CLASS             As Long = &H20&
Public Const STARTF_USESTDHANDLES = &H100&
Public Const STARTF_USESHOWWINDOW = &H1
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
Public Type STARTUPINFO
    Cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
'注册表安全属性类型
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long


Public Sub ExecuteFunc(lngSys As Long, Components As String, Modul As Long, Optional ByVal strPara As String) ', Identity As Byte
    '-------------------------------------------------------------
    '功能：调用执行指定部件的功能程序
    '参数：
    '   frmbrower：主窗体
    '   Components：部件
    '   Modul：模块编号
    '   Identity：可执行者身份要求
    '-------------------------------------------------------------
    Dim rsCheck As New ADODB.Recordset                  '检测版本是否符合系统需求
    Dim IntCount As Integer, intClients As Integer
    Dim objNow As Object                                '创建的部件对象
    Dim BlnExecute As Boolean                           '是否存在该部件
    Dim StrVersion As String, StrCompareVersion As String
    Dim ArrayVersion
    '合法性检查
    Dim intAtom As Integer, strCommon As String
    Dim strSQL  As String
    
    Err = 0: On Error Resume Next
    FrmMainface.MousePointer = 11
    
    IntCount = UBound(gstrObj)
    If Err <> 0 Then IntCount = -1
    Err = 0
    
    BlnExecute = False
    If IntCount >= 0 Then
        For IntCount = 0 To UBound(gstrObj)
            If gstrObj(IntCount) = Components Then
                BlnExecute = True
                Exit For
            End If
        Next
    End If
    
    '使用新病历部件
    If UCase(Components) = UCase("zl9EmrInterface") And BlnExecute = False Then
        IntCount = UBound(gstrObj)
        IntCount = IntCount + 1
        ReDim Preserve gstrObj(IntCount)
        gstrObj(IntCount) = Components
        If FrmMainface.mobjEmr Is Nothing Then
            MsgBox "病历组件创建失败！请检查并重新登录。", vbInformation, gstrSysName
            Exit Sub
        ElseIf FrmMainface.mobjEmr.IsInited = False Then
            MsgBox "病历组件未能初始化," & FrmMainface.mobjEmr.GetError(), vbInformation, gstrSysName
            Exit Sub
        End If
        If Not gobjRelogin.IsEMRProxy Then '使用代理用户登录，则不检查权限
            Dim strSpecify As String '片段，范文权限固定在调用前传递
            If Not FrmMainface.mobjEmr.HasInjectAuthorization(2201) Then
                strSpecify = GetPrivFunc(lngSys, 2201)
                Call FrmMainface.mobjEmr.InjectAuthorization(2201, strSpecify)
            End If
            If Not FrmMainface.mobjEmr.HasInjectAuthorization(2203) Then
                strSpecify = GetPrivFunc(lngSys, 2203)
                Call FrmMainface.mobjEmr.InjectAuthorization(2203, strSpecify)
            End If
        End If
        BlnExecute = True
    End If
    '--如果没有该部件,则创建--
    If BlnExecute = False Then
        Set objNow = CreateObject(Components & ".Cls" & Mid(Components, 4))
    
        If Err = 0 Then
            On Error GoTo errH
            '--检查该部件的版本是否满足系统需求(主版本-3;次版本-3;附版本-3)[自定义报表部件除外]--
            If Not (UCase(Components) = "ZL9REPORT") And Not (UCase(Components) = "ZL9DOC") And Not OS.IsDesinMode Then
                strSQL = " Select nvl(主版本,1) 主版本,nvl(次版本,0) 次版本,nvl(附版本,0) 附版本,名称 " & _
                          " From ZlComponent Where Upper(Rtrim(部件))=[1] And 系统=[2]"
                Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "部件版本检查", UCase(Components), lngSys)
                With rsCheck
                    If .EOF Then
                        MsgBox "系统表部件表ZlComponent数据不完整，请与软件开发商联系！", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    End If
                    
                    '组装版本号为三位主版本、三位次版本及三位附版本
                    StrCompareVersion = String(3 - Len(!主版本), "0") & !主版本 & "." & _
                                        String(3 - Len(!次版本), "0") & !次版本 & "." & _
                                        String(3 - Len(!附版本), "0") & !附版本
                    ArrayVersion = Split(objNow.Version, ".")
                    StrVersion = String(3 - Len(ArrayVersion(0)), "0") & ArrayVersion(0) & "." & _
                                 String(3 - Len(ArrayVersion(1)), "0") & ArrayVersion(1) & "." & _
                                 String(3 - Len(ArrayVersion(2)), "0") & ArrayVersion(2)
                    
                    If StrVersion < StrCompareVersion Then
                        MsgBox "该部件的版本已不能满足系统的需求，请与软件开发商联系！（" & !名称 & "）", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    End If
                End With
            End If
        
            IntCount = 0
            On Error Resume Next
            IntCount = UBound(gstrObj)
            IntCount = IntCount + 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo errH
            ReDim Preserve gobjCls(IntCount)
            Set gobjCls(IntCount) = objNow
            ReDim Preserve gstrObj(IntCount)
            gstrObj(IntCount) = Components
        '创建部件失败，应该提示
        Else
            Screen.MousePointer = 0
            MsgBox "部件 " & Components & ".Cls" & Mid(Components, 4) & " 不能正常创建，请检查安装是否正确！信息：" & vbNewLine & Err.Description, vbExclamation, gstrSysName
            Err.Clear
            Exit Sub
        End If
    End If
    
    Err = 0: On Error GoTo errH
    '--执行该功能--
    If gstrObj(IntCount) = Components Then
        If UCase(Components) = "ZL9REPORT" Then
            If Modul = 菜单基准.其它功能菜单 Then
                gobjCls(IntCount).ReportMan gcnOracle, FrmMainface
            Else
                
'                strPara = "开始日期=2013-01-01"
                If strPara <> "" Then
                    Dim varPara As Variant
                                        
                    varPara = Split(strPara, "|")
'                    varPara(0) = "开始日期=2013-01-01"
'                    varPara(1) = "结束日期=2014-05-01"
                    
                    '最多支持10个参数，超过10个的不管
                    Select Case UBound(varPara)
                    Case 0
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0))
                    Case 1
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1))
                    Case 2
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2))
                    Case 3
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3))
                    Case 4
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4))
                    Case 5
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5))
                    Case 6
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6))
                    Case 7
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7))
                    Case 8
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8))
                    Case 9
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8)), CStr(varPara(9))
                    Case Else
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8)), CStr(varPara(9))
                    End Select
                    
                Else
                    gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface
                End If
                
            End If
        ElseIf UCase(Components) = UCase("zl9EmrInterface") Then
            Dim strFuncs As String, strModul As String
            
            strSQL = " Select 标题　From zlPrograms Where 序号=[1] "
            Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "系统模块检查", Modul)
            With rsCheck
                    If .EOF Then
                        MsgBox "系统表数据不完整，请与软件开发商联系！", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    Else
                        strModul = !标题
                    End If
            End With
            strFuncs = GetPrivFunc(lngSys, Modul)
            Call FrmMainface.mobjEmr.CodeMain(Modul, strModul, FrmMainface.hwnd, gobjRelogin.EMRUser, gobjRelogin.EMRPwd, strFuncs)
        Else
            Call CreateSynonyms(lngSys, Modul)
            
            '用户站点数检测 (正式版及试用版)
            intClients = Val(zlRegInfo("授权站点"))
            If intClients > 0 Then
                If GetCurStates > intClients Then
                    MsgBox "当前用户登录数超过了最大授权数" & intClients & ",系统将自动结束运行！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If

            
            '为通讯原子赋值
            strCommon = Format(Now, "yyyyMMddHHmm")
            strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
            '加入通讯原子
            intAtom = GlobalAddAtom(strCommon)
            Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
            gobjCls(IntCount).CodeMan lngSys, Modul, gcnOracle, FrmMainface, gstrDbUser
            Call GlobalDeleteAtom(intAtom)
            
            '因医保部件只有CodeMan()才能获取系统号，在读取参数时必须知道系统号，特写入注册表，如果医保读不到默认为 100
            Call SaveSetting("ZLSOFT", "公共全局", "系统号", lngSys)
        End If
    End If
    FrmMainface.MousePointer = 0
    Exit Sub
errH:
    FrmMainface.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ReLogin()
    '功能:完成重新重录
    mbln注销 = True
    
    Call gobjRelogin.ReLogin(FrmMainface)
End Sub

Public Function OwnerUser(ByVal strUserName As String) As Boolean
    Dim RecUser As New ADODB.Recordset
    Dim strSQL As String
    OwnerUser = True
'    With RecUser
    On Error GoTo errH
        strSQL = "Select Count(*) 所有者 From ZlSystems Where 所有者='" & strUserName & "'"
         Set RecUser = zlDatabase.OpenSQLRecord(strSQL, "所有者")
'        .Open "Select Count(*) 所有者 From ZlSystems Where 所有者='" & strUserName & "'", gcnOracle By zq
        
        If RecUser.EOF Then
            If Not IsNull(RecUser!所有者) Then
                If RecUser!所有者 = 0 Then OwnerUser = False
            End If
        End If
    Exit Function
errH:
    OwnerUser = False
    If ErrCenter() = 1 Then
        Resume
    End If
'    End With
End Function

Public Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '创建模块所需对象的同义词(如果已创建则不会再创建)
    On Error Resume Next
    If mlngSysPre <> lngSys Then
        strSQL = "Zl_Createsynonyms(" & lngSys & ")"
        zlDatabase.ExecuteProcedure strSQL, "创建同义词"
        mlngSysPre = lngSys
    End If
End Function

Public Sub AddHistory(ByVal strModul As String)
    Dim str系统 As String, str序号 As String, intMax As Integer
    Dim arr系统 As Variant, arr序号 As Variant, strValue As String
    Dim int系统_Cur As Integer, int序号_Cur As Integer
    Dim int系统_Max As Integer, int序号_Max As Integer
    '最近运行的程序，始终在第一个位置；如果已存在于历史记录中，则将其置于第一个位置
    'strModul:系统 & "," & 模块
    
    intMax = 6
    
    strValue = zlDatabase.GetPara("最近使用模块")
    If UBound(Split(strValue, "|")) >= 1 Then
        str系统 = Trim(Split(strValue, "|")(0))
        str序号 = Trim(Split(strValue, "|")(1))
    End If
    If str系统 = "" Or str序号 = "" Then
        str系统 = Split(strModul, ",")(0)
        str序号 = Split(strModul, ",")(1)
        Call zlDatabase.SetPara("最近使用模块", str系统 & "|" & str序号)
        Exit Sub
    End If
    
    arr系统 = Split(str系统, ",")
    arr序号 = Split(str序号, ",")
    int系统_Max = UBound(arr系统)
    int序号_Max = UBound(arr序号)
    str系统 = Split(strModul, ",")(0): str序号 = Split(strModul, ",")(1)
    If int系统_Max > intMax Then int系统_Max = intMax
    
    For int系统_Cur = 0 To int系统_Max
        int序号_Cur = int系统_Cur
        If int序号_Cur > int序号_Max Then Exit For
        If Not (arr系统(int系统_Cur) = Split(strModul, ",")(0) And arr序号(int序号_Cur) = Split(strModul, ",")(1)) Then
            str系统 = str系统 & "," & arr系统(int系统_Cur)
            str序号 = str序号 & "," & arr序号(int序号_Cur)
        End If
    Next
    Call zlDatabase.SetPara("最近使用模块", str系统 & "|" & str序号)
End Sub

Public Sub CheckWinVersion()
    Dim lngVersion As Long
    
    mbln注销 = False
    lngVersion = GetVersion()
    If ((lngVersion And &H80000000) = 0) Then
        mlngWin32 = WINDOWSNT
    Else
        mlngWin32 = WINDOWS95
    End If
End Sub

Public Sub ShutDown(ByVal blnCloseWin As Boolean)
    If mbln注销 Then Exit Sub
    If Not blnCloseWin Then Exit Sub
    If mlngWin32 = WINDOWSNT Then
        'ExitWindowsEx lng关闭计算机及电源 Or EWX_FORCEIFHUNG, 0
        Call AdjustToken
        Call ExitWindowsEx(M_lng关闭计算机及电源 Or EWX_FORCE, 0)
    Else
        Call ExitWindowsEx(M_lng关闭计算机及电源 Or EWX_FORCE, 0)
    End If
End Sub

Public Function AdjustToken() As Boolean
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    'Set the error code of the last thread to zero using the'SetLast Error function
    SetLastError 0
    
    '得到当前进程的句柄
    hdlProcessHandle = GetCurrentProcess()
    If GetLastError <> 0 Then Exit Function
    
    '得到当前进程的权限句柄
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
    If GetLastError <> 0 Then Exit Function
     
    '找到关闭权限并赋给LUID
    'SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege
    'SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    
    tkp.PrivilegeCount = 1    ' One privilege to set
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    'Enable the shutdown privilege in the access token of this process
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
    If GetLastError <> 0 Then Exit Function
    
    AdjustToken = True
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim StrPass As String, strReturn As String, strSource As String, strTarget As String
    
    StrPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(StrPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function RunCommand(ByVal strCommand As String, Optional ByRef strErr As String, Optional ByVal blnCiper As Boolean, Optional ByVal lngWait As Long = INFINITE) As String
'功能：执行命令行，并获取命令行输出
    Dim piProc          As PROCESS_INFORMATION '进程信息
    Dim stStart         As STARTUPINFO '启动信息
    Dim saSecAttr       As SECURITY_ATTRIBUTES '安全属性
    Dim lnghReadPipe    As Long '读取管道句柄
    Dim lnghWritePipe   As Long '写入管道句柄
    Dim lngBytesRead    As Long '读出数据的字节数
    Dim strBuffer       As String * 256 '读取管道的字符串buffer
    Dim lngRet          As Long 'API函数返回值
    Dim lngRetPro       As Long
    Dim strlpOutputs    As String '读出的最终结果
    
    DoEvents
    On Error Resume Next
    '设置安全属性
    With saSecAttr
        .nLength = LenB(saSecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With
    
    '创建管道
    lngRet = CreatePipe(lnghReadPipe, lnghWritePipe, saSecAttr, 0)
    If lngRet = 0 Then
        strErr = "无法创建管道。" & GetLastDllErr()
        Exit Function
    End If
    '设置进程启动前的信息
    With stStart
        .Cb = LenB(stStart)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = lnghWritePipe '设置输出管道
        .hStdError = lnghWritePipe '设置错误管道
    End With
    '启动进程
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS进程以ipconfig.exe为例
    lngRetPro = CreateProcess(vbNullString, strCommand & vbNullChar, saSecAttr, saSecAttr, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, stStart, piProc)
    If lngRetPro = 0 Then
        strErr = "无法启动进程。" & GetLastDllErr()
        lngRet = CloseHandle(lnghWritePipe)
        lngRet = CloseHandle(lnghReadPipe)
        Exit Function
    Else
        '因为无需写入数据，所以先关闭写入管道。而且这里必须关闭此管道，否则将无法读取数据
        lngRet = CloseHandle(lnghWritePipe)
        WaitForSingleObject piProc.hProcess, lngWait
        Do
            lngRet = ReadFile(lnghReadPipe, strBuffer, 256, lngBytesRead, ByVal 0)
            If lngRet <> 0 Then
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            Else
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            End If
            DoEvents
        Loop While (lngRet <> 0) '当ret=0时说明ReadFile执行失败，已经没有数据可读了
        '读取操作完成，关闭各句柄
        lngRet = CloseHandle(lngRetPro)
        lngRet = CloseHandle(piProc.hProcess)
        lngRet = CloseHandle(piProc.hThread)
        lngRet = CloseHandle(lnghReadPipe)
    End If
    RunCommand = Replace(strlpOutputs, vbNullChar, "")
End Function

Public Function GetLastDllErr(Optional ByVal lngErr As Long) As String
    Dim strReturn As String
    If lngErr = 0 Then
        lngErr = GetLastError
    End If
    If lngErr = ERROR_EXTENDED_ERROR Then
        GetLastDllErr = GetWNetErr(lngErr)
    Else
        strReturn = String$(256, 32)
        FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lngErr, 0&, strReturn, Len(strReturn), ByVal 0
        strReturn = Trim(strReturn)
        GetLastDllErr = Replace(Replace(strReturn, Chr(10), ""), Chr(13), "")
    End If
End Function

Private Function GetWNetErr(ByVal lngErr As Long) As String
    Dim strErr As String * 256
    Dim strName As String * 256
    Dim lngRet As Long
    lngRet = WNetGetLastError(lngErr, strErr, Len(strErr), strName, Len(strName))
    GetWNetErr = Replace(Replace("[" & TruncZero(strName) & "]" & TruncZero(strErr), Chr(10), ""), Chr(13), "")
End Function

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function
