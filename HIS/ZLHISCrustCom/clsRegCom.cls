VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================================================================
'                0、模块变量与公共类型
'===========================================================================
Private Const INFINITE                           As Long = &HFFFF&
Private Const SYNCHRONIZE = &H100000
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Public Enum RegFileType
    RFT_NotReg = 0
    RFT_NormalReg = 1               '常规注册，自动识别.NET部件，.NET部件通过Regasm注册，其他通过调用DLLRegServer注册
    RFT_NETGAC = 2                  'NET程序集注册，通过gacutil注册到全局程序集缓存
    RFT_NETServer = 3               'NET服务注册，通过installUtil进行安装卸载。
    RFT_NETComReg = 4               '.NET Com部件注册，通过调用Regasm完成
End Enum

Private mstrNETRegToolsV4       As String
Private mstrNETRegToolsV2       As String

Private mclsPEReader            As New clsPEReader
Private mobjShell               As New clsShell
Private mobjFSO                 As New FileSystemObject
Private mobjRegistry            As New clsRegistry
Private mstrSysBitOpt           As String '64位系统的特殊处理

'===========================================================================
'                1、保存属性的变量
'===========================================================================
Private mstrGACUPath    As String

'===========================================================================
'                2、属性定义
'===========================================================================
Property Get GACUPath() As String
    GACUPath = mstrGACUPath
End Property

Property Let GACUPath(ByVal strGACUPath As String)
    mstrGACUPath = strGACUPath
End Property
'===========================================================================
'                3、公共方法
'===========================================================================
Public Function RegCom(ByVal strFileName As String, Optional ByRef strRegMsg As String, Optional ByVal intRegType As Integer = 1) As Boolean
'功能：注册文件
'参数：strFileName：文件路径
'返回：是否注册成功
    Dim strComLine  As String, strReturn    As String
    strRegMsg = ""
    If intRegType <> 0 Then
        If mclsPEReader.AnalyseFile(strFileName, strRegMsg) Then
            If intRegType = RFT_NETGAC Then
                If Me.GACUPath = "" Then
                    strRegMsg = "Com注册:未找到可用的程序集缓存装卸程序gacutil.exe"
                    Exit Function
                End If
            End If
            If mclsPEReader.IsActivexDLL And Not mclsPEReader.IsActivexEXE And intRegType = RFT_NormalReg And Not mclsPEReader.IsNetLibrary Then
                RegCom = DllRegisterServer(strFileName, strRegMsg, True)
            Else
                If intRegType = RFT_NETGAC Then
                    strComLine = Me.GACUPath & " /i " & strFileName
                ElseIf intRegType = RFT_NETServer Then
                    strComLine = GetNETRegTools(mclsPEReader.CLRVer, 2) & " " & strFileName
                ElseIf mclsPEReader.IsNetLibrary Then
                    strComLine = GetNETRegTools(mclsPEReader.CLRVer, 1) & " " & strFileName & " /tlb:" & mobjFSO.GetBaseName(strFileName) & ".tlb" & " /codebase"
                ElseIf mclsPEReader.IsActivexEXE Then
                    strComLine = strFileName & " /RegServer"
                End If
                If strComLine <> "" Then
                    If mobjShell.Run(strComLine, strReturn, strRegMsg, 120000) Then
                        RegCom = IsRegOk(IIf(intRegType = RFT_NormalReg, IIf(mclsPEReader.IsNetLibrary, 1, 0), intRegType), strReturn)
                    Else
                        RegCom = False
                    End If
                    If Not RegCom And intRegType = RFT_NormalReg And mclsPEReader.IsNetLibrary Then
                        strRegMsg = "注册失败，可能缺失.NET框架4.0+，已经忽略。"
                        RegCom = True
                    End If
                Else
                    strRegMsg = "该文件自动注册类型错误或不需要注册，已经忽略。"
                    RegCom = True
                End If
            End If
        Else
            strRegMsg = "PE分析:" & strRegMsg
        End If
    Else
        RegCom = True
    End If
End Function

Public Function UnRegCom(ByVal strFileName As String, Optional ByRef strRegMsg As String, Optional ByVal intRegType As Integer = 1) As Boolean
'功能：反注册文件
'参数：strFileName：文件路径
'返回：是否反注册成功
    Dim strComLine  As String, strReturn    As String
    strRegMsg = ""
    If intRegType <> 0 Then
        If mclsPEReader.AnalyseFile(strFileName, strRegMsg) Then
            If intRegType = RFT_NETGAC Then
                If Me.GACUPath = "" Then
                    strRegMsg = "Com注册:未找到可用的程序集缓存装卸程序gacutil.exe"
                    Exit Function
                End If
            End If
            If mclsPEReader.IsActivexDLL And Not mclsPEReader.IsActivexEXE And intRegType = RFT_NormalReg And Not mclsPEReader.IsNetLibrary Then
                UnRegCom = DllRegisterServer(strFileName, strRegMsg, False)
            Else
                If intRegType = RFT_NETGAC Then
                    strComLine = Me.GACUPath & " /u " & mobjFSO.GetBaseName(strFileName)
                ElseIf intRegType = RFT_NETServer Then
                    strComLine = GetNETRegTools(mclsPEReader.CLRVer, 2) & " " & strFileName & " -U"
                ElseIf mclsPEReader.IsNetLibrary Then
                    strComLine = GetNETRegTools(mclsPEReader.CLRVer, 1) & " " & strFileName & " /UNREGISTER"
                ElseIf mclsPEReader.IsActivexEXE Then
                    strComLine = strFileName & "  /UnRegServer"
                End If
                If strComLine <> "" Then
                    If mobjShell.Run(strComLine, strReturn, strRegMsg, 120000) Then
                        UnRegCom = IsRegOk(IIf(intRegType = RFT_NormalReg, IIf(mclsPEReader.IsNetLibrary, 1, 0), intRegType), strReturn, False)
                    Else
                        UnRegCom = False
                    End If
                    If Not UnRegCom And intRegType = RFT_NormalReg And mclsPEReader.IsNetLibrary Then
                        strRegMsg = "反注册失败，可能缺失.NET框架4.0+，已经忽略。"
                        UnRegCom = True
                    End If
                Else
                    UnRegCom = True
                End If
            End If
        Else
            strRegMsg = "PE分析:" & strRegMsg
        End If
    Else
        UnRegCom = True
    End If
End Function
'===========================================================================
'                4、私有方法
'===========================================================================
'****************1、DllRegistrServer注册与反注册**************************
Private Function DllRegisterServer(ByVal strFileName As String, Optional ByRef strRegMsg As String, Optional ByVal blnReg As Boolean = True) As Boolean
'功能：调用DLL内部函数"DllRegisterServer"来注册部件或调用DLL内部函数"DllUnregisterServer"来注册部件
'参数：strFileName：注册的部件
'      blnReg:True-注册部件，False-取消部件注册
    Dim lngLibID    As Long, lngProcAdrr    As Long
    Dim lngThreadID As Long, lngThreadHwnd  As Long, blnSuccess  As Boolean
    Dim lngExitCode As Long
    Dim blnReturn   As Boolean
    
    On Error GoTo ErrorH
    '没有提供注册的部件
    If strFileName = "" Then
        strRegMsg = "未找到文件"
    Else
        lngLibID = LoadLibrary(strFileName)
        If lngLibID = 0 Then '将DLL加载到内存失败
            strRegMsg = "将DLL加载到内存失败"
        Else
            lngProcAdrr = GetProcAddress(lngLibID, IIf(blnReg, "DllRegisterServer", "DllUnregisterServer"))
            If lngProcAdrr = 0 Then '没有注册服务器，则不是标准的ActiveDLL
                strRegMsg = "获取注册服务器失败"
            Else
                lngThreadHwnd = CreateThread(ByVal 0, 0, lngProcAdrr, 0, 0, lngThreadID)
                If lngThreadHwnd <> 0 Then
                    blnSuccess = (WaitForSingleObject(lngThreadHwnd, INFINITE) = WAIT_OBJECT_0)
                    If blnSuccess Then
                        blnReturn = True
                    Else
                        strRegMsg = GetLastDllErr
                        Call GetExitCodeThread(lngThreadHwnd, lngExitCode)
                        Call ExitThread(lngExitCode)
                    End If
                Else
                    strRegMsg = "注册服务器进程创建失败"
                End If
            End If
        End If
    End If
ExitRoutine:
    DllRegisterServer = blnReturn
    '关闭进程
    If lngThreadHwnd <> 0 Then Call CloseHandle(lngThreadHwnd)
    '释放DLL库
    If lngLibID <> 0 Then Call FreeLibrary(lngLibID)
    Exit Function
ErrorH:
    strRegMsg = Err.Description
    GoTo ExitRoutine
End Function

Private Function GetNETRegTools(ByVal strNetVer As String, Optional ByVal intType As Integer) As String
'功能：获取RegAsm文件路径
'参数：strNetVer=.NET版本号
'返回：对应的RegAsm文件
'Private mstrNETRegToolsV4 As String
'Private mstrNETRegToolsV2 As String
    Dim strSubKeyValues As Variant, strRegAsmRoot As String
    Dim strToolsName As String
    
    If intType = 1 Then
        strToolsName = "Regasm.exe"
    ElseIf intType = 2 Then
        strToolsName = "InstallUtil.exe"
    End If
    If strNetVer <> "" And mstrNETRegToolsV4 = "" And mstrNETRegToolsV2 = "" Then
        If mobjRegistry.GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\.NETFramework", "InstallRoot", strRegAsmRoot) Then
            strSubKeyValues = mobjRegistry.GetAllKeyValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\.NETFramework\Policy\v4.0")
            If TypeName(strSubKeyValues) = "Empty" Then
                mstrNETRegToolsV4 = " " '标记已经读取数据
            Else
                If Join(strSubKeyValues) <> "" Then
                    mstrNETRegToolsV4 = strRegAsmRoot + "v4.0." + strSubKeyValues(0)
                Else
                    mstrNETRegToolsV4 = ""
                End If
            End If
            strSubKeyValues = mobjRegistry.GetAllKeyValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\.NETFramework\Policy\v2.0")
            If TypeName(strSubKeyValues) = "Empty" Then
                mstrNETRegToolsV2 = " " '标记已经读取数据
            Else
                If Join(strSubKeyValues) <> "" Then
                    mstrNETRegToolsV2 = strRegAsmRoot + "v2.0." + strSubKeyValues(0)
                Else
                    mstrNETRegToolsV2 = ""
                End If
            End If
        End If
    End If
    
    If InStr(strNetVer, "v4") > 0 And Trim(mstrNETRegToolsV4) <> "" Then
        GetNETRegTools = mstrNETRegToolsV4
    ElseIf InStr(strNetVer, "v2") > 0 And Trim(mstrNETRegToolsV2) <> "" Then
        GetNETRegTools = mstrNETRegToolsV2
    Else
        GetNETRegTools = ""
    End If
    If strToolsName <> "" Then
        If GetNETRegTools = "" Then
            GetNETRegTools = strToolsName
        Else
            GetNETRegTools = GetNETRegTools & "\" & strToolsName
        End If
    End If
End Function

Private Function IsRegOk(ByVal intNETType As Integer, ByRef strReSult As String, Optional ByVal blnReg As Boolean = True) As Boolean
'判断NET部件是否注册或反注册成功
'intNETDLLType=0-非NET部件 ,1-NETCOM ,2-NET全局程序集缓存 3-NET服务
'strReSult=返回信息
'blnReg=True-注册，False-反注册
'返回：是否注册成功
    Dim strTmp As String, strKeyWordCN As String, strKeyWordEN As String
    Dim lngPos As Long
    
    IsRegOk = True
    If strReSult = "" Then Exit Function
    strTmp = UCase(strReSult)
    Select Case intNETType
        Case 0
        Case 1
            strTmp = Replace(strTmp, "", "")
            If blnReg Then
                strKeyWordCN = "成功注册了类型"
                strKeyWordEN = "TYPES REGISTERED SUCCESSFULLY"
            Else
                strKeyWordCN = "成功注销了类型"
                strKeyWordEN = "TYPES UN-REGISTERED SUCCESSFULLY"
            End If
        Case 2
            If blnReg Then
                strKeyWordCN = "程序集已成功添加到缓存中"
                strKeyWordEN = "ASSEMBLY SUCCESSFULLY ADDED TO THE CACHE"
            Else
                strKeyWordCN = "失败次数 = 0"
                strKeyWordEN = "NUMBER OF FAILURES = 0"
            End If
        Case 3
            If blnReg Then
                strKeyWordCN = "已完成事务处理安装。"
                strKeyWordEN = "THE TRANSACTED INSTALL HAS COMPLETED."
            Else
                strKeyWordCN = "卸载完成。"
                strKeyWordEN = "THE UNINSTALL HAS COMPLETED."
            End If
    End Select
    If strKeyWordCN <> "" Then
        If InStr(strTmp, strKeyWordCN) > 0 Then
            Exit Function
        End If
    End If
    If strKeyWordEN <> "" Then
        If InStr(strTmp, strKeyWordEN) > 0 Then
            Exit Function
        End If
    End If
    IsRegOk = False
End Function

Private Sub Class_Terminate()
    Set mclsPEReader = Nothing
    Set mobjShell = Nothing
    Set mobjFSO = Nothing
    Set mobjRegistry = Nothing
End Sub
