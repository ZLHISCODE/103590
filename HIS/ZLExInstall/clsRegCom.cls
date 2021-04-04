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
'*************************************************************************
'**模 块 名：clsRegCom
'**创 建 人：lshuo
'**日    期：2016/6/15
'**描    述：本类模块提供了两种方法来注册Com组件
'**          1、DllRegisterServer注册
'**          2、直接操作注册表注册
'**版    本：V1.0
'*************************************************************************
Option Explicit
'===========================================================================
'                0、模块变量与公共类型
'===========================================================================
Const LOAD_WITH_ALTERED_SEARCH_PATH As Long = &H8
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
'RegSvr注册，部件本身的内部函数注册
Private Const NOERRORS As Long = 0
Private Const SYNCHRONIZE = &H100000
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Private Const INFINITE = -1&
'注册表直接注册与反注册
'注册表关键字根类型
Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '记录Windows操作系统中所有数据文件的格式和关联信息，主要记录不同文件的文件名后缀和与之对应的应用程序。其下子键可分为两类，一类是已经注册的各类文件的扩展名，这类子键前面都有一个“。”；另一类是各类文件类型有关信息。
    HKEY_CURRENT_USER = &H80000001 '此根键包含了当前登录用户的用户配置文件信息。这些信息保证不同的用户登录计算机时，使用自己的个性化设置，例如自己定义的墙纸、自己的收件箱、自己的安全访问权限等。
    HKEY_LOCaL_MaCHINE = &H80000002 '此根键包含了当前计算机的配置数据，包括所安装的硬件以及软件的设置。这些信息是为所有的用户登录系统服务的。它是整个注册表中最庞大也是最重要的根键！
    HKEY_USERS = &H80000003 '此根键包括默认用户的信息（Default子键）和所有以前登录用户的信息。
    HKEY_PERFORMANCE_DATA = &H80000004 '在Windows NT/2000/XP注册表中虽然没有HKEY_DYN_DATA键，但是它却隐藏了一个名为“HKEY_ PERFOR MANCE_DATA”键。所有系统中的动态信息都是存放在此子键中。系统自带的注册表编辑器无法看到此键
    HKEY_CURRENT_CONFIG = &H80000005  '此根键实际上是HKEY_LOCAL_MACHINE中的一部分，其中存放的是计算机当前设置，如显示器、打印机等外设的设置信息等。它的子键与HKEY_LOCAL_ MACHINE\ Config\0001分支下的数据完全一样。
    HKEY_DYN_DATA = &H80000006 '此根键中保存每次系统启动时，创建的系统配置和当前性能信息。这个根键只存在于Windows 98中。
End Enum

'注册表数据类型
Private Enum REGValueType
    REG_NONE = 0                       ' No value type
    REG_SZ = 1 'Unicode空终结字符串
    REG_EXPAND_SZ = 2 'Unicode空终结字符串
    REG_BINARY = 3 '二进制数值
    REG_DWORD = 4 '32-bit 数字
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' 二进制数值串
End Enum
'打开错误
Private Enum REGErr
    ERROR_SUCCESS = &H0
    ERROR_FILE_NOT_FOUND = &H2 'The system cannot find the file specified
    ERROR_BADDB = 1009&
    ERROR_BADKEY = 1010&
    ERROR_CANTOPEN = 1011&
    ERROR_CANTREAD = 1012&
    ERROR_CANTWRITE = 1013&
    ERROR_OUTOFMEMORY = 14&
    ERROR_INVALID_PARAMETER = 87&
    ERROR_ACCESS_DENIED = 5&
    ERROR_NO_MORE_ITEMS = 259&
    ERROR_MORE_DATA = 234&
End Enum
'注册表访问权
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Private Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

Private Type SECURITY_ATTRIBUTES
   nLength                  As Long
   lpSecurityDescriptor     As Long
   bInheritHandle           As Boolean
End Type
Private Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal uloptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private mlngHMRoot          As Long 'Local:Classes根目录
Private mlngHMClass         As Long 'Local:CLSID主目录
Private mlngHMType          As Long 'Local:TypeID主目录
Private mlngHMInterface     As Long 'Local:IID主目录

Private mblnReadKey         As Boolean
Private mstrNETRegToolsV4        As String
Private mstrNETRegToolsV2        As String
Private mclsPEReader        As New clsPEReader
Private mstrSysBitOpt       As String '64位系统的特殊处理
'===========================================================================
'                1、保存属性的变量
'===========================================================================


'===========================================================================
'                2、属性定义
'===========================================================================


'===========================================================================
'                3、公共方法
'===========================================================================
Public Function RegCom(ByVal strFileName As String, Optional rftRegType As RegFileType, Optional ByRef strRegMsg As String) As Boolean
'功能：注册文件
'参数：strFileName：文件路径
'返回：是否注册成功
    Dim strComline  As String, strReturn    As String
    Dim strType     As String, intNETType   As Integer
    Dim strTmpMsg   As String
    On Error Resume Next
    intNETType = 0: strRegMsg = ""
    If rftRegType = RFT_NotReg Then
        RegCom = True
    Else
        If mclsPEReader.LoadPEFile(strFileName, strRegMsg) Then
            If mclsPEReader.IsNetLibrary Then
                If rftRegType = RFT_NETGAC Then 'NET全局程序集缓存
                    If gstrGACPath = "" Then
                        RegCom = False
                        strRegMsg = "程序集添加到缓存失败(" & strFileName & "):未找到可用的程序集缓存装卸程序gacutil.exe"
                        gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
                        Exit Function
                    End If
                    strComline = gstrGACPath
                    strComline = strComline & " /i " & strFileName
                    strType = "程序集添加到缓存"
                    intNETType = 2
                Else
                    strComline = GetNETRegTools(mclsPEReader.CLRVer)
                    If strComline = "" Then
                        strComline = "Regasm.exe"
                    Else
                        strComline = strComline & "\Regasm.exe"
                    End If
                    strComline = strComline & " " & strFileName & " /tlb:" & gobjFSO.GetBaseName(strFileName) & ".tlb" & " /codebase"
                    strType = "类型注册"
                    intNETType = 1
                End If
                strReturn = RunCommand(strComline, strRegMsg, , 120000)
                If strRegMsg = "" Then
                    RegCom = IsRegOk(intNETType, strReturn)
                    If RegCom Then
                        gobjTrace.WriteInfo "RegCom", "注册结果", strType & "成功(" & strFileName & ")"
                    Else
                        If rftRegType <> RFT_NETGAC Then
                            strRegMsg = strType & "失败，可能缺失.NET框架4.0+，已经忽略(" & strFileName & "):" & strReturn
                            gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
                            RegCom = True
                        Else
                            strRegMsg = strType & "失败(" & strFileName & "):" & strReturn
                            gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
                        End If
                    End If
                Else
                    If rftRegType <> RFT_NETGAC Then
                        strRegMsg = strType & "失败，可能缺失.NET框架4.0+，已经忽略(" & strFileName & "):" & strReturn
                        gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
                        RegCom = True
                    Else
                        strRegMsg = strType & "失败(" & strFileName & "):" & strReturn
                        gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
                    End If
                End If
            ElseIf mclsPEReader.IsDLL Then
                RegCom = ComRegeditReg(strFileName, strRegMsg, , rftRegType = RFT_VBComReg)
                If RegCom Then
                    gobjTrace.WriteInfo "RegCom", "注册结果", "DLL注册成功(" & strFileName & ")"
                Else
                    strRegMsg = "DLL注册失败(" & strFileName & "):" & strRegMsg
                    gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
                End If
            ElseIf mclsPEReader.IsActivexEXE Then
                RegCom = ComRegeditReg(strFileName, strRegMsg, , rftRegType = RFT_VBComReg)
                If RegCom Then
                    gobjTrace.WriteInfo "RegCom", "注册结果", "ActivexEXE注册成功(" & strFileName & ")"
                Else
                    strRegMsg = "ActivexEXE注册失败(" & strFileName & "):" & strRegMsg
                    gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
                End If
            End If
        ElseIf rftRegType = RFT_NETServer Then 'NET写的服务程序注册
            strComline = GetNETRegTools(mclsPEReader.CLRVer)
            If strComline = "" Then
                strComline = "InstallUtil.exe"
            Else
                strComline = strComline & "\InstallUtil.exe"
            End If
            strComline = strComline & " " & strFileName
            strType = "事务处理安装"
            intNETType = 3
            strRegMsg = ""
            strReturn = RunCommand(strComline, strRegMsg, , 120000)
            If strRegMsg = "" Then
                RegCom = IsRegOk(intNETType, strReturn)
                If RegCom Then
                    gobjTrace.WriteInfo "RegCom", "注册结果", strType & "成功(" & strFileName & ")"
                Else
                    strRegMsg = strType & "失败(" & strFileName & "):" & strReturn
                    gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
                End If
            Else
                strRegMsg = strType & "失败(" & strFileName & "):" & strReturn
                gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
            End If
        Else
            'C:\Windows\System32\QEDIT.DLL判断不是DLL,不是有效PE,因为没有MZ,PE判断也不是DLL，但是存在注册入口
            If UCase(strFileName) Like "*.DLL" Then
                RegCom = ComRegeditReg(strFileName, strTmpMsg, , rftRegType = RFT_VBComReg)
                If RegCom Then
                    gobjTrace.WriteInfo "RegCom", "注册结果", "特殊DLL注册成功(" & strFileName & ")"
                    strRegMsg = ""
                    Exit Function
                End If
            End If
            RegCom = True
            strRegMsg = "无法注册(" & strFileName & "):" & strRegMsg
            gobjTrace.WriteInfo "RegCom", "注册结果", strRegMsg
        End If
    End If
End Function

Public Function UnRegCom(ByVal strFileName As String, Optional rftRegType As RegFileType = RFT_NormalReg, Optional ByRef strRegMsg As String) As Boolean
'功能：反注册文件
'参数：strFileName：文件路径
'返回：是否反注册成功
    Dim strComline  As String, strReturn    As String
    Dim strType     As String, intNETType   As Integer
    Dim strTmpMsg   As String

    On Error Resume Next
    intNETType = 0: strRegMsg = ""
    If rftRegType = RFT_NotReg Then
        UnRegCom = True
    Else
        If mclsPEReader.LoadPEFile(strFileName, strRegMsg) Then
            If mclsPEReader.IsNetLibrary Then
                If rftRegType = RFT_NETGAC Then 'NET全局程序集缓存
                    If gstrGACPath = "" Then
                        UnRegCom = False
                        strRegMsg = "程序集卸载失败(" & strFileName & "):未找到可用的程序集缓存装卸程序gacutil.exe"
                        gobjTrace.WriteInfo "UnRegCom", "反注册结果", strRegMsg
                        Exit Function
                    End If
                    strComline = gstrGACPath
                    strComline = strComline & " /u " & gobjFSO.GetBaseName(strFileName)
                    strType = "程序集卸载"
                    intNETType = 2
                Else
                    strComline = GetNETRegTools(mclsPEReader.CLRVer)
                    If strComline = "" Then
                        strComline = "Regasm.exe"
                    Else
                        strComline = strComline & "\Regasm.exe"
                    End If
                    strComline = strComline & " " & strFileName & "  /UNREGISTER"
                    strType = "类型卸载"
                    intNETType = 1
                End If
                strReturn = RunCommand(strComline, strRegMsg, , 120000)
                If strRegMsg = "" Then
                    UnRegCom = IsRegOk(intNETType, strReturn, False)
                    If UnRegCom Then
                        gobjTrace.WriteInfo "UnRegCom", "反注册结果", strType & "成功(" & strFileName & ")"
                    Else
                        strRegMsg = strType & "失败(" & strFileName & "):" & strRegMsg
                        gobjTrace.WriteInfo "UnRegCom", "反注册结果", strRegMsg
                    End If
                Else
                    strRegMsg = strType & "失败(" & strFileName & "):" & strRegMsg
                    gobjTrace.WriteInfo "UnRegCom", "反注册结果", strRegMsg
                End If
            ElseIf mclsPEReader.IsDLL Then
                UnRegCom = ComRegeditReg(strFileName, strRegMsg, False, rftRegType = RFT_VBComReg)
                If UnRegCom Then
                    gobjTrace.WriteInfo "UnRegCom", "反注册结果", "DLL反注册成功(" & strFileName & ")"
                Else
                    strRegMsg = "DLL反注册失败(" & strFileName & "):" & strRegMsg
                    gobjTrace.WriteInfo "UnRegCom", "反注册结果", strRegMsg
                End If
            ElseIf mclsPEReader.IsActivexEXE Then
                UnRegCom = ComRegeditReg(strFileName, strRegMsg, False, rftRegType = RFT_VBComReg)
                If UnRegCom Then
                    gobjTrace.WriteInfo "UnRegCom", "反注册结果", "ActivexEXE反注册成功(" & strFileName & ")"
                Else
                    strRegMsg = "ActivexEXE反注册失败(" & strFileName & "):" & strRegMsg
                    gobjTrace.WriteInfo "UnRegCom", "反注册结果", strRegMsg
                End If
            End If
        ElseIf rftRegType = RFT_NETServer Then 'NET写的服务程序注册
            strComline = GetNETRegTools(mclsPEReader.CLRVer)
            If strComline = "" Then
                strComline = "InstallUtil.exe"
            Else
                strComline = strComline & "\InstallUtil.exe"
            End If
            strComline = strComline & " " & strFileName & " -U"
            strType = "事务处理卸载"
            intNETType = 3
            strRegMsg = ""
            strReturn = RunCommand(strComline, strRegMsg, , 120000)
            If strRegMsg = "" Then
                UnRegCom = IsRegOk(intNETType, strReturn)
                If UnRegCom Then
                    gobjTrace.WriteInfo "UnRegCom", "反注册结果", strType & "成功(" & strFileName & ")"
                Else
                    strRegMsg = strType & "失败(" & strFileName & "):" & strReturn
                    gobjTrace.WriteInfo "UnRegCom", "反注册结果", strRegMsg
                End If
            Else
                strRegMsg = strType & "失败(" & strFileName & "):" & strRegMsg
                gobjTrace.WriteInfo "UnRegCom", "反注册结果", strRegMsg
            End If
        Else
            'C:\Windows\System32\QEDIT.DLL判断不是DLL,不是有效PE,因为没有MZ,PE判断也不是DLL，但是存在注册入口
            If UCase(strFileName) Like "*.DLL" Then
                UnRegCom = ComRegeditReg(strFileName, strTmpMsg, False, rftRegType = RFT_VBComReg)
                If UnRegCom Then
                    gobjTrace.WriteInfo "UnRegCom", "反注册结果", "特殊DLL反注册成功(" & strFileName & ")"
                    strRegMsg = ""
                    Exit Function
                End If
            End If
            strRegMsg = "无法反注册(" & strFileName & "):" & strRegMsg
            gobjTrace.WriteInfo "UnRegCom", "反注册结果", strRegMsg
            UnRegCom = True
        End If
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
        lngLibID = LoadLibraryEx(strFileName, 0, LOAD_WITH_ALTERED_SEARCH_PATH)
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

Private Function ExeRegisterServer(ByVal strFileName As String, Optional ByRef strRegMsg As String, Optional ByVal blnReg As Boolean = True) As Boolean
'功能：Exe的命令来进行Active EXE的注册与反注册
'参数：strFileName：注册的部件
'      blnReg:True-注册部件，False-取消部件注册
    Dim lngTaskID As Double, lngProcID As Long
    Dim lngReturn As Long

    lngTaskID = Shell(strFileName & IIf(blnReg, " /RegServer", " /UnRegServer"), vbNormalFocus)
    If lngTaskID = 0 Then Exit Function
    lngProcID = OpenProcess(SYNCHRONIZE, False, lngTaskID)
    If lngProcID = 0 Then Exit Function
    lngReturn = WaitForSingleObject(lngProcID, INFINITE)
    lngReturn = CloseHandle(lngProcID)
End Function

Private Function Regsvr32DLL(ByVal strFileName As String, Optional ByRef strRegMsg As String, Optional ByVal blnReg As Boolean = True) As Boolean
    Dim strComline As String, strReturn As String
    strRegMsg = ""
    strComline = gstSysPath & "\regsvr32 " & IIf(blnReg, "/S ", "/U /S ") & strFileName
    strReturn = RunCommand(strComline, strRegMsg, , 120000)
    Regsvr32DLL = strReturn = "" And strRegMsg = ""
End Function

'****************2、注册表直接注册与反注册**************************
Private Function ComRegeditReg(ByVal strFileName As String, Optional ByRef strRegMsg As String, Optional ByVal blnReg As Boolean = True, Optional ByVal blnVB As Boolean) As Boolean
    Dim objinfo         As TypeLibInfo, strTypeName As String
    Dim strProxyClsid   As String, strProxyClsidCur As String
    Dim objTypeInfo     As TypeInfo
    Dim arrVersions     As Variant, i   As Long
    Dim strFileType     As String, intType          As Integer
    Dim strCurVer       As String
    Dim blnOk           As Boolean, blnNewStyle     As Boolean

    strRegMsg = ""
    strFileType = UCase(Right(strFileName, 4))
    If strFileType = ".DLL" Then
        intType = 1
        strProxyClsid = "{00020424-0000-0000-C000-000000000046}"
    ElseIf strFileType = ".OCX" Then
        intType = 2
        strProxyClsid = "{00020424-0000-0000-C000-000000000046}"
    ElseIf strFileType = ".EXE" Then
        intType = 3
        strProxyClsid = "{00020424-0000-0000-C000-000000000046}"
    End If
    gobjTrace.WriteInfo "ComRegeditReg", "注册类型", Decode(intType, 1, "Active DLL", 2, "Active EXE", 3, "OCX")
    If intType = 0 Then
        strRegMsg = "部件非PE文件(.DLL\.OCX\.EXE)，无法注册"
        gobjTrace.WriteInfo "ComRegeditReg", "注册信息", strRegMsg
        Exit Function
    End If
    blnNewStyle = blnVB And Not blnReg
    blnNewStyle = blnVB
    If Not UCase(gobjFSO.GetFileName(strFileName)) Like "ZL*" Then
        blnNewStyle = False
    End If
     '打开各个注册表写入口
    If Not mblnReadKey And blnNewStyle Then
        If Not GetRegKey(strRegMsg) Then
            gobjTrace.WriteInfo "ComRegeditReg", "打开注册表出错", "由于无法打开注册表采用老方式注册。"
            blnNewStyle = False
        End If
    End If

    If blnNewStyle Then
        Err.Clear: On Error Resume Next
        '获取TypeID等信息
        Set objinfo = TypeLibInfoFromFile(gobjFSO.GetFile(strFileName).ShortPath)
        If Err.Number <> 0 Then
            gobjTrace.WriteInfo "ComRegeditReg", "加载类型库", "由于加载类型库失败，采用老方式注册。"
            blnNewStyle = False
            Err.Clear
        Else
            strTypeName = objinfo.Interfaces(1).Name
            If Err.Number <> 0 Then Err.Clear
            If Not strTypeName Like "_*" Then
                blnNewStyle = False
            End If
        End If
    End If
    If Not blnNewStyle Then
        gobjTrace.WriteInfo "ComRegeditReg", "注册方式", "老方式"
        If intType = 3 Then
            Call ExeRegisterServer(strFileName, strRegMsg, blnReg)
        Else
            Call DllRegisterServer(strFileName, strRegMsg, blnReg)
            'OIP.DLL等其他DLL,LOAD失败无法注册，可用RegSvr32注册成功
            If Not IsRegOk(0, strRegMsg, blnReg) Then
                Call Regsvr32DLL(strFileName, strRegMsg, blnReg)
            End If
        End If
        If Not IsRegOk(0, strRegMsg, blnReg) Then
            ComRegeditReg = False
        Else
            ComRegeditReg = True
            strRegMsg = ""
        End If
        Exit Function
    End If
    '1、TypLIB的注册信息
    '[TID\Version\HELPDIR]@="文件路径"
    'EXE,DLL
    '[TID\Version\FLAGS]@="0"
    'OCX
    '[TID\Version\FLAGS]@="2"
    '[TID\Version\0\win32]@="文件路径"
    '[TID\Version]@="HelpString"
    '2、ClsID注册信息
    '[CLSID]@="ProgID"

    'EXE,DLL
    '[CLSID\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}]--409 Automation Objects
    'OCX
    '[CLSID\Control]@=""
    '[CLSID\MiscStatus]@="0"
    '[CLSID\MiscStatus\1]@="229777"(内部资源）
    '[CLSID\ToolboxBitmap32]@="文件路径，资源ID"(内部资源,vb从30000开始）
    '[CLSID\Implemented Categories\{0DE86A52-2BAA-11CF-A229-00AA003D7352}]--409 Class implements IPersistStorage
    '[CLSID\Implemented Categories\{0DE86A53-2BAA-11CF-A229-00AA003D7352}]--409 Class implements IPersistStreamInit
    '[CLSID\Implemented Categories\{0DE86A57-2BAA-11CF-A229-00AA003D7352}]--409 Class implements IPersistPropertyBag
    '[CLSID\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}]--409 Controls:800 Control
    '[CLSID\Required Categories\{D40C2700-FFA1-11CF-8234-00AA00C1AB85}]未知

    'OCX,DLL
    '[CLSID\InprocServer32]@="文件路径"
    '[CLSID\InprocServer32]"ThreadingModel"="Apartment"'单套间，VB不支持多线程
    'EXE
    '[CLSID\LocalServer32]@="文件路径"

    '[CLSID\ProgID]@="ProgID"
    '[CLSID\Programmable]
    '[CLSID\TypeLib]@="TypeID"
    '[CLSID\Version]@="TIDVersion"

    '3、IID注册信息
    '[IID]@="CLSName""
    '[IID\ProxyStubClsid]@="00020424-0000-0000-C000-000000000046"(C:\Windows\system32\oleaut32.dll的PSOAInterface）
    '[IID\ProxyStubClsid32]@="00020424-0000-0000-C000-000000000046"(C:\Windows\system32\oleaut32.dll的PSOAInterface）
    '[IID\TypeLib]@="TypeID"
    '[IID\TypeLib]"Version"="TypeVersion"
    '4、ProgID注册信息
    '[ProgID\Clsid]@="CLSID"

    '说明
    'objinfo.Interfaces可能包含老接口
    '如ZL9comlib的第六个接口为clsComLib___v0
    '该接口的objinfo.Interfaces.NamedItem ("clsComLib___v0").VTableInterface.Name =_clsComLib为新接口
    '产生接口转向问题
    '[IID]@="CLSName""
    '[IID\ProxyStubClsid]@="00020424-0000-0000-C000-000000000046"(C:\Windows\system32\oleaut32.dll的PSOAInterface）
    '[IID\ProxyStubClsid32]@="00020424-0000-0000-C000-000000000046"(C:\Windows\system32\oleaut32.dll的PSOAInterface）
    '[IID\\Forward]@="NewIID"
    'Root节点是SOftWare节点的快捷方式，实际不存在，删除，写入SOftWare会自动产生Root节点
    On Error GoTo ErrH
    blnOk = True
    If blnReg Then
        strCurVer = objinfo.MajorVersion & "." & objinfo.MinorVersion
        arrVersions = GetAllSubKey(mlngHMType, objinfo.Guid)
        '重新写入TypeID，SoftWare节点写入
        blnOk = blnOk And SetRegValue(mlngHMType, objinfo.Guid)
        blnOk = blnOk And SetRegValue(mlngHMType, objinfo.Guid & "\" & strCurVer, "", objinfo.HelpString)
        blnOk = blnOk And SetRegValue(mlngHMType, objinfo.Guid & "\" & strCurVer & "\HELPDIR", "", gobjFSO.GetParentFolderName(strFileName))
        blnOk = blnOk And SetRegValue(mlngHMType, objinfo.Guid & "\" & strCurVer & "\FLAGS", "", IIf(intType = 2, 2, 0))
        blnOk = blnOk And SetRegValue(mlngHMType, objinfo.Guid & "\" & strCurVer & "\0")
        blnOk = blnOk And SetRegValue(mlngHMType, objinfo.Guid & "\" & strCurVer & "\0\win32", "", strFileName)
        For Each objTypeInfo In objinfo.TypeInfos
            If objTypeInfo.TypeKind = TKIND_COCLASS Then
                strTypeName = objTypeInfo.Name
                blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid, "", objTypeInfo.Parent.Name & "." & strTypeName)
                blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\ProgID", "", objTypeInfo.Parent.Name & "." & strTypeName)
                blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\TypeLib", "", objinfo.Guid)
                blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\VERSION", "", strCurVer)
                'ProgID导入
                blnOk = blnOk And SetRegValue(mlngHMRoot, objTypeInfo.Parent.Name & "." & strTypeName, "", objTypeInfo.Parent.Name & "." & strTypeName)
                blnOk = blnOk And SetRegValue(mlngHMRoot, objTypeInfo.Parent.Name & "." & strTypeName & "\Clsid", "", objTypeInfo.Guid)
                If intType = 2 Then
                    '如下三项为次要信息，主要是音乐资源，空间图标资源
'                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\MiscStatus", "", "0")
'                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\MiscStatus\1", "", "229777")
'                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\ToolboxBitmap32", "", strFileName & ",30000")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Control", "", "")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Implemented Categories")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Implemented Categories\{0DE86A52-2BAA-11CF-A229-00AA003D7352}", "", "")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Implemented Categories\{0DE86A53-2BAA-11CF-A229-00AA003D7352}", "", "")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Implemented Categories\{0DE86A57-2BAA-11CF-A229-00AA003D7352}", "", "")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}", "", "")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Required Categories")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Required Categories\{D40C2700-FFA1-11CF-8234-00AA00C1AB85}", "", "")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\InprocServer32", "", strFileName)
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\InprocServer32", "ThreadingModel", "Apartment")
                ElseIf intType = 1 Then
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Implemented Categories")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}", "", "")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\InprocServer32", "", strFileName)
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Programmable", "", "")
                ElseIf intType = 3 Then
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\LocalServer32", "", strFileName)
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Programmable", "", "")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Implemented Categories")
                    blnOk = blnOk And SetRegValue(mlngHMClass, objTypeInfo.Guid & "\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}", "", "")
                End If
            ElseIf objTypeInfo.TypeKind = TKIND_DISPATCH Then '接口
                strTypeName = TrimEx(objTypeInfo.Name, "_")
                If objTypeInfo.Name Like "__*" Then
                    strProxyClsidCur = "{00020420-0000-0000-C000-000000000046}"
                Else
                    strProxyClsidCur = strProxyClsid
                End If
                blnOk = blnOk And SetRegValue(mlngHMInterface, objTypeInfo.Guid, "", strTypeName)
                blnOk = blnOk And SetRegValue(mlngHMInterface, objTypeInfo.Guid & "\ProxyStubClsid", "", strProxyClsidCur)
                blnOk = blnOk And SetRegValue(mlngHMInterface, objTypeInfo.Guid & "\ProxyStubClsid32", "", strProxyClsidCur)
                blnOk = blnOk And SetRegValue(mlngHMInterface, objTypeInfo.Guid & "\TypeLib", "", objinfo.Guid)
                blnOk = blnOk And SetRegValue(mlngHMInterface, objTypeInfo.Guid & "\TypeLib", "Version", strCurVer)
            ElseIf objTypeInfo.TypeKind = TKIND_ALIAS Then  '别名
                strTypeName = TrimEx(objinfo.Interfaces.NamedItem(objTypeInfo.Name).VTableInterface.Name, "_")
                If objinfo.Interfaces.NamedItem(objTypeInfo.Name).VTableInterface.Name Like "__*" Then
                    strProxyClsidCur = "{00020420-0000-0000-C000-000000000046}"
                Else
                    strProxyClsidCur = strProxyClsid
                End If
                blnOk = blnOk And SetRegValue(mlngHMInterface, objTypeInfo.Guid, "", strTypeName)
                blnOk = blnOk And SetRegValue(mlngHMInterface, objTypeInfo.Guid & "\Forward", "", objinfo.Interfaces.NamedItem(objTypeInfo.Name).VTableInterface.Guid)
                blnOk = blnOk And SetRegValue(mlngHMInterface, objTypeInfo.Guid & "\ProxyStubClsid", "", strProxyClsidCur)
                blnOk = blnOk And SetRegValue(mlngHMInterface, objTypeInfo.Guid & "\ProxyStubClsid32", "", strProxyClsidCur)
            End If
        Next
    Else
        ''2、清理残留TypeID
        blnOk = blnOk And DeleteKeyValue(mlngHMType, objinfo.Guid)
        For Each objTypeInfo In objinfo.TypeInfos
            If objTypeInfo.TypeKind = TKIND_COCLASS Then
                strTypeName = objTypeInfo.Name
                'ProgID清理
                blnOk = blnOk And DeleteKeyValue(mlngHMRoot, objTypeInfo.Parent.Name & "." & strTypeName)
                blnOk = blnOk And DeleteKeyValue(mlngHMClass, objTypeInfo.Guid)
            ElseIf objTypeInfo.TypeKind = TKIND_DISPATCH Then '接口
'                blnOk = blnOk And DeleteKeyValue(mlngHMInterface, objTypeInfo.Guid)
            ElseIf objTypeInfo.TypeKind = TKIND_ALIAS Then  '别名
'                blnOk = blnOk And DeleteKeyValue(mlngHMInterface, objTypeInfo.Guid)
            End If
        Next
    End If
    strRegMsg = ""
    ComRegeditReg = blnOk
    Exit Function
ErrH:
    strRegMsg = Err.Description
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

'=========================================================================
'注册表操作
'=========================================================================
Private Function SetRegValue(ByVal lnghKey As Long, ByVal strSubKey As String, Optional ByVal strValueName As String, Optional ByVal strValue As String) As Boolean
'功能：设置注册表中指定位置的值
'参数：strSubKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'          strValue=变量值
'          strValueType=变量类型，默认为字符串
'返回：是否设置成功
    Dim lnghCurKey  As Long, lngReturn      As Long
    Dim lnghDepth   As Long                                  '
    Dim lpAttr      As SECURITY_ATTRIBUTES                   ' 注册表安全类型
    Dim lngLength   As Long
    Dim lngRetHwnd  As Long

    lpAttr.nLength = ActualLen(strValue)                      ' 设置安全属性为缺省值...
    lpAttr.lpSecurityDescriptor = 0
    lpAttr.bInheritHandle = False
    lngReturn = RegOpenKey(lnghKey, strSubKey, 0, KEY_WRITE, lnghCurKey)
    If lngReturn = ERROR_FILE_NOT_FOUND Then
        lngReturn = RegCreateKeyEx(lnghKey, strSubKey, 0, REG_SZ, REG_OPTION_NON_VOLATILE, KEY_WRITE, lpAttr, lnghCurKey, lnghDepth)
        If lngReturn = ERROR_SUCCESS Then
            Call CloseKey(lnghCurKey)
        Else
            gobjTrace.WriteInfo "SetRegValue", "创建子项失败", GetLastDllErr(lngReturn), "Key", "[" & GetRegPath(lnghKey) & "\" & strSubKey & "]""" & strValueName & """=""" & strValue & """"
        End If
        lngReturn = RegOpenKey(lnghKey, strSubKey, 0, KEY_WRITE, lnghCurKey)
    End If

    If lngReturn = ERROR_SUCCESS Then
        '校验句柄If
        SetRegValue = True
        If strValueName <> "" Or strValue <> "" Then
            lngLength = ActualLen(strValue)
            lngReturn = RegSetValue(lnghCurKey, strValueName, 0, REG_SZ, ByVal strValue, lngLength)
            If lngReturn <> ERROR_SUCCESS Then
                gobjTrace.WriteInfo "SetRegValue", "写入键值失败", GetLastDllErr(lngReturn), "Key", "[" & GetRegPath(lnghKey) & "\" & strSubKey & "]""" & strValueName & """=""" & strValue & """"
                SetRegValue = False
            End If
        End If
        Call CloseKey(lnghCurKey)
    Else
        gobjTrace.WriteInfo "SetRegValue", "打开注册表项失败", GetLastDllErr(lngReturn), "Key", "[" & GetRegPath(lnghKey) & "\" & strSubKey & "]""" & strValueName & """=""" & strValue & """"
    End If
    Exit Function
ErrH:
    gobjTrace.WriteInfo "SetRegValue", "未知错误", GetLastDllErr(lngReturn), "Key", "[" & GetRegPath(lnghKey) & "\" & strSubKey & "]""" & strValueName & """=""" & strValue & """"
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetAllSubKey(ByVal lnghKey As Long, ByVal strSubKey As String) As Variant
'功能:获取某项的所有子项
'返回：=子项数组
    Dim lnghCurKey As Long, lngRet As Long, strName As String, lngIdx As Long
    Dim arrSubKey As Variant

    arrSubKey = Array()
    lngIdx = 0
    lngRet = RegOpenKey(lnghKey, strSubKey, 0, KEY_READ, lnghCurKey)
    If lngRet = ERROR_SUCCESS Then
        Do
            strName = String(256, Chr(0))
            lngRet = RegEnumKey(lnghCurKey, lngIdx, strName, Len(strName))
            If lngRet = ERROR_SUCCESS Then
                ReDim Preserve arrSubKey(UBound(arrSubKey) + 1)
                arrSubKey(UBound(arrSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> ERROR_SUCCESS
        If lngRet <> ERROR_SUCCESS And lngRet <> ERROR_NO_MORE_ITEMS Then
            gobjTrace.WriteInfo "GetAllSubKey", "枚举子项失败", GetLastDllErr(lngRet), "Key", GetRegPath(lnghKey) & "\" & strSubKey
        End If
        Call CloseKey(lnghCurKey)
    Else
        If lngRet <> ERROR_FILE_NOT_FOUND Then
            gobjTrace.WriteInfo "GetAllSubKey", "打开注册表项失败", GetLastDllErr(lngRet), "Key", GetRegPath(lnghKey) & "\" & strSubKey
        End If
    End If
    GetAllSubKey = arrSubKey
End Function

Private Function GetAllKeyValue(ByVal lnghKey As Long, ByVal strSubKey As String) As Variant
'功能:获取某项的所有键值
'返回：=键值数组，数组内容为，名称,键值
    Dim lnghCurKey As Long, lngRet As Long, strName As String, strValue As String, lngIdx As Long
    Dim arrKeyValues As Variant

    arrKeyValues = Array()
    lngIdx = 0
    lngRet = RegOpenKey(lnghKey, strSubKey, 0, KEY_READ, lnghCurKey)
    If lngRet = ERROR_SUCCESS Then
        Do
            strName = String(256, Chr(0)): strValue = String(256, Chr(0))
            lngRet = RegEnumValue(lnghCurKey, lngIdx, strName, Len(strName), ByVal 0&, REG_SZ, strValue, Len(strValue))
            If lngRet = ERROR_SUCCESS Then
                ReDim Preserve arrKeyValues(UBound(arrKeyValues) + 1)
                arrKeyValues(UBound(arrKeyValues)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> ERROR_SUCCESS
        If lngRet <> ERROR_SUCCESS And lngRet <> ERROR_NO_MORE_ITEMS Then
            gobjTrace.WriteInfo "GetAllKeyValue", "枚举键值失败", GetLastDllErr(lngRet), "Key", GetRegPath(lnghKey) & "\" & strSubKey
        End If
        Call CloseKey(lnghCurKey)
    Else
        If lngRet <> ERROR_FILE_NOT_FOUND Then
            gobjTrace.WriteInfo "GetAllKeyValue", "打开注册表项失败", GetLastDllErr(lngRet), "Key", GetRegPath(lnghKey) & "\" & strSubKey
        End If
    End If
    GetAllKeyValue = arrKeyValues
End Function

Private Function DeleteKeyValue(ByVal lnghKey As Long, ByVal strSubKey As String) As Boolean
'功能:删除子健以及下面的键值
'strSubKey:子健，对于A\B\C,只会删除最后一个子健以及键值，即C键与C键下的键值,若C下面有子健则递归删除
'返回：是否删除成功
    Dim lngRet          As Long, lnghCurKey As Long
    Dim arrSubKey       As Variant, arrSubKeyValue      As Variant
    Dim i               As Long

    lngRet = RegOpenKey(lnghKey, strSubKey, 0, KEY_READ, lnghCurKey)
    If lngRet = ERROR_SUCCESS Then
        '获取子键，递归调用
        arrSubKey = GetAllSubKey(lnghKey, strSubKey)
        For i = LBound(arrSubKey) To UBound(arrSubKey)
            Call DeleteKeyValue(lnghCurKey, arrSubKey(i))
        Next
        Call CloseKey(lnghCurKey)
        lngRet = RegDeleteKey(lnghKey, strSubKey)
        If lngRet <> ERROR_SUCCESS Then
            gobjTrace.WriteInfo "DeleteKeyValue", "删除注册表项失败", GetLastDllErr(lngRet), "Key", GetRegPath(lnghKey) & "\" & strSubKey
            Exit Function
        End If
    Else
        If lngRet <> ERROR_FILE_NOT_FOUND Then
            gobjTrace.WriteInfo "DeleteKeyValue", "打开注册表项失败", GetLastDllErr(lngRet), "Key", GetRegPath(lnghKey) & "\" & strSubKey
        End If
    End If
    DeleteKeyValue = True
End Function

Private Function GetRegKey(Optional ByRef strErr As String) As Boolean
'功能：获取CLSID,IID,TID目录
    Dim lngReturn As Long

    If gblnIs64Bits Then
        mstrSysBitOpt = "\Wow6432node"
    Else
        mstrSysBitOpt = ""
    End If

    If mlngHMRoot = 0 Then
        lngReturn = RegOpenKey(HKEY_LOCaL_MaCHINE, "SOFTWARE" & mstrSysBitOpt & "\Classes", 0, KEY_READ Or KEY_WRITE, mlngHMRoot)
        If lngReturn <> ERROR_SUCCESS Then
            strErr = GetLastDllErr
            gobjTrace.WriteInfo "GetRegKey", "打开注册表项失败", strErr, "Key", "HKEY_LOCaL_MACHINE\SOFTWARE" & mstrSysBitOpt & "\Classes"
            Exit Function
        End If
    End If
    If mlngHMType = 0 Then
        lngReturn = RegOpenKey(mlngHMRoot, "TypeLib", 0, KEY_READ Or KEY_WRITE, mlngHMType)
        If lngReturn <> ERROR_SUCCESS Then
            strErr = GetLastDllErr
            gobjTrace.WriteInfo "GetRegKey", "打开注册表项失败", strErr, "Key", "HKEY_LOCaL_MACHINE\SOFTWARE" & mstrSysBitOpt & "\Classes\TypeLib"
            Exit Function
        End If
    End If
    If mlngHMClass = 0 Then
        lngReturn = RegOpenKey(mlngHMRoot, "CLSID", 0, KEY_READ Or KEY_WRITE, mlngHMClass)
        If lngReturn <> ERROR_SUCCESS Then
            strErr = GetLastDllErr
            gobjTrace.WriteInfo "GetRegKey", "打开注册表项失败", strErr, "Key", "HKEY_LOCaL_MACHINE\SOFTWARE" & mstrSysBitOpt & "\Classes\CLSID"
            Exit Function
        End If
    End If
    If mlngHMInterface = 0 Then
        lngReturn = RegOpenKey(mlngHMRoot, "Interface", 0, KEY_READ Or KEY_WRITE, mlngHMInterface)
        If lngReturn <> ERROR_SUCCESS Then
            strErr = GetLastDllErr
            gobjTrace.WriteInfo "GetRegKey", "打开注册表项失败", strErr, "Key", "HKEY_LOCaL_MACHINE\SOFTWARE" & mstrSysBitOpt & "\Classes\Interface"
            Exit Function
        End If
    End If
    mblnReadKey = True
    GetRegKey = True
End Function

Private Function GetNETRegTools(ByVal strNetVer As String) As String
'功能：获取RegAsm文件路径
'参数：strNetVer=.NET版本号
'返回：对应的RegAsm文件
'Private mstrNETRegToolsV4 As String
'Private mstrNETRegToolsV2 As String
    Dim lnghCurKey As Long, lngRet As Long, strValue As String
    Dim strSubKeyValues As Variant, strRegAsmRoot As String

    If strNetVer <> "" And mstrNETRegToolsV4 = "" And mstrNETRegToolsV2 = "" Then
        lngRet = RegOpenKey(HKEY_LOCaL_MaCHINE, "SOFTWARE\Microsoft\.NETFramework", 0, KEY_READ, lnghCurKey)
        If lngRet = ERROR_SUCCESS Then
            strValue = String(256, Chr(0))
            lngRet = RegQueryValue(lnghCurKey, "InstallRoot", 0, REG_SZ, strValue, Len(strValue))
            If lngRet = ERROR_SUCCESS Then
                strRegAsmRoot = Left(strValue, InStr(strValue, Chr(0)) - 1)
                strSubKeyValues = GetAllKeyValue(HKEY_LOCaL_MaCHINE, "SOFTWARE\Microsoft\.NETFramework\Policy\v4.0")
                If TypeName(strSubKeyValues) = "Empty" Then
                    mstrNETRegToolsV4 = " " '标记已经读取数据
                Else
                    mstrNETRegToolsV4 = strRegAsmRoot + "v4.0." + strSubKeyValues(0)
                End If
                strSubKeyValues = GetAllKeyValue(HKEY_LOCaL_MaCHINE, "SOFTWARE\Microsoft\.NETFramework\Policy\v2.0")
                If TypeName(strSubKeyValues) = "Empty" Then
                    mstrNETRegToolsV2 = " " '标记已经读取数据
                Else
                    mstrNETRegToolsV2 = strRegAsmRoot + "v2.0." + strSubKeyValues(0)
                End If
            End If
        End If
        Call CloseKey(lnghCurKey)
    End If

    If InStr(strNetVer, "v4") > 0 And Trim(mstrNETRegToolsV4) <> "" Then
        GetNETRegTools = mstrNETRegToolsV4
    ElseIf InStr(strNetVer, "v2") > 0 And Trim(mstrNETRegToolsV2) <> "" Then
        GetNETRegTools = mstrNETRegToolsV2
    Else
        GetNETRegTools = ""
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

Private Function CloseKey(ByVal lnghKey As Long) As Boolean
'功能：是否是64位系统
'返回：
    Dim lngRet As Long
    lngRet = RegCloseKey(lnghKey)
    If lngRet <> ERROR_SUCCESS Then
        gobjTrace.WriteInfo "CloseKey", "关闭注册表句柄出错", GetLastDllErr(lngRet), "Key", GetRegPath(lnghKey)
    Else
        CloseKey = True
    End If
End Function

Private Function GetRegPath(ByVal lngKey As Long) As String
    If lngKey = mlngHMRoot Then
        GetRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE" & mstrSysBitOpt & "\Classes"
    ElseIf lngKey = mlngHMType Then
        GetRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE" & mstrSysBitOpt & "\Classes\TypeLib"
    ElseIf lngKey = mlngHMClass Then
        GetRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE" & mstrSysBitOpt & "\Classes\CLSID"
    ElseIf lngKey = mlngHMInterface Then
        GetRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE" & mstrSysBitOpt & "\Classes\Interface"
    End If
    If GetRegPath <> "" Then Exit Function
    GetRegPath = "Unkown:" & lngKey
End Function

Private Sub Class_Initialize()
    mblnReadKey = False
    mlngHMType = 0
    mlngHMClass = 0
    mlngHMInterface = 0
    mlngHMRoot = 0
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Call CloseKey(mlngHMType)
    Call CloseKey(mlngHMClass)
    Call CloseKey(mlngHMInterface)
    Call CloseKey(mlngHMRoot)
End Sub
