Attribute VB_Name = "mdlAPI"
'@模块 mdlAPI-2019/6/26
'@编写 lshuo
'@功能
'   API相关常用方法
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
'API错误信息获取
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'状态栏图标
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'管道获取CMD输出
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'用来创建一个新的进程和它的主线程，这个新进程运行指定的可执行文件。如果函数执行成功，返回非零值
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'创建一个匿名管道，并从中得到读写管道的句柄。如果函数执行成功，返回非零值
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
'从文件指针指向的位置开始将数据读出到一个文件中， 且支持同步和异步操作
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
'当等待仍在挂起状态时，句柄被关闭，那么函数行为是未定义的。该句柄必须具有 SYNCHRONIZE 访问权限。
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Type STARTUPINFO
    cb                              As Long
    lpReserved                      As String
    lpDesktop                       As String
    lpTitle                         As String
    dwX                             As Long
    dwY                             As Long
    dwXSize                         As Long
    dwYSize                         As Long
    dwXCountChars                   As Long
    dwYCountChars                   As Long
    dwFillAttribute                 As Long
    dwFlags                         As Long
    wShowWindow                     As Integer
    cbReserved2                     As Integer
    lpReserved2                     As Long
    hStdInput                       As Long
    hStdOutput                      As Long
    hStdError                       As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess                        As Long
    hThread                         As Long
    dwProcessId                     As Long
    dwThreadId                      As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength                         As Long
    lpSecurityDescriptor            As Long
    bInheritHandle                  As Long
End Type
Private Const NORMAL_PRIORITY_CLASS  As Long = &H20&
Private Const STARTF_USESTDHANDLES   As Long = &H100&
Private Const STARTF_USESHOWWINDOW   As Long = &H1&
Private Const INFINITE               As Long = &HFFFF&
Private Const SW_HIDE = 0 '隐藏窗口，激活另一个窗口

Public Declare Function GetTickCount Lib "kernel32" () As Long
'功能：检索自系统启动以来已运行的毫秒数，最多可达49.7天。
'返回：返回值是自系统启动以来运行的毫秒数。
'注意事项：GetTickCount函数的解析仅限于系统计时器的精度，通常在10毫秒到16毫秒之间。
'        GetTickCount函数的解析不会受到getsystemtime调适函数的调整的影响?
'        经过的时间存储为DWORD的值?
'        因此，如果系统连续运行49.7天，那么时间将会是零。
'        为了避免这个问题，请使用GetTickCount64函数。
'        否则，在比较时检查溢出条件。
'        如果你需要一个更高的分辨率计时器，可以使用多媒体定时器或高分辨率计时器。
'        为了获得计算机启动后的时间，在注册表关键hkeyperformance cedata的性能数据中检索系统的时间计数器。
'        返回的值是一个8字节的值?
'        要了解更多信息，请参见性能计数器。
'        Note：获得时间系统在工作状态自启动 , 使用QueryUnbiasedInterruptTime函数?
'        调试注意QueryUnbiasedInterruptTime函数产生不同的结果(“checked”)构建Windows,因为中断时间计数是由大约49天。
'        这有助于识别在系统运行很长时间之前可能不会发生的错误?
'        通过Microsoft Developer Network(MSDN)Web站点可以对MSDN的用户进行检查。

'进程获取
Private Type MODULEENTRY32
    dwSize                                      As Long
    th32ModuleID                                As Long
    th32ProcessID                               As Long
    GlblcntUsage                                As Long
    ProccntUsage                                As Long
    modBaseAddr                                 As Byte
    modBaseSize                                 As Long
    hModule                                     As Long
    szModule                                    As String * 256
    szExePath                                   As String * 1024
End Type

Private Type PROCESSENTRY32
      lSize                                     As Long
      lUsage                                    As Long
      lProcessId                                As Long
      lDefaultHeapId                            As Long
      lModuleId                                 As Long
      lThreads                                  As Long
      lParentProcessId                          As Long
      lPriClassBase                             As Long
      lFlags                                    As Long
      sExeFile                                  As String * 1024
End Type

Private Const TH32CS_SNAPPROCESS                As Long = &H2
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long


Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'获取IP的常量与结构
Private Const MAX_ADAPTER_NAME_LENGTH           As Long = 256
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH    As Long = 128
Private Const MAX_ADAPTER_ADDRESS_LENGTH        As Long = 8
Private Const ERROR_SUCCESS                     As Long = 0
Private Type IP_ADDRESS_STRING
    IpAddr(0 To 15)                             As Byte
End Type
Private Type IP_MASK_STRING
    IpMask(0 To 15)                             As Byte
End Type
Private Type IP_ADDR_STRING
    dwNext                                      As Long
    IpAddress                                   As IP_ADDRESS_STRING
    IpMask                                      As IP_MASK_STRING
    dwContext                                   As Long
End Type
Private Type IP_ADAPTER_INFO
  dwNext                                        As Long
  ComboIndex                                    As Long  '保留
  sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3))        As Byte
  sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3)) As Byte
  dwAddressLength                               As Long
  sIPAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1))       As Byte
  dwIndex                                       As Long
  uType                                         As Long
  uDhcpEnabled                                  As Long
  CurrentIpAddress                              As Long
  IpAddressList                                 As IP_ADDR_STRING
  GatewayList                                   As IP_ADDR_STRING
  DhcpServer                                    As IP_ADDR_STRING
  bHaveWins                                     As Long
  PrimaryWinsServer                             As IP_ADDR_STRING
  SecondaryWinsServer                           As IP_ADDR_STRING
  LeaseObtained                                 As Long
  LeaseExpires                                  As Long
End Type
Private Const MAX_IP = 5        'To make a buffer... i dont think you have more than 5 ip on your pc..
Private Type IPINFO
    dwAddr                                      As Long              ' IP address
    dwIndex                                     As Long             ' interface index
    dwMask                                      As Long              ' subnet mask
    dwBCastAddr                                 As Long         ' broadcast address
    dwReasmSize                                 As Long        ' assembly size
    unused1                                     As Integer          ' not currently used
    unused2                                     As Integer          '; not currently used
End Type
Private Type MIB_IPADDRTABLE
    dEntrys                                     As Long             'number of entries in the table
    mIPInfo(MAX_IP)                             As IPINFO   'array of IP address entries
End Type
Private Type IP_Array
    mBuffer                                     As MIB_IPADDRTABLE
    BufferLen                                   As Long
End Type
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (pTcpTable As Any, pdwSize As Long) As Long
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long 'MD5计算

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long

Public Const SYNCHRONIZE                        As Long = &H100000
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Const PROCESS_TERMINATE                  As Long = &H1
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long


'版本信息获取
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (ByVal pBlock As Long, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Public Const FVN_Comments           As String = "Comments"          '注释
Public Const FVN_InternalName       As String = "InternalName"      '内部名称
Public Const FVN_ProductName        As String = "ProductName"       '产品名
Public Const FVN_CompanyName        As String = "CompanyName"       '公司名
Public Const FVN_ProductVersion     As String = "ProductVersion"    '产品版本
Public Const FVN_FileDescription    As String = "FileDescription"   '文件描述
Public Const FVN_OriginalFilename   As String = "OriginalFilename"  '原始文件名
Public Const FVN_FileVersion        As String = "FileVersion"       '文件版本
Public Const FVN_SpecialBuild       As String = "SpecialBuild"      '特殊编译号
Public Const FVN_PrivateBuild       As String = "PrivateBuild"      '私有编译号
Public Const FVN_LegalCopyright     As String = "LegalCopyright"    '合法版权
Public Const FVN_LegalTrademarks    As String = "LegalTrademarks"   '合法商标

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------

'@方法    GetLastDllErr
'   获取API错误描述
'@返回值  String
'
'@参数:
'lngErr Long In
'   错误代码，传递Err.LastDllError
'@备注
'
Public Function GetLastDllErr(ByVal lngErr As Long) As String
    Dim strReturn As String
    
    If lngErr = ERROR_EXTENDED_ERROR Then
        GetLastDllErr = GetWNetErr(lngErr)
    Else
        strReturn = String$(256, 32)
        FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lngErr, 0&, strReturn, Len(strReturn), ByVal 0
        strReturn = TruncZero(Trim(strReturn))
        GetLastDllErr = lngErr & "-" & Replace(Replace(strReturn, Chr(10), ""), Chr(13), "")
    End If
End Function


'@方法    AddIcon
'   在任务栏上增加一个图标
'@返回值
'
'@参数:
'lngHwnd Long In
'   处理事件的对象句柄
'stdIcon StdPicture In
'   图标
'strTip String In (Optional)
'   任务栏图标的右键提示
'@备注
'
Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)
    Shell_NotifyIcon NIM_ADD, t
End Sub

'@方法    RemoveIcon
'   从任务栏上删除图标
'@返回值
'
'@参数:
'lngHwnd Long In
'   处理事件的对象句柄
'@备注
'
Public Sub RemoveIcon(ByVal lngHwnd As Long)
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
End Sub

'@方法    RunCommand
'   执行命令，并通过标准管道读取命令执行结果
'@返回值  String
'
'@参数:
'commandline String In
'   cmd 命令行
'@备注
'
Public Function RunCommand(commandline As String) As String
    Dim si As STARTUPINFO                                                       'used to send info the CreateProcess
    Dim pi As PROCESS_INFORMATION                                               'used to receive info about the created process
    Dim retval As Long                                                          'return value
    Dim hRead As Long                                                           'the handle to the read end of the pipe
    Dim hWrite As Long                                                          'the handle to the write end of the pipe
    Dim sBuffer(0 To 63) As Byte                                                'the buffer to store data as we read it from the pipe
    Dim lgSize As Long                                                          'returned number of bytes read by readfile
    Dim sa As SECURITY_ATTRIBUTES
    Dim strResult As String                                                     'returned results of the command line
    
    'set up security attributes structure
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1&                                                    'inherit, needed for this to work
        .lpSecurityDescriptor = 0&
    End With
    'create our anonymous pipe an check for success
    ' note we use the default buffer size
    ' this could cause problems if the process tries to write more than this buffer size
    retval = CreatePipe(hRead, hWrite, sa, 0&)
    If retval = 0 Then
'        MsgBox "错误提示:创建管道失败!"
        RunCommand = ""
        Exit Function
    End If
    'set up startup info
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW                 'tell it to use (not ignore) the values below
        .wShowWindow = SW_HIDE
        .hStdOutput = hWrite                                                    'pass the write end of the pipe as the processes standard output
    End With
    'run the command line and check for success
    retval = CreateProcess(vbNullString, commandline & vbNullChar, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, si, pi)
    If retval Then
        'wait until the command line finishes
        ' trouble if the app doesn't end, or waits for user input, etc
        WaitForSingleObject pi.hProcess, INFINITE
        'read from the pipe until there's no more (bytes actually read is less than what we told it to)
        Do While ReadFile(hRead, sBuffer(0), 64, lgSize, ByVal 0&)
            'convert byte array to string and append to our result
            strResult = strResult & StrConv(sBuffer(), vbUnicode)
            'TODO = what's in the tail end of the byte array when lgSize is less than 64???
            Erase sBuffer()
            If lgSize <> 64 Then Exit Do
            DoEvents
        Loop
        'close the handles of the process
        CloseHandle pi.hProcess
        CloseHandle pi.hThread
    Else
'        MsgBox "错误提示:创建进程失败!" & vbCrLf
        RunCommand = ""
        Exit Function
    End If
    'close pipe handles
    CloseHandle hRead
    CloseHandle hWrite
    'return the command line output
    RunCommand = Replace(strResult, vbNullChar, "")
End Function

Public Function GetVersionInfo(ByVal strFileName As String, ByVal strEntryName As String) As String
    Dim i               As Long
    Dim lngVerSize      As Long
    Dim bytVerBlock()   As Byte
    Dim strSubBlock  As String
    Dim bytTranslate()  As Byte, lngAdrTranslate    As Long, lngTranslateSize       As Long
    Dim bytBuffer()     As Byte, lngBuffer          As Long, lngAdrBuffer           As Long
    
    On Error GoTo ErrH
    If Not gobjFSO.FileExists(strFileName) Then Exit Function
    lngVerSize = GetFileVersionInfoSize(strFileName, 0&)
    If lngVerSize <= 0 Then Exit Function
    
    ReDim bytVerBlock(lngVerSize - 1)
    Call GetFileVersionInfo(strFileName, 0&, lngVerSize, bytVerBlock(0))
    
    VerQueryValue VarPtr(bytVerBlock(0)), "\\VarFileInfo\\Translation", lngAdrTranslate, lngTranslateSize
    ReDim bytTranslate(lngTranslateSize - 1)
    RtlMoveMemory bytTranslate(0), ByVal lngAdrTranslate, lngTranslateSize
    For i = 1 To lngTranslateSize / (UBound(bytTranslate) + 1)
        strSubBlock = "\\StringFileInfo\\"
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 0, 1, True)
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 2, 3, True)
        strSubBlock = strSubBlock & "\\" & strEntryName
        
        VerQueryValue VarPtr(bytVerBlock(0)), strSubBlock, lngAdrBuffer, lngBuffer
        If lngAdrBuffer <> 0 And lngBuffer <> 0 Then
            ReDim bytBuffer(lngBuffer - 1)
            RtlMoveMemory bytBuffer(0), ByVal lngAdrBuffer, lngBuffer
            ReDim Preserve bytBuffer(InStrB(bytBuffer, ChrB(0)) - 2)
            GetVersionInfo = StrConv(bytBuffer, vbUnicode)
        End If
    Next
    Exit Function
ErrH:
    Err.Clear
End Function


'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------

'@方法    GetWNetErr
'   获取网路扩展错误
'@返回值  String
'
'@参数:
'lngErr Long In
'   错误代码
'@备注
'
Private Function GetWNetErr(ByVal lngErr As Long) As String
    Dim strErr As String * 256
    Dim strName As String * 256
    Dim lngRet As Long
    lngRet = WNetGetLastError(lngErr, strErr, Len(strErr), strName, Len(strName))
    GetWNetErr = lngErr & "-" & Replace(Replace("[" & TruncZero(strName) & "]" & TruncZero(strErr), Chr(10), ""), Chr(13), "")
End Function

Private Function ConvertAddressToString(longAddr As Long, Optional ByRef strErr As String) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    strErr = ""
    On Error GoTo ErrH
    RtlMoveMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrH:
    strErr = Err.Description
    Err.Clear
End Function

Private Function Byte2Hex(bytArray() As Byte, Optional ByVal lngStart As Long = 0, Optional ByVal lngEnd As Long = -1, Optional fReversed As Boolean = False) As String
    Dim i     As Long
    lngStart = IIf(lngStart < 0, 0, lngStart)
    lngEnd = IIf(lngEnd < 0, UBound(bytArray), lngEnd)
    
    If fReversed Then
        For i = lngEnd To lngStart Step -1
            Byte2Hex = Byte2Hex & Right$("00" & Hex(bytArray(i)), 2)
        Next
    Else
        For i = lngStart To lngEnd
            Byte2Hex = Byte2Hex & Right$("00" & Hex(bytArray(i)), 2)
        Next
    End If
End Function
'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------


