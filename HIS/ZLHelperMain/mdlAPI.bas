Attribute VB_Name = "mdlAPI"
'@ģ�� mdlAPI-2019/6/26
'@��д lshuo
'@����
'   API��س��÷���
'@����
'
'@��ע
'
Option Explicit
'---------------------------------------------------------------------------
'                0��API�ͳ�������
'---------------------------------------------------------------------------
'API������Ϣ��ȡ
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'״̬��ͼ��
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

'�ܵ���ȡCMD���
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'��������һ���µĽ��̺��������̣߳�����½�������ָ���Ŀ�ִ���ļ����������ִ�гɹ������ط���ֵ
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'����һ�������ܵ��������еõ���д�ܵ��ľ�����������ִ�гɹ������ط���ֵ
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
'���ļ�ָ��ָ���λ�ÿ�ʼ�����ݶ�����һ���ļ��У� ��֧��ͬ�����첽����
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
'���ȴ����ڹ���״̬ʱ��������رգ���ô������Ϊ��δ����ġ��þ��������� SYNCHRONIZE ����Ȩ�ޡ�
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
Private Const SW_HIDE = 0 '���ش��ڣ�������һ������

Public Declare Function GetTickCount Lib "kernel32" () As Long
'���ܣ�������ϵͳ�������������еĺ����������ɴ�49.7�졣
'���أ�����ֵ����ϵͳ�����������еĺ�������
'ע�����GetTickCount�����Ľ���������ϵͳ��ʱ���ľ��ȣ�ͨ����10���뵽16����֮�䡣
'        GetTickCount�����Ľ��������ܵ�getsystemtime���ʺ����ĵ�����Ӱ��?
'        ������ʱ��洢ΪDWORD��ֵ?
'        ��ˣ����ϵͳ��������49.7�죬��ôʱ�佫�����㡣
'        Ϊ�˱���������⣬��ʹ��GetTickCount64������
'        �����ڱȽ�ʱ������������
'        �������Ҫһ�����ߵķֱ��ʼ�ʱ��������ʹ�ö�ý�嶨ʱ����߷ֱ��ʼ�ʱ����
'        Ϊ�˻�ü�����������ʱ�䣬��ע���ؼ�hkeyperformance cedata�����������м���ϵͳ��ʱ���������
'        ���ص�ֵ��һ��8�ֽڵ�ֵ?
'        Ҫ�˽������Ϣ����μ����ܼ�������
'        Note�����ʱ��ϵͳ�ڹ���״̬������ , ʹ��QueryUnbiasedInterruptTime����?
'        ����ע��QueryUnbiasedInterruptTime����������ͬ�Ľ��(��checked��)����Windows,��Ϊ�ж�ʱ��������ɴ�Լ49�졣
'        ��������ʶ����ϵͳ���кܳ�ʱ��֮ǰ���ܲ��ᷢ���Ĵ���?
'        ͨ��Microsoft Developer Network(MSDN)Webվ����Զ�MSDN���û����м�顣

'���̻�ȡ
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

'��ȡIP�ĳ�����ṹ
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
  ComboIndex                                    As Long  '����
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
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long 'MD5����

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long

Public Const SYNCHRONIZE                        As Long = &H100000
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Const PROCESS_TERMINATE                  As Long = &H1
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long


'�汾��Ϣ��ȡ
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (ByVal pBlock As Long, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Public Const FVN_Comments           As String = "Comments"          'ע��
Public Const FVN_InternalName       As String = "InternalName"      '�ڲ�����
Public Const FVN_ProductName        As String = "ProductName"       '��Ʒ��
Public Const FVN_CompanyName        As String = "CompanyName"       '��˾��
Public Const FVN_ProductVersion     As String = "ProductVersion"    '��Ʒ�汾
Public Const FVN_FileDescription    As String = "FileDescription"   '�ļ�����
Public Const FVN_OriginalFilename   As String = "OriginalFilename"  'ԭʼ�ļ���
Public Const FVN_FileVersion        As String = "FileVersion"       '�ļ��汾
Public Const FVN_SpecialBuild       As String = "SpecialBuild"      '��������
Public Const FVN_PrivateBuild       As String = "PrivateBuild"      '˽�б����
Public Const FVN_LegalCopyright     As String = "LegalCopyright"    '�Ϸ���Ȩ
Public Const FVN_LegalTrademarks    As String = "LegalTrademarks"   '�Ϸ��̱�

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'---------------------------------------------------------------------------
'                1���������
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                2�����Ա����붨��
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3����������
'---------------------------------------------------------------------------

'@����    GetLastDllErr
'   ��ȡAPI��������
'@����ֵ  String
'
'@����:
'lngErr Long In
'   ������룬����Err.LastDllError
'@��ע
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


'@����    AddIcon
'   ��������������һ��ͼ��
'@����ֵ
'
'@����:
'lngHwnd Long In
'   �����¼��Ķ�����
'stdIcon StdPicture In
'   ͼ��
'strTip String In (Optional)
'   ������ͼ����Ҽ���ʾ
'@��ע
'
Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)
    Shell_NotifyIcon NIM_ADD, t
End Sub

'@����    RemoveIcon
'   ����������ɾ��ͼ��
'@����ֵ
'
'@����:
'lngHwnd Long In
'   �����¼��Ķ�����
'@��ע
'
Public Sub RemoveIcon(ByVal lngHwnd As Long)
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼�����������
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
End Sub

'@����    RunCommand
'   ִ�������ͨ����׼�ܵ���ȡ����ִ�н��
'@����ֵ  String
'
'@����:
'commandline String In
'   cmd ������
'@��ע
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
'        MsgBox "������ʾ:�����ܵ�ʧ��!"
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
'        MsgBox "������ʾ:��������ʧ��!" & vbCrLf
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
'                4��˽�з���
'---------------------------------------------------------------------------

'@����    GetWNetErr
'   ��ȡ��·��չ����
'@����ֵ  String
'
'@����:
'lngErr Long In
'   �������
'@��ע
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
'                5�����󷽷����¼�
'---------------------------------------------------------------------------


