Attribute VB_Name = "mdlPubDefine"
Option Explicit
'==========================================================
'������ṹ
'==========================================================
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
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (pTcpTable As Any, pdwSize As Long) As Long
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long 'MD5����
'API������Ϣ��ȡ
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
'MD5����
Public Declare Function CreateFileA Lib "kernel32.dll" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Const GENERIC_READ                      As Long = &H80000000
Private Const FILE_SHARE_READ                   As Long = &H1
Private Const OPEN_EXISTING                     As Long = &H3
Private Const FILE_ATTRIBUTE_NORMAL             As Long = &H80
Private Const INVALID_HANDLE_VALUE              As Long = (-1)
Private Const PAGE_READONLY                     As Long = &H2
Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As Long, ByRef lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal R As String)
Private Declare Function CryptAcquireContextA Lib "advapi32.dll" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Const CRYPT_NEWKEYSET                   As Long = &H8
Private Const PROV_RSA_FULL                     As Long = 1
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As Long, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Const FILE_MAP_READ                     As Long = &H4
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Const HP_HASHVAL                        As Long = 2
Private Const HP_HASHSIZE                       As Long = 4
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Enum HashAlgorithm
    HA_MD2 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
    HA_MD4 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
    HA_MD5 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
    HA_SHA = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA
End Enum
Private Type LARGE_INTEGER
    lowpart     As Long
    highpart    As Long
End Type
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

'�ܵ�����
Private Type PROCESS_INFORMATION
    hProcess    As Long
    hThread     As Long
    dwProcessID As Long
    dwThreadID  As Long
End Type

Private Type STARTUPINFO
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

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'���ȴ����ڹ���״̬ʱ��������رգ���ô������Ϊ��δ����ġ��þ��������� SYNCHRONIZE ����Ȩ�ޡ�
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const NORMAL_PRIORITY_CLASS             As Long = &H20&
Private Const STARTF_USESTDHANDLES              As Long = &H100&
Private Const STARTF_USESHOWWINDOW              As Long = &H1
Private Const SW_HIDE                           As Integer = 0 '���ش��ڣ�������һ������
Public Const INFINITE                           As Long = &HFFFF&

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
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Private Const TH32CS_SNAPMODULE                 As Long = &H8
Public Const SYNCHRONIZE                       As Long = &H100000
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Const PROCESS_TERMINATE                 As Long = &H1
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'ϵͳ��Ȩ���
Private Const ANYSIZE_ARRAY = 1
Private Const TOKEN_ADJUST_PRIVILEGES = (&H20)
Private Const TOKEN_QUERY = (&H8)
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Type LUID
    lowpart As Long
    highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type


Public Const SE_DEBUG_NAME = "SeDebugPrivilege"
Public Const SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege"
Public Const SE_AUDIT_NAME = "SeAuditPrivilege"
Public Const SE_BACKUP_NAME = "SeBackupPrivilege"
Public Const SE_CHANGE_NOTIFY_NAME = "SeChangeNotifyPrivilege"
Public Const SE_CREATE_PAGEFILE_NAME = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT_NAME = "SeCreatePermanentPrivilege"
Public Const SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege"
Public Const SE_INC_BASE_PRIORITY_NAME = "SeIncreaseBasePriorityPrivilege"
Public Const SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"
Public Const SE_LOAD_DRIVER_NAME = "SeLoadDriverPrivilege"
Public Const SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege"
Public Const SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege"
Public Const SE_PROF_SINGLE_PROCESS_NAME = "SeProfileSingleProcessPrivilege"
Public Const SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege"
Public Const SE_RESTORE_NAME = "SeRestorePrivilege"
Public Const SE_SECURITY_NAME = "SeSecurityPrivilege"
Public Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT_NAME = "SeSystemEnvironmentPrivilege"
Public Const SE_SYSTEM_PROFILE_NAME = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege"
Public Const SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege"
Public Const SE_TCB_NAME = "SeTcbPrivilege"
Public Const SE_UNSOLICITED_INPUT_NAME = "SeUnsolicitedInputPrivilege"
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long
'��ͣ(Wait)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'ϵͳ�ж�
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long
'����Ա�ж�
Private Const ANYSIZE_ARRAY_EX = 20 'Fixed at this size for comfort. Could be bigger or made dynamic.
' Security APIs
Private Const TokenUser = 1
Private Const TokenGroups = 2
Private Const TokenPrivileges = 3
Private Const TokenOwner = 4
Private Const TokenPrimaryGroup = 5
Private Const TokenDefaultDacl = 6
Private Const TokenSource = 7
Private Const TokenType = 8
Private Const TokenImpersonationLevel = 9
Private Const TokenStatistics = 10
' Token Specific Access Rights
Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = &H2
Private Const TOKEN_IMPERSONATE = &H4
Private Const TOKEN_QUERY_SOURCE = &H10
Private Const TOKEN_ADJUST_GROUPS = &H40
Private Const TOKEN_ADJUST_DEFAULT = &H80
' NT well-known SIDs
Private Const SECURITY_DIALUP_RID = &H1
Private Const SECURITY_NETWORK_RID = &H2
Private Const SECURITY_BATCH_RID = &H3
Private Const SECURITY_INTERACTIVE_RID = &H4
Private Const SECURITY_SERVICE_RID = &H6
Private Const SECURITY_ANONYMOUS_LOGON_RID = &H7
Private Const SECURITY_LOGON_IDS_RID = &H5
Private Const SECURITY_LOCAL_SYSTEM_RID = &H12
Private Const SECURITY_NT_NON_UNIQUE = &H15
Private Const SECURITY_BUILTIN_DOMAIN_RID = &H20
' Well-known domain relative sub-authority values (RIDs)
Private Const DOMAIN_ALIAS_RID_ADMINS = &H220
Private Const DOMAIN_ALIAS_RID_USERS = &H221
Private Const DOMAIN_ALIAS_RID_GUESTS = &H222
Private Const DOMAIN_ALIAS_RID_POWER_USERS = &H223
Private Const DOMAIN_ALIAS_RID_ACCOUNT_OPS = &H224
Private Const DOMAIN_ALIAS_RID_SYSTEM_OPS = &H225
Private Const DOMAIN_ALIAS_RID_PRINT_OPS = &H226
Private Const DOMAIN_ALIAS_RID_BACKUP_OPS = &H227
Private Const DOMAIN_ALIAS_RID_REPLICATOR = &H228

Private Const SECURITY_NT_AUTHORITY = &H5
Type SID_AND_ATTRIBUTES
    Sid         As Long
    Attributes  As Long
End Type

Type TOKEN_GROUPS
    GroupCount  As Long
    Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type

Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type

Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function OpenThreadToken Lib "advapi32" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "advapi32" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function IsValidSid Lib "advapi32" (ByVal pSid As Long) As Long
Private Declare Function EqualSid Lib "advapi32" (pSid1 As Any, pSid2 As Any) As Long
Private Declare Sub FreeSid Lib "advapi32" (pSid As Any)


'==========================================================
'��������
'==========================================================
Public Function IsDesinMode() As Boolean
'���ܣ� ȷ����ǰģʽΪ���ģʽ
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

'������ܳ���
Public Function Cipher(ByVal strText As String) As String
    Const MIN_ASC = 32    '��СASCII��
    Const MAX_ASC = 126 '���ASCII�� �ַ�
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    Dim strSeed As String
    
    If strText = "" Then Exit Function
    '��ȡ�������
    '������ӵ������Ϊ999
    Rnd (-1)
    Randomize (999)
    strSeed = "456"
    intSeedLen = Len(strSeed)
    strDeText = Chr(intSeedLen + MIN_ASC)
    For i = 1 To intSeedLen
        intChr = Asc(Mid(strSeed, i, 1)) 'ȡ��ĸת���ASCII��
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        End If
    Next
    Rnd (-1)
    Randomize (Val(strSeed))
    intlen = Len(strText)
    For i = 1 To intlen
        intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        ElseIf intChr < 0 Then
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    Cipher = strDeText
End Function

Public Function DeCipher(ByVal strText As String) As String
'������ܳ���
    Const MIN_ASC = 32    '��СASCII��
    Const MAX_ASC = 126 '���ASCII�� �ַ�
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '������ӳ���
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intlen = Len(strText)
    '���þɵ�����㷨
    If intSeedLen > 0 And intSeedLen < intlen - 3 And intSeedLen < 5 Then
        '��ȡ�������
        '������ӵ������Ϊ999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
            If intChr >= MIN_ASC And intChr <= MAX_ASC Then
                intChr = intChr - MIN_ASC
                lngOffset = Int((NUM_ASC + 1) * Rnd())
                intChr = ((intChr - lngOffset) Mod NUM_ASC)
                If intChr < 0 Then
                    intChr = intChr + NUM_ASC
                End If
                intChr = intChr + MIN_ASC
                strDeText = strDeText & Chr(intChr)
            End If
        Next
        If Not IsNumeric(strDeText) Then
            strDeText = "123"
            intStart = 1
        Else
            intStart = 2 + intSeedLen
        End If
    Else
        strDeText = "123"
        intStart = 1
    End If
        
    '���ݽ��ܵ�����
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intlen
        intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr - lngOffset) Mod NUM_ASC)
            If intChr < 0 Then
                intChr = intChr + NUM_ASC
            End If
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        Else
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    DeCipher = strDeText
End Function

Public Function SQLAdjust(ByVal varInput As Variant, Optional ByVal lngMaxLength As Long) As String
'���ܣ�������"'"���ŵ��ַ�������ΪOracle����ʶ����ַ�����,�����մ�ת��ΪNull
'lngMaxLength=������Ƴ��ȣ���0��������
'˵�����Զ�(����)�����߼�"'"�綨����

    Dim i As Long, strTmp As String, strOneChar As String
    Dim strReturn As String
    Dim lngLine As Long
    Dim lngLength As Long
    
    strReturn = varInput & ""
    If strReturn & "" = "" Then SQLAdjust = "Null": Exit Function
    If InStr(1, strReturn, "'") = 0 And InStr(1, strReturn, Chr(10)) = 0 And InStr(1, strReturn, Chr(13)) = 0 Then
        SQLAdjust = "'" & strReturn & "'"
        Exit Function
    End If
    For i = 1 To Len(strReturn)
        strOneChar = Mid(strReturn, i, 1)
        Select Case strOneChar
            Case "'"
                If i = 1 Then
                    strTmp = "CHR(39)||'"
                ElseIf i = Len(strReturn) Then
                    strTmp = strTmp & "'||CHR(39)"
                Else
                    strTmp = strTmp & "'||CHR(39)||'"
                End If
                lngLine = lngLine + 1 '��ʶ�зǻ����ַ�
            Case Chr(10), Chr(13)
                If i = 1 Then
                    strTmp = "CHR(13)||'"
                ElseIf lngLine = 0 Then '���Ŷ�����У�����һ��
                    If i = Len(strReturn) Then '���һ���ǻ���
                        strTmp = strTmp & "'"
                    End If
                ElseIf i = Len(strReturn) Then
                    strTmp = strTmp & "'||CHR(13)"
                Else
                    strTmp = strTmp & "'||CHR(13)||'"
                End If
                lngLine = 0 '��ʶ�Ѿ��л���
            Case Else
                If i = 1 Then
                    strTmp = "'" & Mid(strReturn, i, 1)
                ElseIf i = Len(strReturn) Then
                    strTmp = strTmp & Mid(strReturn, i, 1) & "'"
                Else
                    strTmp = strTmp & Mid(strReturn, i, 1)
                End If
                lngLine = lngLine + 1 '��ʶ�зǻ����ַ�
        End Select
    Next
    SQLAdjust = strTmp
End Function

Public Function TrimEx(ByVal strTrim As String, Optional ByVal strTrmChar As String = " ") As String
'���ܣ�ȥ��strTrim���ߵ�strTrmChar,��������Trim
'         ����strTrmChar���ߴ��ո�ʱ���൱Trim
    Dim i As Integer, intB As Integer, intE As Integer
    
    If strTrim = "" Or strTrmChar = "" Then TrimEx = strTrim: Exit Function
    If strTrmChar = " " Then TrimEx = Trim(strTrim): Exit Function
    
    intB = 1
    For i = 1 To Len(strTrim)
        If Mid(strTrim, i, 1) <> strTrmChar Then intB = i: Exit For
    Next
    intE = Len(strTrim)
    For i = Len(strTrim) To 1 Step -1
        If Mid(strTrim, i, 1) <> strTrmChar Then intE = i: Exit For
    Next
    TrimEx = Mid(strTrim, intB, intE - intB + 1)
End Function

Public Function ComputerName() As String
'���ܣ���ȡ��������
    Dim strComputer As String * 256
    
    Call GetComputerName(strComputer, 255)
    ComputerName = strComputer
    ComputerName = Trim(Replace(ComputerName, Chr(0), ""))
End Function

Public Function IP(Optional ByRef strErr As String) As String
'���ܣ�ͨ��API��ȡ��ʱIP
    Dim ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    Dim strTmpErr As String, strALLErr As String
    
    strErr = ""
    On Error GoTo ErrHand
    GetIpAddrTable ByVal 0&, ret, True
    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
    'retrieve the data
    GetIpAddrTable bBytes(0), ret, False
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    For Tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr, strTmpErr)
        If strTmpErr <> "" Then strALLErr = strALLErr & IIf(strALLErr = "", "", "|") & strTmpErr
    Next Tel
    'Sort Out The IP For WAN
    TempIP = TempList(0)
    For Tempi = 0 To Listing.dEntrys - 1
        L3 = Left(TempList(Tempi), 3)
        If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
            TempIP = TempList(Tempi)
        End If
    Next Tempi
    IP = TempIP 'Return The TempIP
    strErr = strALLErr
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    strErr = strALLErr & IIf(strALLErr = "", "", "|") & Err.Description
    Err.Clear
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

'SIZE��ÿ��Ӱ����ļ���С ֻ����2��N�η�  ��: 2^27=2��27�η�=128M
Public Function FileMD5(ByVal szFilePath As String, Optional ByVal haCur As Long = HA_MD5, Optional ByVal Block_Size As Long = 32768) As String
    Dim lnghFile As Long, lnghMapFile As Long, lnglpBaseMap As Long
    Dim lnghCtx As Long, lngRet As Long, lnghHash As Long, lngLen As Long
    Dim i As Long, j As Long, lngPoint As Long
    Dim lintFI As LARGE_INTEGER, lintCurrent As LARGE_INTEGER, dblCurrentPoint As Double
    Dim lngTmp As Long, lngBlocks As Long, lngLastBlock As Long, Block() As Byte
    Dim lngSize As Long
    '�����ļ�ָ��
    DoEvents
    lngSize = 2 ^ 27
    lnghFile = CreateFileA(szFilePath, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If lnghFile <> INVALID_HANDLE_VALUE Then
        lintFI.lowpart = GetFileSize(lnghFile, lintFI.highpart) '�ɹ��� ��ȡ�ļ���С
        If lintFI.highpart > 0 Then lngBlocks = ((2 ^ 32 / lngSize) * lintFI.highpart) ' ��λ   Ϊ1���� 2^32���ֽ�  Ҳ����4�ֽ��޷��ų�������ֵ
        If lintFI.lowpart < 0 Then        '��λ
            lngBlocks = lngBlocks + (2 ^ 31 / lngSize) '��λΪ���� ��Ȼ����2^31�η�  ��Ϊ������2^31  VB����������ʾ
            lngTmp = LongToUnsigned(lintFI.lowpart) - 2 ^ 31 'תΪ�޷������ͼ���2^31�� VB����������ʾ��������
            lngLastBlock = lngTmp \ lngSize
            lngBlocks = lngBlocks + lngLastBlock
            lngLastBlock = lngTmp - lngLastBlock * lngSize
        Else
            lngTmp = lintFI.lowpart \ lngSize
            lngBlocks = lngBlocks + lngTmp
            lngLastBlock = lintFI.lowpart - lngTmp * lngSize
        End If
        
        lnghMapFile = CreateFileMapping(lnghFile, ByVal 0&, PAGE_READONLY, lintFI.highpart, lintFI.lowpart, 0) '�����ļ�ӳ�����
        lngRet = CryptAcquireContextA(lnghCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)
        If Err.LastDllError = &H80090016 Then lngRet = CryptAcquireContextA(lnghCtx, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
        lngRet = CryptCreateHash(lnghCtx, haCur, 0, 0, lnghHash)
        ReDim Block(Block_Size) As Byte
        
        For i = 1 To lngBlocks '�ɹ������ָ����С ��ʼӰ���ļ����ڴ�ռ�
            lnglpBaseMap = MapViewOfFile(lnghMapFile, FILE_MAP_READ, lintCurrent.highpart, lintCurrent.lowpart, lngSize)
            If lnglpBaseMap Then
                lngPoint = lnglpBaseMap
                For j = 1 To lngSize / Block_Size ' 2��N�η�  ��Ȼ����
                    
                    lngRet = CryptHashData(lnghHash, lngPoint, Block_Size, 0)
                    lngPoint = lngPoint + Block_Size
                Next
                UnmapViewOfFile (lnglpBaseMap)
            End If
            dblCurrentPoint = dblCurrentPoint + lngSize
            lintCurrent = Currency2LargeInteger(dblCurrentPoint / 10000@) '�����ļ��ߵ�λ
        Next
            
        If lngLastBlock > 0 Then 'ӳ������
            lnglpBaseMap = MapViewOfFile(lnghMapFile, FILE_MAP_READ, lintCurrent.highpart, lintCurrent.lowpart, lngLastBlock)
            If lnglpBaseMap Then
                lngPoint = lnglpBaseMap
                lngTmp = lngLastBlock \ Block_Size '��һ������ ������FOR ѭ�����ٴμ���
                
                For j = 1 To lngTmp
                    lngRet = CryptHashData(lnghHash, lngPoint, Block_Size, 0)
                    lngPoint = lngPoint + Block_Size
                Next
                lngTmp = lngLastBlock - lngTmp * Block_Size
                lngRet = CryptHashData(lnghHash, lngPoint, lngTmp, 0)
                UnmapViewOfFile (lnglpBaseMap)
            End If
        End If
        Call CloseHandle(lnghMapFile)
        If lngRet Then
            lngRet = CryptGetHashParam(lnghHash, HP_HASHSIZE, lngLen, 4, 0)
            If lngRet Then
                ReDim hash(lngLen) As Byte
                lngRet = CryptGetHashParam(lnghHash, HP_HASHVAL, hash(0), lngLen, 0)
                If lngRet Then
                    For j = 0 To UBound(hash) - 1
                        FileMD5 = FileMD5 & Right$("0" & Hex$(hash(j)), 2)
                    Next
                End If
                CryptDestroyHash lnghHash
            End If
        End If
        CryptReleaseContext lnghCtx, 0
        CloseHandle (lnghFile)
        
        If FileMD5 = "" Then
            On Error Resume Next
            FileMD5 = MD5File(szFilePath)
        End If
    End If
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
    CopyMemory bytTranslate(0), ByVal lngAdrTranslate, lngTranslateSize
    For i = 1 To lngTranslateSize / (UBound(bytTranslate) + 1)
        strSubBlock = "\\StringFileInfo\\"
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 0, 1, True)
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 2, 3, True)
        strSubBlock = strSubBlock & "\\" & strEntryName
        
        VerQueryValue VarPtr(bytVerBlock(0)), strSubBlock, lngAdrBuffer, lngBuffer
        If lngAdrBuffer <> 0 And lngBuffer <> 0 Then
            ReDim bytBuffer(lngBuffer - 1)
            CopyMemory bytBuffer(0), ByVal lngAdrBuffer, lngBuffer
            ReDim Preserve bytBuffer(InStrB(bytBuffer, ChrB(0)) - 2)
            GetVersionInfo = StrConv(bytBuffer, vbUnicode)
        End If
    Next
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function RunCommand(ByVal strCommand As String, Optional ByRef strErr As String, Optional ByVal blnCiper As Boolean, Optional ByVal lngWait As Long = INFINITE) As String
'���ܣ�ִ�������У�����ȡ���������
    Dim piProc          As PROCESS_INFORMATION '������Ϣ
    Dim stStart         As STARTUPINFO '������Ϣ
    Dim saSecAttr       As SECURITY_ATTRIBUTES '��ȫ����
    Dim lnghReadPipe    As Long '��ȡ�ܵ����
    Dim lnghWritePipe   As Long 'д��ܵ����
    Dim lngBytesRead    As Long '�������ݵ��ֽ���
    Dim strBuffer       As String * 256 '��ȡ�ܵ����ַ���buffer
    Dim lngRet          As Long 'API��������ֵ
    Dim lngRetPro       As Long
    Dim strlpOutputs    As String '���������ս��
    
    DoEvents
    On Error Resume Next
    If blnCiper Then
        gobjTrace.WriteInfo "RunCommand", "����", Cipher(strCommand)
    Else
        gobjTrace.WriteInfo "RunCommand", "����", strCommand
    End If
    '���ð�ȫ����
    With saSecAttr
        .nLength = LenB(saSecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With
    
    '�����ܵ�
    lngRet = CreatePipe(lnghReadPipe, lnghWritePipe, saSecAttr, 0)
    If lngRet = 0 Then
        strErr = "�޷������ܵ���" & GetLastDllErr()
        gobjTrace.WriteInfo "RunCommand", "����ܵ�����ʧ��", strErr
        Exit Function
    End If
    '���ý�������ǰ����Ϣ
    With stStart
        .Cb = LenB(stStart)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = lnghWritePipe '��������ܵ�
        .hStdError = lnghWritePipe '���ô���ܵ�
    End With
    '��������
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS������ipconfig.exeΪ��
    lngRetPro = CreateProcess(vbNullString, strCommand & vbNullChar, saSecAttr, saSecAttr, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, stStart, piProc)
    If lngRetPro = 0 Then
        strErr = "�޷��������̡�" & GetLastDllErr()
        gobjTrace.WriteInfo "RunCommand", "�����������ʧ��", strErr
        lngRet = CloseHandle(lnghWritePipe)
        lngRet = CloseHandle(lnghReadPipe)
        Exit Function
    Else
        '��Ϊ����д�����ݣ������ȹر�д��ܵ��������������رմ˹ܵ��������޷���ȡ����
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
        Loop While (lngRet <> 0) '��ret=0ʱ˵��ReadFileִ��ʧ�ܣ��Ѿ�û�����ݿɶ���
        '��ȡ������ɣ��رո����
        lngRet = CloseHandle(lngRetPro)
        lngRet = CloseHandle(piProc.hProcess)
        lngRet = CloseHandle(piProc.hThread)
        lngRet = CloseHandle(lnghReadPipe)
    End If
    RunCommand = Replace(strlpOutputs, vbNullChar, "")
End Function

Public Function FindExitsProcess(ByVal strProcessName As String, Optional ByVal lngCurProcID As Long) As Long
'���ܣ����ݳ������Ʋ��Ҳ�������
    Dim uProcess As PROCESSENTRY32
    Dim lngProcID As Long
    Dim lngSnapShot As Long, lngRet As Long
    Dim strFindName As String, lngPos As Long
    Dim lngPid As Long
    
    FindExitsProcess = 0
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot <> 0 Then
        uProcess.lSize = 1060
        If (Process32First(lngSnapShot, uProcess)) Then
            Do
                lngPos = InStr(1, uProcess.sExeFile, Chr(0))
                strFindName = UCase(Left(uProcess.sExeFile, lngPos - 1))
                If strFindName = strProcessName Then
                    lngPid = uProcess.lProcessId
                    If lngCurProcID <> lngPid Then
                        lngProcID = OpenProcess(1&, -1&, lngPid)
                        FindExitsProcess = lngProcID
                    End If
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
        lngRet = CloseHandle(lngSnapShot)
    End If
End Function

Public Function GetCommpentVersion(ByVal strFile As String) As String
    '����:��ȡָ���ؼ��İ汾���ɵ��ַ���
    
    Dim strVer As String, varVersion As Variant
    Dim strTmp As String
    
    Err = 0: On Error Resume Next
    '��ȡ�ļ��汾��
    strVer = gobjFSO.GetFileVersion(strFile)
    If Err <> 0 Then
        Err.Clear: Err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
    If Trim(strVer) <> "" Then
        varVersion = Split(strVer, ".")
        If UBound(varVersion) > 2 Then
            strVer = Val(varVersion(0)) * 10 ^ 8 + Val(varVersion(1)) * 10 ^ 4 + Val(varVersion(3))
        ElseIf UBound(varVersion) = 2 Then
            strVer = Val(varVersion(0)) * 10 ^ 8 + Val(varVersion(1)) * 10 ^ 4 + Val(varVersion(2))
        End If
        If strVer <> "" Then
            strTmp = GetVersionInfo(strFile, FVN_FileDescription)
            If IsNumeric(strTmp) Then
                strVer = strVer & Format(strTmp, "0000")
            End If
        End If
    End If
    GetCommpentVersion = strVer
End Function

Public Function DeCompression(ByVal strDesFile As String, ByVal strSourceFile As String, Optional ByVal intRate As Integer, Optional ByVal blnCompression As Boolean, Optional ByRef strErr As String) As Boolean
'���ܣ�����ѹ����ѹ(��ǰ��֧�ֵ��ļ���
'������
'       strDesFile=������ļ�·��������
'       strSourceFile=ԭʼ�ļ�
'       intRate=ѹ���ȼ���ѹ��ʹ�á�
'                   ѹ���ȼ� ѹ���㷨 �ֵ��С �����ֽ� ƥ���� ������ ����
'                   0           Copy    ��ѹ��
'                   1           LZMA    64KB     32       HC4   BCJ   ���ѹ��
'                   3           LZMA    1MB      32       HC4   BCJ   ����ѹ��
'                   5           LZMA    16MB     32       BT4   BCJ   ����ѹ��(Ĭ�ϵȼ���
'                   7           LZMA    32MB     64       BT4   BCJ   ���ѹ��
'                   9           LZMA    64MB     64       BT4   BCJ2  ����ѹ��
'       blnCompression=True-ѹ����False-��ѹ
'���أ��Ƿ�ɹ�
'˵������ѹ���ļ�������,��ɾ��ѹ��ԭʼ�ļ�
    Dim strCommand As String, strReturn As String
    '��ȡ����7Z�ļ�·������ֱ���˳�
    If gstr7ZPath = "" Then
        strErr = "7Z.EXE��ѹ���򲻴���"
        Exit Function
    End If
    If Not gobjFSO.FileExists(strSourceFile) Then
        strErr = "Դ�ļ�" & strSourceFile & "������"
        Exit Function
    End If
    If gobjFSO.FileExists(strDesFile) Then
        On Error Resume Next
        'ɾ�����ڵ�Ŀ���ļ�
        If FileSystem.GetAttr(strDesFile) <> vbNormal Then
             Call FileSystem.SetAttr(strDesFile, vbNormal)
        End If
        Call gobjFSO.DeleteFile(strDesFile, True)
        If Err.Number <> 0 Then Err.Clear
    End If
    On Error GoTo ErrH
    If blnCompression Then
        '-m �̶������ַ� x=���õȼ� mt������رն��߳�ѹ��ģʽ
        strCommand = """" & gstr7ZPath & """  a -y """ & strDesFile & """ """ & strSourceFile & """ -mx=" & intRate & " -mmt"
    Else
        '-o �̶������ַ�
        strCommand = """" & gstr7ZPath & """  e -y """ & strSourceFile & """ -o""" & gobjFSO.GetParentFolderName(strDesFile) & """"
    End If
    strReturn = RunCommand(strCommand, strErr, , 30000)
    If strErr = "" And strReturn <> "" Then strErr = strReturn
    If gobjFSO.FileExists(strDesFile) Then
        DeCompression = True
        If Not blnCompression Then
            On Error Resume Next
            If FileSystem.GetAttr(strSourceFile) <> vbNormal Then
                 Call FileSystem.SetAttr(strSourceFile, vbNormal)
            End If
            'ɾ��ԭʼ�ļ�
            Call gobjFSO.DeleteFile(strSourceFile, True)
            If Err.Number <> 0 Then Err.Clear
        End If
    End If
    Exit Function
ErrH:
    If strErr = "" Then strErr = Err.Description
    MsgBox Err.Description, vbInformation, App.Title
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
'���ܣ�����VB���֧�ֵİ汾����ʽ:9999.9999.9999,��С�汾��0000.0000.0000
'������strVer=��ǰ�汾��
'           blnMax=True,����Ϊ�գ��򷵻����֧�ְ汾��False=����Ϊ�գ��򷵻���С֧�ְ汾
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '����һ�Σ��Լ�������SP�汾��
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

Public Function IsVerSion(ByVal strVer As String) As Boolean
'���ܣ��ж��ַ����Ƿ��ǰ汾��
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If i = 3 Then
            If Format(Val(arrVer(i)), "0000") <> Format(Trim(arrVer(i)), "0000") Then Exit Function
        Else
            If Val(arrVer(i)) & "" <> Trim(arrVer(i)) Then Exit Function
        End If
    Next
    
    IsVerSion = True
End Function

Public Function TerminatePID(ByVal lngPid As Long) As Boolean
    '����:����ָ���Ľ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 11:06:16

    Dim lngProcess As Long, Phandle As Long, ret As Long
    TerminatePID = False
    
    On Error GoTo ErrHand:
    Phandle = OpenProcess(SYNCHRONIZE, False, lngPid)
    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
    Call TerminateProcess(lngProcess, 1&)
    ret = WaitForSingleObject(Phandle, INFINITE)
    ret = CloseHandle(Phandle)
    TerminatePID = True
ErrHand:
End Function

Public Function zlGetFileProcess(ByVal strFile As String, ByRef cllOutProcess As Collection) As Boolean
'����:��ȡָ���ļ�����ؽ���
'���:strFile-ָ����DLL�ļ�
'����:cllOutProcess-���ر����õĽ���ֵ
'����:�ɹ�,����true,���򷵻�False
'����:���˺�
'����:2009-01-20 13:59:35

    Dim uProcess As PROCESSENTRY32, uMdlInfor As MODULEENTRY32
    Dim lngMdlProcess As Long, strExeName As String, lngSnapShot As Long, strDLLName As String
    
    On Error GoTo ErrHand:
    '�������̿���
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
      uProcess.lSize = Len(uProcess)
      If Process32First(lngSnapShot, uProcess) Then
        Do
          '��ý��̵ı�ʶ��
          strExeName = UCase(Left(Trim(uProcess.sExeFile), InStr(1, Trim(uProcess.sExeFile), vbNullChar) - 1))
          If strExeName Like "*" & UCase(strFile) & "*" Then
             'һ����˵ֻ��Exe�ļ��Ż����
            On Error Resume Next
            cllOutProcess.Add Array(uProcess.lProcessId, strExeName, uProcess.lProcessId), "B" & uProcess.lProcessId
            If Err <> 0 Then
                cllOutProcess.Remove "B" & uMdlInfor.th32ProcessID
                cllOutProcess.Add Array(uProcess.lProcessId, strExeName, uProcess.lProcessId), "B" & uProcess.lProcessId
            End If
            On Error GoTo ErrHand:
          Else
                lngMdlProcess = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, uProcess.lProcessId)
                If lngMdlProcess > 0 Then
                    uMdlInfor.dwSize = Len(uMdlInfor)
                    If Module32First(lngMdlProcess, uMdlInfor) Then
                          Do
                                strDLLName = UCase(Left(Trim(uMdlInfor.szExePath), InStr(1, Trim(uMdlInfor.szExePath), vbNullChar) - 1))
                                If uProcess.lProcessId = uMdlInfor.th32ProcessID Then
                                    If strDLLName Like "*" & UCase(strFile) & "*" Then
                                        On Error Resume Next
                                        cllOutProcess.Add Array(uProcess.lProcessId, strExeName, uMdlInfor.th32ProcessID), "K" & uMdlInfor.th32ProcessID
                                        If Err <> 0 Then
                                            cllOutProcess.Remove "K" & uMdlInfor.th32ProcessID
                                            cllOutProcess.Add Array(uProcess.lProcessId, strExeName, uMdlInfor.th32ProcessID), "K" & uMdlInfor.th32ProcessID
                                        End If
                                        On Error GoTo ErrHand:
                                    End If
                                End If
                          Loop Until (Module32Next(lngMdlProcess, uMdlInfor) < 1)
                    End If
                    CloseHandle (lngMdlProcess)
                End If
            End If
        Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
      End If
      CloseHandle (lngSnapShot)
    End If
    zlGetFileProcess = True
    Exit Function
ErrHand:
End Function

'��������Ȩ��
Public Function EnablePrivilege(ByVal hProc As Long, ByVal strPrivilegeName As String) As Boolean
    Dim hToken As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    Dim lngRet As Long
    
    lngRet = OpenProcessToken(hProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
    lngRet = LookupPrivilegeValue(vbNullString, strPrivilegeName, tmpLuid)
    tkp.PrivilegeCount = 1
    tkp.Privileges(0).pLuid = tmpLuid
    tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    EnablePrivilege = AdjustTokenPrivileges(hToken, False, tkp, Len(tkp), tkpNewButIgnored, lBufferNeeded)
    CloseHandle hToken
End Function

Public Function Is64bit() As Boolean
'���ܣ��Ƿ���64λϵͳ
'���أ�
    Dim Handle As Long
    Dim blnFunc As Boolean
        
    blnFunc = False
    Handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If Handle > 0 Then
        IsWow64Process GetCurrentProcess(), blnFunc
    End If
    Is64bit = blnFunc
End Function

Public Function IsAdmin() As Boolean
    Dim hProcessToken       As Long
    Dim BufferSize          As Long
    Dim psidAdmin           As Long
    Dim lResult             As Long
    Dim X                   As Integer
    Dim tpTokens            As TOKEN_GROUPS
    Dim tpSidAuth           As SID_IDENTIFIER_AUTHORITY

    IsAdmin = False
    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
    ' Obtain current process token
    If Not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, hProcessToken) Then
        Call OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken)
    End If
    If hProcessToken Then
        ' Deternine the buffer size required
        Call GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) ' Determine required buffer size
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
            ' Retrieve your token information
            lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
            If lResult <> 1 Then Exit Function
            ' Move it from memory into the token structure
            Call CopyMemory(tpTokens, InfoBuffer(0), Len(tpTokens))
            ' Retreive the admins sid pointer
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
            If lResult <> 1 Then Exit Function
            If IsValidSid(psidAdmin) Then
                For X = 0 To tpTokens.GroupCount
                    ' Run through your token sid pointers
                    If IsValidSid(tpTokens.Groups(X).Sid) Then
                        ' Test for a match between the admin sid equalling your sid's
                        If EqualSid(ByVal tpTokens.Groups(X).Sid, ByVal psidAdmin) Then
                            IsAdmin = True
                            Exit For
                        End If
                    End If
                Next
            End If
            If psidAdmin Then Call FreeSid(psidAdmin)
        End If
        Call CloseHandle(hProcessToken)
    End If
End Function
'==========================================================
'˽�з���
'==========================================================
Private Function GetWNetErr(ByVal lngErr As Long) As String
    Dim strErr As String * 256
    Dim strName As String * 256
    Dim lngRet As Long
    lngRet = WNetGetLastError(lngErr, strErr, Len(strErr), strName, Len(strName))
    GetWNetErr = Replace(Replace("[" & TruncZero(strName) & "]" & TruncZero(strErr), Chr(10), ""), Chr(13), "")
End Function

Private Function ConvertAddressToString(longAddr As Long, Optional ByRef strErr As String) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    strErr = ""
    On Error GoTo ErrH
    CopyMemory myByte(0), longAddr, 4
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

Private Function Currency2LargeInteger(ByVal curDistance As Currency) As LARGE_INTEGER
    CopyMemory Currency2LargeInteger, curDistance, 8
End Function


Private Function LongToUnsigned(Value As Long) As Double
    If Value < 0 Then
        LongToUnsigned = Value + 2 ^ 32
    Else
        LongToUnsigned = Value
    End If
End Function

Private Function MD5File(f As String) As String
    Dim R As String * 32
    R = Space(32)
    MDFile f, R
    MD5File = UCase(R)
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

Public Function XCopy(ByVal strSourceFolder As String, ByVal strDesFolder As String) As Boolean
'���ܣ����ļ����Լ�����Ŀ¼���Ƶ���һ��Ŀ¼
    Dim objFile As File, objFolder As Folder
    
    On Error Resume Next
    For Each objFolder In gobjFSO.GetFolder(strSourceFolder).SubFolders
        If Not gobjFSO.FolderExists(strDesFolder & "\" & objFolder.Name) Then
            Call gobjFSO.CreateFolder(strDesFolder & "\" & objFolder.Name)
        End If
        Call XCopy(strSourceFolder & "\" & objFolder.Name, strDesFolder & "\" & objFolder.Name)
    Next
    
    For Each objFile In gobjFSO.GetFolder(strSourceFolder).Files
        If gobjFSO.FileExists(strDesFolder & "\" & objFile.Name) Then
            If FileSystem.GetAttr(strDesFolder & "\" & objFile.Name) <> vbNormal Then
                FileSystem.SetAttr strDesFolder & "\" & objFile.Name, vbNormal
            End If
        End If
        gobjFSO.CopyFile objFile.Path, strDesFolder & "\" & objFile.Name, True
    Next
    XCopy = Err.Number = 0
    If Err.Number <> 0 Then Err.Clear
End Function
