Attribute VB_Name = "mdlPubDefine"
Option Explicit


'API常量定义
'********************************************************************
'---------------------------------------------------------------
'- 注册表 Api 常数...
'---------------------------------------------------------------
' Reg Data Types...
Public Const REG_SZ = 1                         ' Unicode空终结字符串
Public Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Public Const REG_DWORD = 4                      ' 32-bit 数字

' 注册表创建类型值...
Public Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留

' 注册表关键字安全选项...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Public Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' 返回值...
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0

'OpenFolder函数的回调函数使用
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const WM_USER = &H400
Public Const BFFM_SETSELECTION = (WM_USER + 102)
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_STATUSTEXT = &H4

'-----------------------------------------------------------------
'其他API常数
'-----------------------------------------------------------------
Public Const GW_CHILD = 5
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYFULLSCREEN = 17
Public Const LVM_SETCOLUMNWIDTH = &H101E
Public Const LVSCW_AUTOSIZE = -1
Public Const GW_HWNDNEXT = 2
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const Process_Query_Information = &H400
Public Const Still_Active = &H103
Public Const MAX_PATH = 260
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000 'Forces a top-level window onto the taskbar when the window is visible.强制一个可见的顶级视窗到工具栏上
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const TH32CS_SNAPPROCESS = &H2
Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const OFS_MAXPATHNAME = 128
Public Const OF_EXIST = &H4000
Public Const SND_SYNC = &H0    ' Play synchronously (default).  '播放时不能进行其它操作.只适用于非常短的声音
Public Const SND_ASYNC = &H1   ' Play asynchronously (see note below).播放时可以进行窗体的其它
Public Const SND_NODEFAULT = &H2 ' Do not use default sound.
Public Const SND_MEMORY = &H4  ' lpszSoundName points to a memory file.直接从内存中读取数据
Public Const SND_LOOP = &H8    ' Loop the sound until next sndPlaySound.
Public Const SND_NOSTOP = &H10 ' Do not stop any currently playing sound.
Public Const TVM_SETBKCOLOR = 4381&
Public Const TVM_GETBKCOLOR = 4383&
Public Const TVS_HASLINES = 2&
Public Const EM_EXGETSEL = WM_USER + 52
Public Const CB_SHOWDROPDOWN = &H14F
Public Const GWL_WNDPROC = -4
Public Const SB_TOP = 6
Public Const WM_VSCROLL = &H115
Public Const BF_SOFT = &H1000
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
'剪贴版数据格式定义
Public Const CF_TEXT = 1
Public Const CF_BITMAP = 2
Public Const CF_METAFILEPICT = 3
Public Const CF_SYLK = 4
Public Const CF_DIF = 5
Public Const CF_TIFF = 6
Public Const CF_OEMTEXT = 7
Public Const CF_DIB = 8
Public Const CF_PALETTE = 9
Public Const CF_PENDATA = 10
Public Const CF_RIFF = 11
Public Const CF_WAVE = 12
Public Const CF_UNICODETEXT = 13
Public Const CF_ENHMETAFILE = 14
Public Const CF_HDROP = 15
Public Const CF_LOCALE = 16
Public Const CF_MAX = 17

' 内存操作定义
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MODIFY = &H80
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_LOWER = GMEM_NOT_BANKED
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Const FO_COPY = &H2

Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const STARTF_USESTDHANDLES = &H100&
Public Const STARTF_USESHOWWINDOW = &H1
Public Const GENERIC_READ As Long = &H80000000
Public Const FILE_SHARE_READ As Long = &H1
Public Const OPEN_EXISTING As Long = 3
Public Const INVALID_HANDLE_VALUE As Long = (-1)
Public Const PAGE_READONLY As Long = &H2
Public Const PROV_RSA_FULL = 1
Public Const CRYPT_NEWKEYSET = &H8
Public Const SECTION_MAP_READ As Long = &H4
Public Const FILE_MAP_READ As Long = SECTION_MAP_READ
Public Const HP_HASHSIZE = 4
Public Const HP_HASHVAL = 2

Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const CONNECT_UPDATE_PROFILE = &H1
Public Const NO_ERROR = 0

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const CW_USEDEFAULT = &H80000000
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2


Public Const SB_HORZ = &H0
Public Const SB_VERT = &H1

Public Const WM_PASTE = &H302

'API类型参数定义
'********************************************************************
'读物理内存大小
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'注册表安全属性类型
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type
'进度结构
Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 1024
End Type

Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type DROPFILES
   pFiles As Long
   pt As POINTAPI
   fNC As Long
   fWide As Long
End Type

Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
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

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Public Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Public Type ZIPUSERFUNCTIONS
    DllPrnt As Long
    DLLPASSWORD As Long
    DLLCOMMENT As Long
    DLLSERVICE As Long
End Type

'ZPOPT is used to set options in the zip32.dll
Public Type ZPOPT
    fSuffix As Long
    fEncrypt As Long
    fSystem As Long
    fVolume As Long
    fExtra As Long
    fNoDirEntries As Long
    fExcludeDate As Long
    fIncludeDate As Long
    fVerbose As Long
    fQuiet As Long
    fCRLF_LF As Long
    fLF_CRLF As Long
    fJunkDir As Long
    fRecurse As Long
    fGrow As Long
    fForce As Long
    fMove As Long
    fDeleteEntries As Long
    fUpdate As Long
    fFreshen As Long
    fJunkSFX As Long
    fLatestTime As Long
    fComment As Long
    fOffsets As Long
    fPrivilege As Long
    fEncryption As Long
    fRepair As Long
    flevel As Byte
    date As String ' 8 bytes long
    szRootDir As String ' up to 256 bytes long
End Type

' Userfunctions structure
Public Type USERFUNCTION
    DllPrnt As Long
    DLLSND As Long
    DLLREPLACE As Long
    DLLPASSWORD As Long
    DLLMESSAGE As Long
    DLLSERVICE As Long
    TotalSizeComp As Long
    TotalSize As Long
    CompFactor As Long
    NumMembers As Long
    cchComment As Integer
End Type

' DCL structure
Public Type DCLIST
    ExtractOnlyNewer As Long
    SpaceToUnderscore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nUflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    fPrivilege As Long
    Zip As String
    ExtractDir As String
End Type

' Unzip32.dll version structure
Public Type UZPVER
    structlen As Long
    flag As Long
    beta As String * 10
    date As String * 20
    zlib As String * 10
    unzip(1 To 4) As Byte
    zipinfo(1 To 4) As Byte
    os2dll As Long
    windll(1 To 4) As Byte
End Type
'API声明
'********************************************************************
Public Declare Function apiOpenFile Lib "kernel32" Alias "OpenFile" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateFileA Lib "kernel32.dll" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As Long, ByRef lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function CryptAcquireContextA Lib "advapi32.dll" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Public Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As Long, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (pTcpTable As Any, pdwSize As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function messagebeep Lib "User32.dll" (ByVal wtype As Integer) As Integer
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long '判断数组是否为空
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySounda" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function StrCSpn Lib "shlwapi.dll" Alias "StrCSpnW" (ByVal lpStr&, ByVal lpCharacters&) As Long
Public Declare Function StrCSpnI Lib "shlwapi.dll" Alias "StrCSpnIW" (ByVal lpStr&, ByVal lpCharacters&) As Long
Public Declare Function StrRStr Lib "shell32.dll" Alias "StrRStrW" (ByVal lpStart&, ByVal lpEnd&, ByVal lpSrch&) As Long
Public Declare Function StrRStrI Lib "shell32.dll" Alias "StrRStrIW" (ByVal lpStart&, ByVal lpEnd&, ByVal lpSrch&) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Long
Public Declare Sub UzpVersion2 Lib "unzip32.dll" (uzpv As UZPVER)
Public Declare Function windll_unzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As ZIPnames, ByVal xfnc As Long, ByRef xfnv As ZIPnames, dcll As DCLIST, Userf As USERFUNCTION) As Long
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Public Declare Function ZpInit Lib "zip32.dll" (ByRef Zipfun As ZIPUSERFUNCTIONS) As Long ' Set Zip Callbacks
Public Declare Function ZpSetOptions Lib "zip32.dll" (ByRef Opts As ZPOPT) As Long ' Set Zip options
Public Declare Function ZpGetOptions Lib "zip32.dll" () As ZPOPT ' used to check encryption flag only
Public Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long ' Real zipping action

Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long

'公共部份菜单ID定义
'********************************************************************
Public Const conMenu_FilePopup = 1              '文件
Public Const conMenu_ManagePopup = 2            '管理
Public Const conMenu_EditPopup = 3              '编辑
Public Const conMenu_ReportPopup = 4            '报表

Public Const conMenu_ToolPopup = 5              '工具 2006-07-11 add by 陈东

Public Const conMenu_ViewPopup = 7              '查看
Public Const conMenu_HelpPopup = 9              '帮助

'文件菜单
Public Const conMenu_File_PrintSet = 101        '打印设置(&S)…
Public Const conMenu_File_Preview = 102         '预览(&V)
Public Const conMenu_File_Print = 103           '打印(&P)
Public Const conMenu_File_Excel = 104           '输出到&Excel…
Public Const conMenu_File_RemoveTools = 161     '卸载管理工具   2006-07-11 add by 陈东
Public Const conMenu_File_LogOut = 171           '注销         2006-07-11 add by 陈东
Public Const conMenu_File_Parameter = 181       '参数设置(&M)
Public Const conMenu_File_Exit = 191            '退出(&X)

'编辑菜单
Public Const conMenu_Edit_NewParent = 301       '新分类(&N)
Public Const conMenu_Edit_NewItem = 302         '新项目(&A)
Public Const conMenu_Edit_Modify = 303          '修改(&M)
Public Const conMenu_Edit_Delete = 304          '删除(&D)
Public Const conMenu_Edit_Audit = 305           '审核(&U)
Public Const conMenu_Edit_Blankoff = 306        '作废(&B)
Public Const conMenu_Edit_Disuse = 307          '停用(&P)
Public Const conMenu_Edit_Reuse = 308           '启用(&R)

'工具菜单  2006-07-11 add by 陈东
Public Const conMenu_Tool_LoadAndUnload = 501       '装卸管理(&I)
Public Const conMenu_Tool_DataMana = 502         '数据管理(&D)
Public Const conMenu_Tool_RunMana = 503          '运行管理(&E)
Public Const conMenu_Tool_Popedom = 504          '权限管理(&G)
Public Const conMenu_Tool_Expert = 505          '专项工具(&R)
Public Const conMenu_Tool_DBA = 506           'DBA工具

Public Const conMenu_Edit_Untread = 50502       '恢复过程
Public Const conMenu_Edit_Confirm = 5030306     '确认调整
Public Const conMenu_Process_Zoom = 50503       '差异检查
Public Const conMenu_Edit_Word = 5050502        '搜集更新
Public Const conMenu_Manage_Change_PaitNote = 50504  'conMenu_Manage_Change_PaitNote
Public Const conMenu_Edit_SaveExit = 50506      '完成
Public Const conMenu_Edit_Save = 50505          '暂存
'查看菜单
Public Const conMenu_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_View_StatusBar = 702            '状态栏(&S)
Public Const conMenu_View_Expend = 711               '展开/折叠组(&X)
Public Const conMenu_View_Expend_AllCollapse = 7111     '折叠所有组(&L)
Public Const conMenu_View_Expend_AllExpend = 7112       '展开所有组(&X)
Public Const conMenu_View_Expend_CurCollapse = 7113     '折叠当前组(&C)
Public Const conMenu_View_Expend_CurExpend = 7114       '展开当前组(&E)
Public Const conMenu_View_Filter = 721               '过滤(&G)
Public Const conMenu_View_Find = 722                 '查找(&F)
Public Const conMenu_View_FindNext = 723             '查找下一个(&N)
Public Const conMenu_View_Refresh = 791              '刷新(&R)

Public Const conMenu_View_ToolsList = 703            '工具列表 2006-07-11 add by 陈东
Public Const conMenu_View_ToolsPwd = 704            '非所有者用户的数据库密码

'帮助菜单
Public Const conMenu_Help_Help = 901        '帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Forum = 9023      '中联论坛(&F)
Public Const conMenu_Help_Web_Mail = 9022       '发送反馈(&M)
Public Const conMenu_Help_About = 991       '关于(&A)…

'其它常量定义
'********************************************************************
'CommandBar固有常量定义
Public Const XTP_ID_WINDOW_LIST = 35000 '窗体列表
Public Const XTP_ID_TOOLBARLIST = 59392 '工具栏列表
Public Const ID_INDICATOR_CAPS = 59137 '状态栏（大写）
Public Const ID_INDICATOR_NUM = 59138 '状态栏（数字）
Public Const ID_INDICATOR_SCRL = 59139 '状态栏（滚动）

'CommandBar辅助热键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16
'********************************************************************

Public Const MSTR_DBLINK_KEY As String = "zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325"  '密码加解密秘钥
