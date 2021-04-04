Attribute VB_Name = "mdlApi"
Option Explicit


'使用API函数修改MsgBox，使其可以在调用的时候，指定父窗体
Public Const HWND_BROADCAST = &HFFFF&

Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_APPLMODAL = &H0&
Public Const MB_COMPOSITE = &H2         '  use composite chars
Public Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Public Const MB_DEFBUTTON1 = &H0&
Public Const MB_DEFBUTTON2 = &H100&
Public Const MB_DEFBUTTON3 = &H200&
Public Const MB_DEFMASK = &HF00&
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONMASK = &HF0&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONSTOP = MB_ICONHAND
Public Const MB_MISCMASK = &HC000&
Public Const MB_MODEMASK = &H3000&
Public Const MB_NOFOCUS = &H8000&
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_PRECOMPOSED = &H1         '  use precomposed chars
Public Const MB_RETRYCANCEL = &H5&
Public Const MB_SETFOREGROUND = &H10000
Public Const MB_SYSTEMMODAL = &H1000&
Public Const MB_TASKMODAL = &H2000&
Public Const MB_TYPEMASK = &HF&
Public Const MB_USEGLYPHCHARS = &H4         '  use glyph chars, not ctrl chars
Public Const MB_YESNO = &H4&
Public Const MB_YESNOCANCEL = &H3&

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const ETO_OPAQUE = 2
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_CLOSE = &H10
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const WM_GETMINMAXINFO = &H24
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Const WS_EX_APPWINDOW = &H40000

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000


Public Type PointAPI
        x As Long
        y As Long
End Type


Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


'注册热键

Public Const HOOK_HC_ACTION = 0
Public Const HOOK_WH_JOURNALRECORD = 0
Public Const HOOK_WH_KEYBOARD As Long = 2
Public Const HOOK_WH_MOUSE_LL As Long = 14   '监视输入到线程消息队列中的鼠标消息
Public Const HOOK_WH_KEYBOARD_LL As Long = 13 '监视输入到线程消息队列中的键盘消息
Public Const HOOK_WM_MOUSEMOVE = &H200
Public Const HOOK_WM_LBUTTONDOWN = &H201
Public Const HOOK_WM_LBUTTONUP = &H202
Public Const HOOK_WM_LBUTTONDBLCLK = &H203
Public Const HOOK_WM_RBUTTONDOWN = &H204
Public Const HOOK_WM_RBUTTONUP = &H205
Public Const HOOK_WM_RBUTTONDBLCLK = &H206
Public Const HOOK_WM_MBUTTONDOWN = &H207
Public Const HOOK_WM_MBUTTONUP = &H208
Public Const HOOK_WM_MBUTTONDBLCLK = &H209
Public Const HOOK_WM_MOUSEACTIVATE = &H21
Public Const HOOK_WM_MOUSEFIRST = &H200
Public Const HOOK_WM_MOUSELAST = &H209
Public Const HOOK_WM_MOUSEWHEEL = &H20A '以上是鼠标的各个值


Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4 'MENU=ALT
Public Const VK_RMENU = &HA5
Public Const HC_ACTION = &H0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101


Public Type MSLLHOOKSTRUCT
    pt As PointAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Type KeyMsgs
    lngVirtualKey As Long
    lngScanKey As Long
    lngFlag As Long
    lngTime As Long
End Type

Public Type KBDLLHOOKSTRUCT
    VKCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4


'快速设置ComboBox数据
Public Const CB_ADDSTRING = &H143
Public Const CB_SETITEMDATA = &H151
Public Const CB_SETCURSEL = &H14E

'快速清除TreeView
Public Const WM_SETREDRAW = &HB


'处理鼠标滚轮
Public Const WM_MOUSEWHEEL = &H20A
Public Type POINTL
    x As Long
    y As Long
End Type

'判断滚动条的可见性
Public Const SB_VERT = &H1
Public Type SCROLLINFO
    cbSize   As Long
    npos   As Long
    ntrackpos   As Long
    nmax   As Long
    npage   As Long
    nmin   As Long
    fmask   As Long
End Type

Public Const WS_HSCROLL = &H100000
Public Const WS_VSCROLL = &H200000

' 网络资源
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
Public Const RESOURCE_PUBLICNET = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const CONNECT_UPDATE_PROFILE = &H1

Public Type MINMAXINFO
        ptReserved As PointAPI
        ptMaxSize As PointAPI
        ptMaxPosition As PointAPI
        ptMinTrackSize As PointAPI
        ptMaxTrackSize As PointAPI
End Type

Public Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As PointAPI
End Type


'读取网卡的多个IP
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1

Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128


Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type


Public Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type



Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type


Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(256) As RGBQUAD
End Type

Public Type TBitMap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type



Public Const WH_CBT = 5
Public Const HCBT_ACTIVATE = 5
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7
Public Const IDPROMPT = &HFFFF&
 

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

Public Const GA_PARENT = 1      '获取父窗口。这不包括所有者，功能同GetParent功能
Public Const GA_ROOT = 2        '通过遍历父窗口链获取根窗口
Public Const GA_ROOTOWNER = 3   '通过遍历父窗口链和使用GetParent函数返回的所有者窗口来获取根窗口

Public Const VK_LCTRL = &HA2
Public Const VK_RCTRL = &HA3

Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12


Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long

'消息发送
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAsAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowTextAsLong Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal LPCSTR As Long) As Long ' C BOOL
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWndChild As Long) As Long
Public Declare Function GetAncestor Lib "user32" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetComboData Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindComboStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Public Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Public Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
 
Public Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByRef lpMinPos As Long, ByRef lpMaxPos As Long) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, ByRef lpscrollinfo As SCROLLINFO) As Long


Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long


Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long



Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Public Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)


Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

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
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' 返回值...
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- 注册表安全属性类型...
'---------------------------------------------------------------
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type


'---------------------------------------------------------------
'- 注册表API定义...
'---------------------------------------------------------------
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long



'---------------------------------------------------------------
'- 监视器操作API...
'---------------------------------------------------------------

Public Const MONITORINFOF_PRIMARY = &H1
Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const MONITOR_DEFAULTTONULL = &H0
Public Const MONITOR_DEFAULTTOPRIMARY = &H1
  
'显示器信息
Public Type MONITORINFO
        cbSize   As Long
        rcMonitor   As RECT
        rcWork   As RECT
        dwFlags   As Long
End Type


Public Type Monitorinfos
  monitorHandle As Long
  monitorInf As MONITORINFO
End Type
 

Public Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Public Declare Function MonitorFromPoint Lib "user32.dll" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Public Declare Function MonitorFromRect Lib "user32.dll" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Public Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
Public Declare Function EnumDisplayMonitors Lib "user32.dll" (ByVal hdc As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long

