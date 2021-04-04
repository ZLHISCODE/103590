Attribute VB_Name = "mdlApiDeclare"
Option Explicit
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type



'常数声明
Public glngTXTProc As Long '保存默认的消息函数的地址
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息

Public Const GWL_EXSTYLE = (-20)

Public Const CB_GETCURSEL = &H147
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_GETCOUNT = &H146

Public Const SB_TOP = 6
Public Const WM_VSCROLL = &H115
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
Public Const EM_GETFIRSTVISIBLELINE = &HCE 'lngR(>=0)
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
Public Const GWL_STYLE = (-16)
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCaL_MaCHINE = &H80000002
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const KLF_REORDER = &H8
Public Const LVM_SETCOLUMNWIDTH = &H101E
Public Const LVSCW_AUTOSIZE = -1
Public Const MAX_PATH = 256
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_SHOWWINDOW = &H40
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9

Public Const MONITORINFOF_PRIMARY = &H1
Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const MONITOR_DEFAULTTONULL = &H0
Public Const MONITOR_DEFAULTTOPRIMARY = &H1
  
'OpenDir函数的回调函数使用
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const WM_USER = &H400
Public Const BFFM_SETSELECTION = (WM_USER + 102)
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)

Public Const WM_MOUSEWHEEL = &H20A '鼠标滚动消息

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type


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
Public Const WM_GETMINMAXINFO = &H24
Public Const WH_KEYBOARD = 2

 
Public Const SPI_GETWORKAREA = 48

Public Const GWL_HWNDPARENT = (-8)
 
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type


Private Type TY_WindowsRect
    MaxW As Long
    MaxH As Long
    MinW  As Long
    MinH As Long
End Type
Public gWinRect As TY_WindowsRect


Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'输入法控制API----------------------------------------------------------------------------------------------
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long

'API定义
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function GetMonitorInfo Lib "User32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Public Declare Function MonitorFromPoint Lib "User32.dll" (ByVal X As Long, ByVal Y As Long, ByVal dwFlags As Long) As Long
Public Declare Function MonitorFromRect Lib "User32.dll" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Public Declare Function MonitorFromWindow Lib "User32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Public Declare Function EnumDisplayMonitors Lib "User32.dll" (ByVal hDC As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
'打开文件夹，并设置初始路径
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const WM_NCACTIVATE = &H86
Public Const WM_CLOSE = &H10
Private mlngOldProc As Long
'OpenDir初始路径设置
Public gstrAPIPath As String
Public gMonitors() As Monitorinfos
