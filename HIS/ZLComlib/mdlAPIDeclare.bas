Attribute VB_Name = "mdlAPIDeclare"
Option Explicit

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

'API定义
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

Public Const SPI_GETWORKAREA = 48
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const WM_NCACTIVATE = &H86
Public Const WM_CLOSE = &H10
Private mlngOldProc As Long
'OpenDir初始路径设置
Public gstrAPIPath As String
Public gMonitors() As Monitorinfos

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT
    Err = 0: On Error GoTo Errhand:
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
Errhand:
End Function


Public Function WindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_NCACTIVATE Then
        SendMessage hWnd, WM_CLOSE, wParam, lParam
    End If
    WindowProc = CallWindowProc(mlngOldProc, hWnd, Msg, wParam, lParam)
End Function

Public Sub Hook(ByVal hWnd As Long)
    mlngOldProc = GetWindowLong(hWnd, GWL_WNDPROC)
    SetWindowLong hWnd, GWL_WNDPROC, AddressOf WindowProc
End Sub

Public Sub UnHook(ByVal hWnd As Long)
    SetWindowLong hWnd, GWL_WNDPROC, mlngOldProc
End Sub

Public Function OpenDirCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
 '功能：claCommFun.OpenDir回调函数，用来设置打开的文件的初始路径
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
  
    On Error Resume Next
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal gstrAPIPath)
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH * 2)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, ByVal sBuffer)
            End If
    End Select
    
    OpenDirCallbackProc = 0
End Function
