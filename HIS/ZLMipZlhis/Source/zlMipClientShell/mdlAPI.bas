Attribute VB_Name = "mdlAPI"
Option Explicit

Public Type PointAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const SND_ASYNC = &H1   ' Play asynchronously (see note below).播放时可以进行窗体的其它
Public Const SND_NODEFAULT = &H2 ' Do not use default sound.
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_COPYRETURNORG = &H4
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const ILC_COLOR4 = &H4
Public Const ILC_COLOR8 = &H8
Public Const ILC_COLOR16 = &H10
Public Const ILC_COLOR24 = &H18
Public Const ILC_COLOR32 = &H20
Public Const ILD_NORMAL = 0

Public Const PI    As Double = 3.14159265358979
Public Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
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

Public Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)

' Constants
Public Const SM_CXFULLSCREEN = 16   ' Width of window client area
Public Const SM_CYFULLSCREEN = 17   ' Height of window client area
Public Const SND_SYNC = &H0
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
