Attribute VB_Name = "mdlTrayIcon"
Option Explicit
'------任务栏图标处理
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_ACTIVATE = &H6
Public Const WM_KEYDOWN = &H100
Public Const WM_PAINT = &HF

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'Public Const GWL_EXSTYLE = (-20)
'Public Const WinStyle = &H40000
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '功能：在任务栏上增加一个图标
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "仪器数据接收", strTip) & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub ModifyIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "", Optional ByVal blnMessage As Boolean = True)
    
    '功能：在任务栏上增加一个图标
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = IIf(blnMessage, WM_MOUSEMOVE, 0)
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "仪器数据接收", strTip) & Chr$(0)
    Shell_NotifyIcon NIM_MODIFY, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '功能：从任务栏上删除图标
    Dim t As NOTIFYICONDATA
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub


