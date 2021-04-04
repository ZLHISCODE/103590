Attribute VB_Name = "mdlHookKey"
Option Explicit


'Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long



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

'鼠标位置
Public Type POINTAPI1
    X As Long
    Y As Long
End Type


Public Type MSLLHOOKSTRUCT
    pt As POINTAPI1
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type


Public Type KBDLLHOOKSTRUCT
    VKCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type


Public hookKey As Object



Public Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'************************************************************************************************
'线程钩子回调函数
'
'nCode:
'wParam:消息类型
'lParam:数据
'
'************************************************************************************************
    On Error Resume Next
    HookProc = hookKey.HookProcess(nCode, wParam, lParam)
End Function




