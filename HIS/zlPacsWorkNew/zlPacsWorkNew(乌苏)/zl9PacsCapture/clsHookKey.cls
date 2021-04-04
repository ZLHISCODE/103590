VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHookKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mhHook As Long
Private mlngActiveHwnd As Long        '当前窗口句柄
Private mblnIsOnlyActive As Boolean   '是否仅仅为当前窗口句柄时才执行hook处理
Private mlngMsgType As Long

Public Event OnKeyBoardLHook(ByVal lngMsg As Long, ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)


Public Sub EnableHook(Optional lngRegMsg As Long = 0, Optional ByVal blnIsGlobalHot As Boolean = False)
'************************************************************************************************
'注册线程钩子'
'************************************************************************************************

    mlngMsgType = lngRegMsg
    
    If mhHook = 0 Then
        'mhHook = SetWindowsHookEx(HOOK_WH_MOUSE_LL, AddressOf HookProc, App.hInstance, 0) '鼠标钩子
        'mhHook = SetWindowsHookEx(HOOK_WH_KEYBOARD, AddressOf HookProc, App.hInstance, 0)  '键盘钩子
        
        Set hookKey = Me
        
        '监视所有进程
        mhHook = SetWindowsHookEx(HOOK_WH_KEYBOARD_LL, AddressOf HookProc, App.hInstance, 0)  '键盘钩子    'IIf(blnIsGlobalHot, 0, App.hInstance)
        
    End If
End Sub

Public Sub FreeHook()
'************************************************************************************************
'释放线程钩子'
'************************************************************************************************
    If mhHook <> 0 Then
        Call UnhookWindowsHookEx(mhHook)
        mhHook = 0
        
        Set hookKey = Nothing
    End If
End Sub
    
    
    
Private Function GetKeyDataStruc(ByVal lParam As Long) As KBDLLHOOKSTRUCT
'返回按键状态
    Dim ks As KBDLLHOOKSTRUCT
    
    Call CopyMemory(VarPtr(ks), ByVal lParam, Len(ks))
    
    GetKeyDataStruc = ks
End Function

    
Public Function HookProcess(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'************************************************************************************************
'线程钩子回调函数
'
'nCode:
'wParam:消息类型
'lParam:数据
'
'************************************************************************************************
    On Error GoTo errHandle
    Dim pid As Long
    Dim ks As KBDLLHOOKSTRUCT
    
    
BUGEX "HookProcess 1"
    If (nCode <> 0) Then
        HookProcess = CallNextHookEx(mhHook, nCode, wParam, lParam)
        Exit Function
    End If

BUGEX "HookProcess 2"

'    '判断触发消息的是否为当前进程
'    Call GetWindowThreadProcessId(GetActiveWindow(), pid)
'
'    If GetCurrentProcessId = pid Then
    
BUGEX "HookProcess 3"
    If mlngMsgType <> 0 And wParam <> mlngMsgType Then
        HookProcess = CallNextHookEx(mhHook, nCode, wParam, lParam)
        Exit Function
    End If
    
        ks = GetKeyDataStruc(lParam)

        RaiseEvent OnKeyBoardLHook(wParam, ks.VKCode, ks.scanCode, ks.flags)
'    End If
    
BUGEX "HookProcess End"
errHandle:
    HookProcess = CallNextHookEx(mhHook, nCode, wParam, lParam)
End Function



'---------------------------------------------------------------


Public Property Let ActiveHwnd(lngHwnd As Long)
    mlngActiveHwnd = lngHwnd
End Property

Public Property Get ActiveHwnd() As Long
    ActiveHwnd = mlngActiveHwnd
End Property

Public Property Let IsOnlyActive(blnIsOnlyActive As Boolean)
    mblnIsOnlyActive = blnIsOnlyActive
End Property

Public Property Get IsOnlyActive() As Boolean
    IsOnlyActive = mblnIsOnlyActive
End Property



Private Sub Class_Terminate()
    'Set hookKey = Nothing
End Sub
