Attribute VB_Name = "mdlHookKey"
Option Explicit



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




