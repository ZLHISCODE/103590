Attribute VB_Name = "mdlHookKey"
Option Explicit



Public hookKey As Object



Public Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'************************************************************************************************
'�̹߳��ӻص�����
'
'nCode:
'wParam:��Ϣ����
'lParam:����
'
'************************************************************************************************
    On Error Resume Next
    HookProc = hookKey.HookProcess(nCode, wParam, lParam)
End Function




