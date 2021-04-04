Attribute VB_Name = "mdlDefendGetText"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = -4
Private Const GWL_GETTEXT = 13

Private mlngOffset As Long

Public Function HookDefend(ByVal hwnd As Long) As Long
    'ָ���Զ���Ĵ��ڹ���
    'hwnd- ����ľ��
    mlngOffset = GetWindowLong(hwnd, GWL_WNDPROC)
    SetWindowLong hwnd, GWL_WNDPROC, AddressOf DefendFromSpy
    
    HookDefend = mlngOffset
End Function

Public Function DefendFromSpy(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '�Զ����ı����庯��,��ֹ�ⲿ����GetText
    '˵��:��������ҹ������󣬵�����Ϣ�����ں�ϵͳ�ͻ�����������
    
    If uMsg = GWL_GETTEXT Then
        Exit Function
    End If
    
    '�ָ�ԭ���Ĵ��ڹ���,��һ��ز�����
    DefendFromSpy = CallWindowProc(mlngOffset, hw, uMsg, wParam, lParam)
End Function
