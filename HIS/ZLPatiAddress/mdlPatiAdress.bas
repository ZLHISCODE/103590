Attribute VB_Name = "mdlPatiAdress"
Option Explicit
Public glngTXTProc As Long '��ֹ�Ҽ��˵�
Public gblnCanPaste As Boolean
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const WM_PASTE = &H302 'Ӧ�ó����ʹ���Ϣ��һ���༭���ComboBox�ԴӼ������еõ�����
Public gobjPati As PatiAddress
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_WNDPROC = -4&
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessageMenu(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then
        WndMessageMenu = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
    Else
        Call gobjPati.PopMenu
    End If
End Function

Public Function WndMessagePaste(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_Paste���͵���Ĭ�ϵĴ��ں�������
    If msg = WM_PASTE Then
        If Not gblnCanPaste Then  '�ṹ�����Ʋ�������
        Else
            WndMessagePaste = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
        End If
    Else
        WndMessagePaste = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
    End If
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
 '���:strInfor-ԭ��
 '         lngStart-ֱʼλ��
'         lngLen-����
'����:�Ӵ�
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo errH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
errH:
    Err.Clear
    SubB = ""
End Function
