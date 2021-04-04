Attribute VB_Name = "mdlNotify"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2019/3/19
'ģ��           mdlNotify
'˵��
'==================================================================================================
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
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

'Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
'Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'
'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Private Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type
'Private Type POINTAPI
'        x As Long
'        y As Long
'End Type
'Private Const RDW_INVALIDATE = &H1
'Private Const RDW_ERASE = &H4
'Private Const RDW_UPDATENOW = &H100
'Private Const SM_CXSMICON = 49
'Private Const SM_CYSMICON = 50
'
'Public Sub RemoveDeadIconFromSysTray()
'    Dim TrayWindow As Long
'    Dim WindowRect As RECT
'    Dim SmallIconWidth As Long
'    Dim SmallIconHeight As Long
'    Dim CursorPos As POINTAPI
'    Dim Row As Long
'    Dim Col As Long
'    '�������������ͱ߿�
'    TrayWindow = FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString)
'    If GetWindowRect(TrayWindow, WindowRect) = 0 Then Exit Sub
'    '���Сͼ���С
'    SmallIconWidth = GetSystemMetrics(SM_CXSMICON)
'    SmallIconHeight = GetSystemMetrics(SM_CYSMICON)
'    '���浱ǰ���λ��
'    Call GetCursorPos(CursorPos)
'    'ʹ�����ٻ���ÿ��ͼ��
'    For Row = 0 To (WindowRect.Bottom - WindowRect.Top) / SmallIconHeight
'        For Col = 0 To (WindowRect.Right - WindowRect.Left) / SmallIconWidth
'            Call SetCursorPos(WindowRect.Left + Col * SmallIconWidth, WindowRect.Top + Row * SmallIconHeight)
'            Call Sleep(10)  '��������ط�����Ϊ 0 ��ʱ����ʱ���ǲ�����
'        Next
'    Next
'    '�ָ����λ��
'    Call SetCursorPos(CursorPos.x, CursorPos.y)
'    '�ػ�������
'    Call RedrawWindow(TrayWindow, 0, 0, RDW_INVALIDATE Or RDW_ERASE Or RDW_UPDATENOW)
'End Sub


Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '���ܣ���������������һ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '���ܣ�����������ɾ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼�����������
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
    
End Sub
