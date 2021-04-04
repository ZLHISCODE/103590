Attribute VB_Name = "mdlAPI"
Option Explicit

'���̼䴫���ڴ�ռ䣬���Դ��ַ���
Public Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type


Public Const WM_COPYDATA As Long = &H4A

'��ȡ�ַ�����Ϣ
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long


'�ػ�ϵͳ��Ϣ
Private Const GWL_WNDPROC = -4
Public Const GWL_USERDATA = (-21)
Public Const WM_SIZE = &H5
Public Const WM_USER = &H400
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    

'�����ļ���
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Public Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Public Type NETRESOURCE ' ������Դ
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
Public Const RESOURCE_PUBLICNET = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const CONNECT_UPDATE_PROFILE = &H1

Public Function Hook(ByVal hwnd As Long) As Long
    'ָ���Զ���Ĵ��ڹ���
    '���ز�����ԭ��Ĭ�ϵĴ��ڹ���ָ��
  Hook = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Function

Public Sub Unhook(ByVal hwnd As Long, ByVal lpWndProc As Long)
  Dim temp As Long
  
  temp = SetWindowLong(hwnd, GWL_WNDPROC, lpWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��Ϣ�������ר�Ŵ����ض��� WM_COPYDATA ��Ϣ
    If uMsg = WM_COPYDATA Then
        If wParam = 123 Then Call subMsgCopyData(lParam)
    End If
  
    '����ԭ���Ĵ��ڹ���
    WindowProc = CallWindowProc(plngPreWndProc, hw, uMsg, wParam, lParam)
End Function

Private Sub subMsgCopyData(ByVal lParam As Long)
    Dim buf(1 To 1024) As Byte
    Dim strMsg As String
    
    On Error GoTo err
    
    '��lParam�����ݸ��Ƶ��ṹ��
    Call CopyMemory(dss, ByVal lParam, Len(dss))
    
    If dss.dwData = 3 Then
        'ֻ����dwData=3����Ϣ
        '��lpData�����ݸ��Ƶ�buf��
'
        Call CopyMemory(buf(1), ByVal dss.lpData, dss.cbData)
        strMsg = StrConv(buf, vbUnicode)
        strMsg = Left$(strMsg, InStr(1, strMsg, Chr$(0)) - 1)
        
        Call MsgInQueue(strMsg)

    End If
    Exit Sub
err:
    '�޴���������˳��������
End Sub
