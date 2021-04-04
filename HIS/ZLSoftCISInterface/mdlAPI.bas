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
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    

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

Public dss As COPYDATASTRUCT        '�����ַ�����Ϣ���ڴ�ṹ

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'������ʾ����ǰ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'SM4����
'/**
' * \brief          SM4-ECB block encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param input    input block
' * \param output   output block
' */
Public Declare Function sm4_crypt_ecb Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, key As Byte, in_put As Byte, out_put As Byte) As Long

'��ȡZLSM4���޸İ汾
'1:ֻ֧��sm4_crypt_ecb,sm4_crypt_cbc
'2:����֧��sm3��sm3_file��sm3_hmac��sm_version
'/**
' * \brief          Output = zlSM4.DLL Version
' */
Public Declare Function get_sm_version Lib "zlSm4.dll" Alias "sm_version" () As Long

Public Const HWND_TOP = 0 ' {��ǰ��}
Public Const HWND_BOTTOM = 1 ' {�ں���}
Public Const HWND_TOPMOST = -1 '{��ǰ��, λ���κζ������ڵ�ǰ��}
Public Const HWND_NOTOPMOST = -2 '{��ǰ��, λ�������������ڵĺ���}

Public M_SM4_VERSION As Long
Public Const SM4_CRYPT_RANDOMIZE_KEY    As Long = 999  'sm4�����㷨��Կ���������������
Public Const SM4_CRYPT_RANDOMIZE_IV     As Long = 666   'sm4�����㷨��ʼ�������������������
Public Const G_PASSWORD_KEY             As String = "3357F1F2CA0341A5B75DBA7F35666715"

Public Enum CrypeMode
    CM_Encrypt = 1   '����
    CM_Decrypt = 0   '����
End Enum

'uFlags ������ѡֵ:
'SWP_NOSIZE = 1; {���� cx��cy, ���ִ�С}
'SWP_NOMOVE = 2; {���� X��Y, ���ı�λ��}
'SWP_NOZORDER = 4; {���� hWndInsertAfter, ���� Z ˳��}
'SWP_NOREDRAW = 8; {���ػ�}
'SWP_NOACTIVATE = $10; {������}
'SWP_FRAMECHANGED = $20; {ǿ�Ʒ��� WM_NCCALCSIZE ��Ϣ, һ��ֻ���ڸı��Сʱ�ŷ��ʹ���Ϣ}
'SWP_SHOWWINDOW = $40; {��ʾ����}
'SWP_HIDEWINDOW = $80; {���ش���}

Public Function Hook(ByVal hWnd As Long) As Long
    'ָ���Զ���Ĵ��ڹ���
    '���ز�����ԭ��Ĭ�ϵĴ��ڹ���ָ��

    Hook = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Function

Public Sub Unhook(ByVal hWnd As Long, ByVal lpWndProc As Long)
    Dim temp As Long
  
    temp = SetWindowLong(hWnd, GWL_WNDPROC, lpWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��Ϣ�������ר�Ŵ����ض��� WM_COPYDATA ��Ϣ
    If uMsg = WM_COPYDATA Then
        If wParam = 223 Then Call subMsgCopyData(lParam)
    End If
  
    '����ԭ���Ĵ��ڹ���
    WindowProc = CallWindowProc(plngPreWndProc, hw, uMsg, wParam, lParam)
End Function

Private Sub subMsgCopyData(ByVal lParam As Long)
'���ƺͷַ���Ϣ

    Dim buf(1 To 1024) As Byte
    Dim strmsg As String
    
    On Error GoTo err
    
    '��lParam�����ݸ��Ƶ��ṹ��
    Call CopyMemory(dss, ByVal lParam, Len(dss))
    
    
    '���û������Ϣ���ڣ�ֱ���˳�
    If mfrmShowHisForms Is Nothing Then
        Call CloseAllForms
        Exit Sub
    End If
        
    If dss.dwData = 32 Then '�ر����д���
        
        Call CloseAllForms
        
    ElseIf dss.dwData = 33 Then
        'ֻ����dwData=33����Ϣ
        '��lpData�����ݸ��Ƶ�buf��
        Call CopyMemory(buf(1), ByVal dss.lpData, dss.cbData)
        strmsg = StrConv(buf, vbUnicode)
        strmsg = Left$(strmsg, InStr(1, strmsg, Chr$(0)) - 1)
        
        Call ProcessMessage(strmsg)

    End If
    Exit Sub
err:
End Sub
