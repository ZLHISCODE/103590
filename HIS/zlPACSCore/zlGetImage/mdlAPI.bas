Attribute VB_Name = "mdlAPI"
Option Explicit

'进程间传递内存空间，可以传字符串
Public Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type


Public Const WM_COPYDATA As Long = &H4A

'读取字符串消息
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long


'截获系统消息
Private Const GWL_WNDPROC = -4
Public Const GWL_USERDATA = (-21)
Public Const WM_SIZE = &H5
Public Const WM_USER = &H400
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    

'共享文件夹
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Public Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Public Type NETRESOURCE ' 网络资源
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
    '指定自定义的窗口过程
    '返回并保存原来默认的窗口过程指针
  Hook = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Function

Public Sub Unhook(ByVal hwnd As Long, ByVal lpWndProc As Long)
  Dim temp As Long
  
  temp = SetWindowLong(hwnd, GWL_WNDPROC, lpWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'消息处理程序，专门处理特定的 WM_COPYDATA 消息
    If uMsg = WM_COPYDATA Then
        If wParam = 123 Then Call subMsgCopyData(lParam)
    End If
  
    '调用原来的窗口过程
    WindowProc = CallWindowProc(plngPreWndProc, hw, uMsg, wParam, lParam)
End Function

Private Sub subMsgCopyData(ByVal lParam As Long)
    Dim buf(1 To 1024) As Byte
    Dim strMsg As String
    
    On Error GoTo err
    
    '把lParam的内容复制到结构中
    Call CopyMemory(dss, ByVal lParam, Len(dss))
    
    If dss.dwData = 3 Then
        '只处理dwData=3的消息
        '把lpData的内容复制到buf中
'
        Call CopyMemory(buf(1), ByVal dss.lpData, dss.cbData)
        strMsg = StrConv(buf, vbUnicode)
        strMsg = Left$(strMsg, InStr(1, strMsg, Chr$(0)) - 1)
        
        Call MsgInQueue(strMsg)

    End If
    Exit Sub
err:
    '无处理，出错就退出这个过程
End Sub
