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
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    

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

Public dss As COPYDATASTRUCT        '传递字符串消息的内存结构

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'窗口显示到最前端
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOP = 0 ' {在前面}
Public Const HWND_BOTTOM = 1 ' {在后面}
Public Const HWND_TOPMOST = -1 '{在前面, 位于任何顶部窗口的前面}
Public Const HWND_NOTOPMOST = -2 '{在前面, 位于其他顶部窗口的后面}

'uFlags 参数可选值:
'SWP_NOSIZE = 1; {忽略 cx、cy, 保持大小}
'SWP_NOMOVE = 2; {忽略 X、Y, 不改变位置}
'SWP_NOZORDER = 4; {忽略 hWndInsertAfter, 保持 Z 顺序}
'SWP_NOREDRAW = 8; {不重绘}
'SWP_NOACTIVATE = $10; {不激活}
'SWP_FRAMECHANGED = $20; {强制发送 WM_NCCALCSIZE 消息, 一般只是在改变大小时才发送此消息}
'SWP_SHOWWINDOW = $40; {显示窗口}
'SWP_HIDEWINDOW = $80; {隐藏窗口}

Public Function Hook(ByVal hWnd As Long) As Long
    '指定自定义的窗口过程
    '返回并保存原来默认的窗口过程指针

    Hook = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Function

Public Sub Unhook(ByVal hWnd As Long, ByVal lpWndProc As Long)
    Dim temp As Long
  
    temp = SetWindowLong(hWnd, GWL_WNDPROC, lpWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'消息处理程序，专门处理特定的 WM_COPYDATA 消息
    If uMsg = WM_COPYDATA Then
        If wParam = 223 Then Call subMsgCopyData(lParam)
    End If
  
    '调用原来的窗口过程
    WindowProc = CallWindowProc(glngPreWndProc, hw, uMsg, wParam, lParam)
End Function

Private Sub subMsgCopyData(ByVal lParam As Long)
'复制和分发消息

    Dim buf(1 To 1024) As Byte
    Dim strMsg As String
    
    On Error GoTo err
    
    '把lParam的内容复制到结构中
    Call CopyMemory(dss, ByVal lParam, Len(dss))
    
    Call WriteCommLog("zlSoftViewImage.subMsgCopyData", "复制和分发消息", "参数为：dss.dwData = " & dss.dwData, ltDebug)
    
    '如果没有主消息窗口，直接退出
    If gfrmViewImage Is Nothing Then
        Exit Sub
    End If
        
    If dss.dwData = 32 Then '关闭所有窗口
        Call CloseAllForms
    ElseIf dss.dwData = 33 Then
        '只处理dwData=33的消息
        '把lpData的内容复制到buf中
        Call CopyMemory(buf(1), ByVal dss.lpData, dss.cbData)
        strMsg = StrConv(buf, vbUnicode)
        strMsg = Left$(strMsg, InStr(1, strMsg, Chr$(0)) - 1)
        
        Call ProcessMessage(strMsg)

    End If
    Exit Sub
err:
    Call WriteCommLog("zlSoftViewImage.subMsgCopyData", "复制和分发消息，出现错误", err.Description, ltError)
End Sub

