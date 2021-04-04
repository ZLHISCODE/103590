Attribute VB_Name = "mdlPatiAdress"
Option Explicit
Public glngTXTProc As Long '防止右键菜单
Public gblnCanPaste As Boolean
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const WM_PASTE = &H302 '应用程序发送此消息给一个编辑框或ComboBox以从剪贴板中得到数据
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
'去掉TextBox的默认右键菜单
Public Function WndMessageMenu(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then
        WndMessageMenu = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
    Else
        Call gobjPati.PopMenu
    End If
End Function

Public Function WndMessagePaste(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_Paste，就调用默认的窗口函数处理
    If msg = WM_PASTE Then
        If Not gblnCanPaste Then  '结构化复制不作处理
        Else
            WndMessagePaste = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
        End If
    Else
        WndMessagePaste = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
    End If
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'功能:读取指定字串的值,字串中可以包含汉字
 '入参:strInfor-原串
 '         lngStart-直始位置
'         lngLen-长度
'返回:子串
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo errH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
errH:
    Err.Clear
    SubB = ""
End Function
