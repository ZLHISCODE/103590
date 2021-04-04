Attribute VB_Name = "mdlVSFlexGrid"
Option Explicit

'######################################################################################################################
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Enum Color
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E4E7
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
    
    原始单据 = 0
    冲销记录 = &HFF
    停用项目 = &H8000000C
    启用项目 = 0
    
    公共模块色 = &HC00000
    
    报警背景色 = &H40C0&
    报警前景色 = &H8000000E
    超标背景色 = &H80C0FF
    低标背景色 = &H80FFFF
    超标前景色 = &H80000012
    默认前景色 = &H80000008
    锁色 = &HF5F5F5
    启用色 = 0
    停用色 = 255
End Enum


Private Type POINTAPI
     X As Long
     Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Const SB_TOP = 6
Public Const WM_VSCROLL = &H115

Private mlngTXTProc As Long

Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Sub TxtSelAll(objTxt As Object)
'功能：将编辑框的的文本全部选中
'参数：objTxt=需要全选的编辑控件,该控件具有SelStart,SelLength属性
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    If TypeName(objTxt) = "TextBox" Then
        If objTxt.MultiLine Then
            SendMessage objTxt.hwnd, WM_VSCROLL, SB_TOP, 0
        End If
    End If
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
    '       实际数据存储长度
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function


Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByVal hwnd As Long = 0, Optional str项目 As String) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If str项目 = "" Then str项目 = "所输入内容"
    
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        MsgBox str项目 & "含有非法字符。", vbExclamation, ""
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox str项目 & "不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, ""
            If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
            Exit Function
        End If
    End If
    
    StrIsValid = True
End Function

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function SetNewWindowLong(ByVal lngHwnd As Long, ByVal dwNewLong As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngTXTProc = GetWindowLong(lngHwnd, GWL_WNDPROC)
    Call SetWindowLong(lngHwnd, GWL_WNDPROC, dwNewLong)
        
    SetNewWindowLong = True
    
End Function

Public Function RestoreWindowLong(ByVal lngHwnd As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call SetWindowLong(lngHwnd, GWL_WNDPROC, mlngTXTProc)
    
    RestoreWindowLong = True
End Function

Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    '******************************************************************************************************************
    '功能：去掉TextBox的默认右键菜单
    '参数：
    '返回：
    '说明：如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    '******************************************************************************************************************
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(mlngTXTProc, hwnd, msg, wp, lp)
End Function

Public Function ErrCenter() As Byte
    MsgBox Err.Description
End Function

