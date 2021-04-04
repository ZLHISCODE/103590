Attribute VB_Name = "mdlPublic"
Option Explicit

Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public glngTXTProc As Long '保存默认的消息函数的地址

Public Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Function GetGUID() As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim udtGUID As GUID
    
    On Error GoTo ErrHand
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
                String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
                String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
                IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
                IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
                IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
                IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
                IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
                IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
                IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
                IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
    
    Exit Function
ErrHand:
    'MsgBox Err.Description
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function CDateEx(ByVal strDate As String) As Date
    '---------------------------------------------------------------------------------------
    ' 功能 : 将一个不带日期格式时间字符串，格式化为有效格式的日期时间
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/5/21 17:26
    '---------------------------------------------------------------------------------------
    Dim intLimit As Integer
    Dim strDateTime As String, strTmp As String
    On Error GoTo ErrHand
    If Not IsNumeric(strDate) Then CDateEx = CDate(0): Exit Function
    
    strDateTime = Left(strDate, 4)  '年
    strDate = Mid(strDate, 5)
    If Val(strDateTime) < 1900 Or Val(strDateTime) > 3000 Then CDateEx = CDate(0): Exit Function
    
    If Val(Left(strDate, 2)) > 12 Then '月
        strDateTime = strDateTime & "-" & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    Else
        strDateTime = strDateTime & "-" & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    End If
    
    If IsDate(strDateTime & "-" & Left(strDate, 2)) Then '日
        strDateTime = strDateTime & "-" & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    Else
        strDateTime = strDateTime & "-" & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    End If
    
    If Val(Left(strDate, 2)) > 24 Then '时
        strDateTime = strDateTime & " " & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    Else
        strDateTime = strDateTime & " " & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    End If
    
    If Val(Left(strDate, 2)) > 60 Then '分
        strDateTime = strDateTime & ":" & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    Else
        strDateTime = strDateTime & ":" & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    End If
    
    If Val(Left(strDate, 2)) > 60 Then '秒
        strDateTime = strDateTime & ":" & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    Else
        strDateTime = strDateTime & ":" & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    End If
    
    CDateEx = strDateTime
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
