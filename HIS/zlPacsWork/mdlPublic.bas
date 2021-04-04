Attribute VB_Name = "mdlPublic"
Option Explicit
Public lngTXTProc As Long '保存默认的消息函数的地址
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long型最大值

Public Const ETO_OPAQUE = 2
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const WM_GETMINMAXINFO = &H24
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWndChild As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'''''''''''''''''''''''''''''''''''''''''''''设置全局热键'''''''''''''''''''''''''''''''''''''''''''
Public Declare Function RegisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long) As Long
Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4

Public preWinProc As Long
Public Modifiers As Long, uVirtKey As Long, idHotKey As Long

Private Type taLong
    ll As Long
End Type

Private Type t2Int
    lWord As Integer
    hWord As Integer
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'快速设置ComboBox数据
Public Const CB_ADDSTRING = &H143
Public Const CB_SETITEMDATA = &H151
Public Const CB_SETCURSEL = &H14E

Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetComboData Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindComboStr Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

'DICOM图象参数
Public Const ATTR_检查日期 As String = "8:20"
Public Const ATTR_检查时间 As String = "8:30"
Public Const ATTR_影像类别 As String = "8:60"
Public Const ATTR_检查设备 As String = "8:1090"

Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
    Dim vRect As RECT, vPos As POINTAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.x >= vRect.Left And vPos.x <= vRect.Right _
        And vPos.y >= vRect.Top And vPos.y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.Hwnd, vRect)
    lngStyle = GetWindowLong(objForm.Hwnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.Hwnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.Hwnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function MoveObj(lngHwnd As Long) As RECT
'功能：在对象的MouseDown事件中调用,对象必须具有Hwnd属性
'返回：相对屏幕的像素值
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
'功能：在下拉式工具按钮中弹出一个菜单
    Dim vRect As RECT, vDot1 As POINTAPI, vDot2 As POINTAPI
    
    Call GetWindowRect(ToolBar.Hwnd, vRect)
    vDot1.x = vRect.Left: vDot1.y = vRect.Top
    vDot2.x = vRect.Right: vDot2.y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.Hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.Hwnd, vDot2)
    
    vDot1.x = vDot1.x * 15: vDot1.y = vDot1.y * 15
    vDot2.x = vDot2.x * 15: vDot2.y = vDot2.y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.x + Button.Left, vDot2.y
End Sub

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal LngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.x = lngX / Screen.TwipsPerPixelX: vPoint.y = LngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.x = vPoint.x * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'功能：将VB的系统颜色转换为RGB色
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal Hwnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, Hwnd, Msg, wp, lp)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, 500, varValue)
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, 500, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

    '执行返回记录集
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function

Public Function OpenRecord(rsTmp As ADODB.Recordset, strSQL As String, ByVal strTitle As String, _
    Optional CursorType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, strTitle, strSQL)
    rsTmp.Open strSQL, gcnOracle, CursorType, LockType
    Call SQLTest
    
    Set OpenRecord = rsTmp
End Function

Public Sub ExecuteProc(ByVal strSQL As String, ByVal strCaption As String)
'功能：执行SQL语句
'    Call SQLTest(App.ProductName, strCaption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    Call zl9comlib.SQLTest(App.ProductName, strCaption, strSQL)
    Call zl9comlib.zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Call zl9comlib.SQLTest
End Sub

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function GetFullDate(ByVal strText As String) As String
'功能：根据输入的日期简串,返回完整的日期串(yyyy-MM-dd HH:mm)
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = zlDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '输入串中包含日期分隔符
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                '只输入了日期部份
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                '只输入了时间部份
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '输入非法日期,返回原内容
            strTmp = strText
        End If
    Else
        '不包含日期分隔符
        If Len(strTmp) <= 2 Then
            '当作输入dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '当作输入MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '当作输入yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '当作输入MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '当作输入yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '当作输入yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    GetFullDate = strTmp
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Public Function StringMask(ByVal strText As String, ByVal strMask As String) As Boolean
'功能：检查字符串是否只包含指定的字符
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strMask, Mid(strText, i, 1)) = 0 Then Exit Function
    Next
    StringMask = True
End Function

Public Function ExeTimeValid(ByVal strTime As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String) As Boolean
'功能：检查指定的执行时间是否合法
    Dim arrTime() As String, strTmp As String, i As Integer
    Dim strPreTime As String, intPreDay As Long, intCurDay As Long
    
    If strTime = "" Then Exit Function
    
    If str间隔单位 = "周" Then
        '1/8:00-3/15:00-5/9:00；1/8:00-3/15-5/9:00
        If Not StringMask(strTime, "0123456789:-/") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> int频率次数 Then Exit Function
        
        For i = 0 To UBound(arrTime)
            If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
            '星期部份
            strTmp = Split(arrTime(i), "/")(0)
            If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
            intCurDay = Val(strTmp)
            If intCurDay < 1 Or intCurDay > 7 Then Exit Function
            If intPreDay <> 0 Then
                If intCurDay < intPreDay Then Exit Function
            End If
            
            '绝对时间部分
            strTmp = Split(arrTime(i), "/")(1)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Then Exit Function
            If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            
            strPreTime = Format(strTmp, "HH:mm")
            intPreDay = intCurDay
        Next
    ElseIf str间隔单位 = "天" Then
        If int频率间隔 = 1 Then
            '8:00-12:00-14:00；8:00-12-14:00
            If Not StringMask(strTime, "0123456789:-") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If UBound(arrTime) + 1 <> int频率次数 Then Exit Function
            
            For i = 0 To UBound(arrTime)
                strTmp = arrTime(i)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Then Exit Function
                If strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                strPreTime = Format(strTmp, "HH:mm")
            Next
        Else
            '1/8:00-1/15:00-2/9:00；1/8:00-1/15-2/9:00
            If Not StringMask(strTime, "0123456789:-/") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If UBound(arrTime) + 1 <> int频率次数 Then Exit Function
            
            For i = 0 To UBound(arrTime)
                If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
                '相对天数部份
                strTmp = Split(arrTime(i), "/")(0)
                If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
                intCurDay = Val(strTmp)
                If intCurDay < 1 Or intCurDay > int频率间隔 Then Exit Function
                If intPreDay <> 0 Then
                    If intCurDay < intPreDay Then Exit Function
                End If
                
                '绝对时间部分
                strTmp = Split(arrTime(i), "/")(1)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Then Exit Function
                If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                
                strPreTime = Format(strTmp, "HH:mm")
                intPreDay = intCurDay
            Next
        End If
    ElseIf str间隔单位 = "小时" Then
        '1:30-2-3:30
        If Not StringMask(strTime, "0123456789:-") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> int频率次数 Then Exit Function
        
        For i = 0 To UBound(arrTime)
            strTmp = arrTime(i)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) < 1 Or Val(Split(strTmp, ":")(0)) > int频率间隔 Or Split(strTmp, ":")(0) = "" Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Then Exit Function
            If strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            strPreTime = Format(strTmp, "HH:mm")
        Next
    End If
    
    ExeTimeValid = True
End Function

Public Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean)
'功能：在ComboBox中查找并定位
'参数：blnEvent=定位时是否触发Click事件
'说明：未能定位时,设置ListIndex=-1
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If NeedName(objCbo.List(i)) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlControl.CboSetIndex(objCbo.Hwnd, i)
            End If
            Exit Sub
        End If
    Next
    If blnEvent Then
        objCbo.ListIndex = -1
    Else
        Call zlControl.CboSetIndex(objCbo.Hwnd, -1)
    End If
End Sub

Public Function GetColFormat(vsGrid As Object) As String
'功能：获取指定表格的列格式属性串
    Dim strTmp As String, i As Long
    
    For i = 0 To vsGrid.Cols - 1
        '列宽,列可见,列对齐
        strTmp = strTmp & ";" & vsGrid.ColWidth(i) & "," & IIf(vsGrid.ColHidden(i), 0, 1) & "," & vsGrid.ColAlignment(i)
    Next
    GetColFormat = Mid(strTmp, 2)
End Function

Public Sub SetColFormat(vsGrid As Object, ByVal strFormat As String)
'功能：恢复指定表格的列格式
    Dim arrCol As Variant, i As Long
    If strFormat = "" Then Exit Sub
    
    arrCol = Split(strFormat, ";")
    For i = 0 To UBound(arrCol)
        vsGrid.ColWidth(i) = Val(Split(arrCol(i), ",")(0))
        vsGrid.ColHidden(i) = Val(Split(arrCol(i), ",")(1)) = 0
        vsGrid.ColAlignment(i) = Val(Split(arrCol(i), ",")(2))
        vsGrid.Cell(2, vsGrid.FixedRows, i, vsGrid.Rows - 1, i) = Val(Split(arrCol(i), ",")(2))
    Next
    vsGrid.Cell(2, 0, 0, vsGrid.FixedRows - 1, vsGrid.Cols - 1) = 4
End Sub

Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * vNumber)
End Function

Public Function Between(x, a, B) As Boolean
'功能：判断x是否在a和b之间
    If a < B Then
        Between = x >= a And x <= B
    Else
        Between = x >= B And x <= a
    End If
End Function

Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个时间是否在暂停的时间段中
'参数：strPause="暂停时间,开始时间;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '可能尚未启用或暂停的时候被停止
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function

Public Function DateIsPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个日期是否在暂停的时间段中
'参数：strPause="暂停时间,开始时间;...."
'说明：不按时点判断,对暂停日期按算始不算止规则判断
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Format(Split(arrPause(i), ",")(0), "yyyy-MM-dd")
        strEnd = Format(Split(arrPause(i), ",")(1), "yyyy-MM-dd")
        If strEnd = "" Then strEnd = "3000-01-01" '可能尚未启用或暂停的时候被停止
        If strEnd > strBegin Then
            If Between(Format(vDate, "yyyy-MM-dd"), strBegin, _
                Format(DateAdd("d", -1, CDate(strEnd)), "yyyy-MM-dd")) Then
                DateIsPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function TimeisLastPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个时间是否在最后一次暂停的时间内,且最后一次暂停没有启用
'说明：因为这种情况下,如果长嘱没有终止时间,某些计算会死循环
    Dim arrPause() As String
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    
    strBegin = Split(arrPause(UBound(arrPause)), ",")(0)
    strEnd = Split(arrPause(UBound(arrPause)), ",")(1)
    If strEnd = "" Then
        strEnd = "3000-01-01 00:00:00"
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeisLastPause = True: Exit Function
        End If
    End If
End Function

Public Function Calc次数分解时间(lng次数 As Long, ByVal dat开始时间 As Date, dat终止时间 As Date, strPause As String, _
    ByVal str执行时间 As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String) As String
'功能：按次数计算各次的分解执行时间,要求<=终止时间及不在暂停时间段内
'参数：dat开始时间=医嘱的开始执行时间
'      dat终止时间=医嘱的执行终止时间,没有时传入"3000-01-01"
'      strPause=医嘱的暂停时间段
'返回：1."时间1,时间2,...."(yyyy-MM-dd HH:mm:ss)
'      2.lng次数=实际能够分解的次数
'说明：1.因为终止时间的限制,因此分解出来的时间个数可能小于要分解的次数
'      2.本函数是假定在执行时间及频率性质完全正确的情况下计算。
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime() As String, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    vCurTime = dat开始时间
    arrTime = Split(str执行时间, "-")
    
    If str间隔单位 = "周" Then
        vCurTime = GetWeekBase(dat开始时间)
        Do While lng次数 > 0
            '1/8:00-3/15:00-5/9:00
            For i = 1 To int频率次数
                vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                    strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                Else
                    strTmp = Split(arrTime(i - 1), "/")(1)
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                If vTmpTime > dat终止时间 Then
                    Exit Do
                ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                    Exit Do
                ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    lng次数 = lng次数 - 1
                    If lng次数 = 0 Then Exit Do
                End If
            Next
            vCurTime = vCurTime + 7
        Loop
    ElseIf str间隔单位 = "天" Then
        Do While lng次数 > 0
            If int频率间隔 = 1 Then
                '8:00-12:00-14:00；8-12-14
                For i = 1 To int频率次数
                    If InStr(arrTime(i - 1), ":") = 0 Then
                        strTmp = arrTime(i - 1) & ":00"
                    Else
                        strTmp = arrTime(i - 1)
                    End If
                    vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    
                    If vTmpTime > dat终止时间 Then
                        Exit Do
                    ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                        Exit Do
                    ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        lng次数 = lng次数 - 1
                        If lng次数 = 0 Then Exit Do
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To int频率次数
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime > dat终止时间 Then
                        Exit Do
                    ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                        Exit Do
                    ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        lng次数 = lng次数 - 1
                        If lng次数 = 0 Then Exit Do
                    End If
                Next
            End If
            vCurTime = vCurTime + int频率间隔
        Loop
    ElseIf str间隔单位 = "小时" Then
        '10:00-20:00-40:00；10-20-40；02:30
        Do While lng次数 > 0
            For i = 1 To int频率次数
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                If vTmpTime > dat终止时间 Then
                    Exit Do
                ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                    Exit Do
                ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    lng次数 = lng次数 - 1
                    If lng次数 = 0 Then Exit Do
                End If
            Next
            vCurTime = vCurTime + int频率间隔 / 24
        Loop
    End If
    lng次数 = UBound(Split(Mid(strDetailTime, 2), ",")) + 1
    Calc次数分解时间 = Mid(strDetailTime, 2)
End Function

Public Function Calc段内分解时间(ByVal datBegin As Date, ByVal datEnd As Date, ByVal strPause As String, _
    ByVal str执行时间 As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String) As String
'功能：按时间段计算各次的分解执行时间及次数
'参数：datBegin-datEnd=要计算的时间段,其中datBegin应为每个周期的开始基准时间
'      strPause=暂停的时间段
'返回："时间1,时间2,...."(yyyy-MM-dd HH:mm:ss),时间个数即为次数
'说明：1.时间段内要排除暂停的时间段,次数可能因此而减少
'      2.本函数是假定在执行时间及频率性质完全正确的情况下计算。
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime() As String, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    vCurTime = datBegin
    arrTime = Split(str执行时间, "-")
    
    If str间隔单位 = "周" Then
        vCurTime = GetWeekBase(datBegin)
        Do While vCurTime <= datEnd
            '1/8:00-3/15:00-5/9:00
            For i = 1 To int频率次数
                vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                    strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                Else
                    strTmp = Split(arrTime(i - 1), "/")(1)
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + 7, "yyyy-MM-dd") '！！！
        Loop
    ElseIf str间隔单位 = "天" Then
        Do While vCurTime <= datEnd
            If int频率间隔 = 1 Then
                '8:00-12:00-14:00；8-12-14
                For i = 1 To int频率次数
                    If InStr(arrTime(i - 1), ":") = 0 Then
                        strTmp = arrTime(i - 1) & ":00"
                    Else
                        strTmp = arrTime(i - 1)
                    End If
                    vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To int频率次数
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                Next
            End If
            vCurTime = Format(vCurTime + int频率间隔, "yyyy-MM-dd") '！！！
        Loop
    ElseIf str间隔单位 = "小时" Then
        '10:00-20:00-40:00；10-20-40；02:30
        Do While vCurTime <= datEnd
            For i = 1 To int频率次数
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = vCurTime + int频率间隔 / 24
        Loop
    End If
    Calc段内分解时间 = Mid(strDetailTime, 2)
End Function

Public Function Calc缺省药品总量(ByVal dbl单量 As Double, ByVal int疗程 As Integer, _
    ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, Optional ByVal str执行时间 As String, _
    Optional ByVal dbl剂量系数 As Double, Optional ByVal dbl包装系数 As Double, Optional ByVal int分零 As Integer) As Double
'功能：按疗程及分零特性计算药品临嘱的缺省总量(或配方缺省付数)
'参数：dbl单量=按剂量单位的一次用量
'      int疗程=一个疗程的天数
'      int分零=0-可分零,1-不分零,2-一次性(即时失效),-N-N天内分零使用有效
'      dbl包装系数=门诊包装或住院包装
'返回：按住院单位计算的药品总量
'说明：
'     1.药品分零特性是针对门诊或住院包装而言的。
'     2.dbl剂量系数,dbl包装系数,int分零=中药不传递,只计算付数
    Dim dbl天次 As Double, dbl总量 As Double
    Dim dbl剩余 As Double, dblOne As Double
    Dim intStep As Integer, dblEnd As Double
    Dim arrTime() As String, strBegin As String
    Dim strTime As String, i As Integer, j As Integer
    
    '疗程不足一个频率周期时就不管疗程
    If str间隔单位 = "周" Then
        If int疗程 < 7 Then int疗程 = 1
    ElseIf str间隔单位 = "天" Then
        If int疗程 < int频率间隔 Then int疗程 = 1
    ElseIf str间隔单位 = "小时" Then
        If int疗程 < int频率间隔 / 24 Then int疗程 = 1
    End If
    
    '一个频率周期的次数(按天)
    If str间隔单位 = "周" Then
        dbl天次 = int频率次数 / 7
    ElseIf str间隔单位 = "天" Then
        dbl天次 = int频率次数 / int频率间隔
    ElseIf str间隔单位 = "小时" Then
        dbl天次 = (int频率次数 / int频率间隔) * 24
    End If
    
    If dbl剂量系数 = 0 And dbl包装系数 = 0 Then
        '中药总量(付数) = 单付*疗程*(频率次数/频率间隔)
        dbl总量 = IntEx(int疗程 * dbl天次)
    Else
        '药品临嘱总量 = 门诊/住院包装(单量*疗程*(频率次数/频率间隔))
        If int分零 = 0 Then
            '可分零
            dbl总量 = dbl单量 * int疗程 * dbl天次 / dbl剂量系数 / dbl包装系数
        ElseIf int分零 = 1 Then
            '不分零
            dbl总量 = IntEx(dbl单量 * int疗程 * dbl天次 / dbl剂量系数 / dbl包装系数)
        ElseIf int分零 = 2 Then
            '一次性(即时失效)
            dbl总量 = IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * IntEx(int疗程 * dbl天次)
        ElseIf int分零 < 0 Then
            'ABS(int分零)天内分零使用有效(但不分零计算)
            If str执行时间 <> "" Then
                '一次门诊/住院包装的剂量
                dblOne = IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * (dbl剂量系数 * dbl包装系数)
                '缺省执行的次数和时间分解
                strTime = Calc次数分解时间(IntEx(int疗程 * dbl天次), Date, CDate("3000-01-01"), "", str执行时间, int频率次数, int频率间隔, str间隔单位)
                If strTime <> "" Then
                    arrTime = Split(strTime, ",")
                    dbl剩余 = dblOne: dbl总量 = 1
                    strBegin = arrTime(0)
                    
                    '计算总量
                    For i = 0 To UBound(arrTime)
                        If dbl剩余 < dbl单量 Or CDate(arrTime(i)) - CDate(strBegin) > Abs(int分零) Then
                            If CDate(arrTime(i)) - CDate(strBegin) > Abs(int分零) Then
                                dbl剩余 = dblOne
                            Else
                                dbl剩余 = dbl剩余 + dblOne
                            End If
                            dbl总量 = dbl总量 + 1
                            strBegin = arrTime(i)
                        End If
                        dbl剩余 = dbl剩余 - dbl单量
                    Next
                End If
            End If
        End If
    End If
    Calc缺省药品总量 = dbl总量
End Function

Public Function Calc发送药品总量(ByVal dat开始执行时间 As Date, lng次数 As Long, str分解时间 As String, _
    ByVal dbl单量 As Double, ByVal dbl剂量系数 As Double, ByVal dbl包装系数 As Double, _
    ByVal int分零 As Integer, ByVal dat终止时间 As Date, ByVal strPause As String, ByVal str执行时间 As String, _
    ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String) As Double
'功能：按发送次数及分零特性计算成药总量
'参数：dat开始执行时间=医嘱的开始执行时间,用于计算下一执行周期开始基准时间
'      lng次数=本次计划要发送的次数
'      dbl单量=按剂量单位的一次用量
'      int分零=0-可分零,1-不分零,2-一次性(即时失效),-N-N天内分零使用有效(按24小时计算)
'      dbl包装系数=门诊包装或住院包装
'下列参数用于不分零药品计算(包括-N型)：
'      str分解时间=本次发送计划执行的分解时间,与次数对应
'      strPause=医嘱的暂停时间段
'      dat终止时间=医嘱的执行终止时间,没有时传入"3000-01-01"
'返回：1.按门诊/住院单位计算的药品总量
'      2.lng次数=不分零药品(包括-N型分零药品)计算后的实际执行次数(增加)
'      3.str分解时间=不分零药品(包括-N型分零药品)计算后的分解时间(增加)
'说明：药品分零特性是针对门诊或住院包装而言的。
    Dim dbl总量 As Double, dbl剩余 As Double
    Dim arrTime() As String, dblOne As Double
    Dim strBegin As String, datBase As Date
    Dim strTmp As String, i As Long
    
    If int分零 = 0 Then
        '可分零
        dbl总量 = dbl单量 * lng次数 / dbl剂量系数 / dbl包装系数
    ElseIf int分零 = 1 Then
        '不分零
        dbl总量 = IntEx(dbl单量 * lng次数 / dbl剂量系数 / dbl包装系数)
        '按不分零计算时,多余的尽可能使用,从而使发送次数增加
        dbl剩余 = dbl总量 * dbl包装系数 * dbl剂量系数 - dbl单量 * lng次数
        If dbl剩余 >= dbl单量 And dbl单量 <> 0 Then
            '剩余理论可以执行的次数
            i = Int(dbl剩余 / dbl单量)
            '剩余实际可以执行的次数及时间分解(受终止时间限制)
            arrTime = Split(str分解时间, ",")
            datBase = Calc本周期开始时间(dat开始执行时间, CDate(arrTime(UBound(arrTime))), int频率间隔, str间隔单位)
            
            '在往后扩展时间时,最后一个周期内已执行的时间不再计算,按暂停处理
            strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
            If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
            
            strTmp = Calc次数分解时间(i, datBase, dat终止时间, strPause, str执行时间, int频率次数, int频率间隔, str间隔单位)
            If strTmp <> "" Then
                lng次数 = lng次数 + i
                str分解时间 = str分解时间 & "," & strTmp
            End If
        End If
    ElseIf int分零 = 2 Then
        '一次性(即时失效)
        dbl总量 = IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * lng次数
    ElseIf int分零 < 0 Then
        'ABS(int分零)天内分零使用有效(但不分零计算)
        arrTime = Split(str分解时间, ",")
        strBegin = arrTime(0)
        
        '一次门诊/住院包装的剂量(剂量单位)
        dblOne = IntEx(dbl单量 / dbl剂量系数 / dbl包装系数) * (dbl剂量系数 * dbl包装系数)
        '一次门诊/住院包装的剂量(包装单位)
        dbl总量 = IntEx(dbl单量 / dbl剂量系数 / dbl包装系数)
        
        '计算总量
        dbl剩余 = dblOne
        For i = 0 To UBound(arrTime)
            '第一次循环肯定够,所以不进入条件
            If dbl剩余 < dbl单量 Or CDate(arrTime(i)) - CDate(strBegin) > Abs(int分零) Then
                If CDate(arrTime(i)) - CDate(strBegin) > Abs(int分零) Then
                    dbl剩余 = dblOne
                    dbl总量 = dbl总量 + IntEx(dbl单量 / dbl剂量系数 / dbl包装系数)
                Else
                    If dbl剩余 + dbl剂量系数 * dbl包装系数 >= dbl单量 Then
                        '只需剩余加一个包装单位即够
                        dbl剩余 = dbl剩余 + dbl剂量系数 * dbl包装系数
                        dbl总量 = dbl总量 + 1
                    Else
                        '需要剩余加一次包装单位才够
                        dbl剩余 = dbl剩余 + dblOne
                        dbl总量 = dbl总量 + IntEx(dbl单量 / dbl剂量系数 / dbl包装系数)
                    End If
                End If
                strBegin = arrTime(i)
            End If
            dbl剩余 = dbl剩余 - dbl单量
        Next
        
        '剩余部分继续在有效期内按不分零计算,从而使发送次数增加
        If dbl剩余 >= dbl单量 And dbl单量 <> 0 Then
            '剩余理论可以执行的次数
            i = Int(dbl剩余 / dbl单量)
            '剩余实际可以执行的次数及时间分解(受终止时间限制)
            datBase = Calc本周期开始时间(dat开始执行时间, CDate(arrTime(UBound(arrTime))), int频率间隔, str间隔单位)
            
            '在往后扩展时间时,最后一个周期内已执行的时间不再计算,按暂停处理
            strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
            If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
            
            strTmp = Calc次数分解时间(i, datBase, dat终止时间, strPause, str执行时间, int频率次数, int频率间隔, str间隔单位)
            If strTmp <> "" Then
                arrTime = Split(strTmp, ",")
                For i = 0 To UBound(arrTime)
                    If dbl剩余 < dbl单量 Or CDate(arrTime(i)) - CDate(strBegin) > Abs(int分零) Then
                        Exit For
                    End If
                    lng次数 = lng次数 + 1
                    str分解时间 = str分解时间 & "," & arrTime(i)
                    dbl剩余 = dbl剩余 - dbl单量
                Next
            End If
        End If
    End If
    Calc发送药品总量 = dbl总量
End Function

Public Function Calc本周期开始时间(ByVal dat开始执行时间 As Date, ByVal dat某次执行时间 As Date, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String) As Date
'功能：根据长嘱的某次执行时间，得到它在该周期内的开始基准时间
    Dim datBegin As Date, datCurr As Date
    
    datCurr = dat开始执行时间
    datBegin = datCurr
    If str间隔单位 = "周" Then datCurr = GetWeekBase(datCurr)
    
    Do While datCurr <= dat某次执行时间
        datBegin = datCurr
        If str间隔单位 = "周" Then
            datCurr = datCurr + 7
        ElseIf str间隔单位 = "天" Then
            datCurr = datCurr + int频率间隔
        ElseIf str间隔单位 = "小时" Then
            datCurr = DateAdd("h", int频率间隔, datCurr)
        End If
    Loop
    Calc本周期开始时间 = datBegin
End Function

Public Function Trim分解时间(ByVal lng次数 As Long, ByVal str分解时间 As String) As String
'功能：将医嘱执行的分解时间按次数进行截断
    Dim arrTime() As String, strTmp As String, i As Long
    
    arrTime = Split(str分解时间, ",")
    For i = 0 To lng次数 - 1
        strTmp = strTmp & "," & arrTime(i)
    Next
    Trim分解时间 = Mid(strTmp, 2)
End Function

Public Function Calc持续性长嘱次数(ByVal datBegin As Date, ByVal datEnd As Date, _
    ByVal str上次执行时间 As String, ByVal str执行终止时间 As String, _
    ByVal strPause As String, Optional str首次时间 As String, _
    Optional str末次时间 As String, Optional str分解时间 As String) As Long
'功能：对持续性非药长嘱计算它本次应该发送的次数,及首末时间
'参数：str上次执行时间=不一定等于本次发送的开始时间
'      str执行终止时间=终止这天不发送
'返回：本次该医嘱发送的次数
'      str首次时间,str末次时间=返回yyyy-MM-dd HH:mm:ss
'说明：持续性长嘱按日期每天发送一次处理,处理规则与床位费类似(暂停时间算始不算止;终止当天不发送)
    Dim curDate As Date, lng次数 As Long, blnSend As Boolean
    
    str首次时间 = "": str末次时间 = "": str分解时间 = ""
    curDate = CDate(Format(datBegin, "yyyy-MM-dd"))
    Do While curDate <= CDate(Format(datEnd, "yyyy-MM-dd"))
        If Not DateIsPause(curDate, strPause) Then
            blnSend = True
            If str上次执行时间 <> "" Then
                If Format(curDate, "yyyy-MM-dd") <= Format(str上次执行时间, "yyyy-MM-dd") Then
                    blnSend = False '应大于上次执行时间才执行
                End If
            End If
            If str执行终止时间 <> "" Then
                If Format(curDate, "yyyy-MM-dd") >= Format(str执行终止时间, "yyyy-MM-dd") Then
                    blnSend = False '应小于执行终止时间才执行
                End If
            End If
            If blnSend Then
                lng次数 = lng次数 + 1
                If str首次时间 = "" Then
                    str首次时间 = Format(curDate, "yyyy-MM-dd 00:00:00") '定为零点执行
                End If
                str末次时间 = Format(curDate, "yyyy-MM-dd 00:00:00")
                str分解时间 = str分解时间 & "," & str末次时间
            End If
        End If
        curDate = curDate + 1
    Loop
    str分解时间 = Mid(str分解时间, 2)
    Calc持续性长嘱次数 = lng次数
End Function

Public Function CheckScope(varL As Double, varR As Double, varI As Double) As String
'功能：判断输入金额是否在原价和现从限定的范围内
'参数：varL=原价,varR=现价,varI=输入金额
'返回：如果不在范围内,则为提示信息,否则为空串
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '如果数值符号相同,则用绝对值判断
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "输入的价格绝对值不在范围(" & FormatEx(Abs(varL), 5) & "-" & FormatEx(Abs(varR), 5) & ")内."
        End If
    Else
        '如果符号不相同,则用原始范围判断
        If varI < varL Or varI > varR Then
            CheckScope = "输入的价格值不在范围(" & FormatEx(varL, 5) & "-" & FormatEx(varR, 5) & ")内."
        End If
    End If
End Function

Public Sub GetCboIndex(objCbo As Object, strFind As String, Optional Keep As Boolean)
'功能：由字符串在ComboBox中查找索引
'参数：Keep=如果未匹配，是否保持原索引
    Dim i As Integer
    
    '先精确查找
    For i = 0 To objCbo.ListCount - 1
        If objCbo.List(i) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        ElseIf NeedName(objCbo.List(i)) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        End If
    Next
    
    '最后模糊查找
    If strFind <> "" Then
        For i = 0 To objCbo.ListCount - 1
            If InStr(objCbo.List(i), strFind) > 0 Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not Keep Then objCbo.ListIndex = -1
End Sub

Public Sub FindCboIndex(objCbo As Object, lngData As Long, Optional Keep As Boolean)
'功能：由项目值查找ComboBox的项目索引
'参数：Keep=如果未匹配，是否保持原索引
    Dim i As Integer
    
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not Keep Then objCbo.ListIndex = -1
End Sub

Public Function SeekCboIndex(objCbo As Object, lngData As Long) As Long
'功能：由ItemData查找ComboBox的索引值
    Dim i As Integer
    
    SeekCboIndex = -1
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                SeekCboIndex = i: Exit Function
            End If
        Next
    End If
End Function

Public Function GetNextControl(ByVal intTab As Integer, ByVal frmForm As Object) As Object
'功能：获取下一个光标顺序的控件
    Dim objNext As Object, i As Long
    
    '先找比当前控件TabIndex大的
    For i = 0 To frmForm.Controls.Count - 1
        If InStr("TextBox,ComboBox,VSFlexGrid", TypeName(frmForm.Controls(i))) > 0 Then
            If frmForm.Controls(i).Enabled And frmForm.Controls(i).Visible And frmForm.Controls(i).TabStop Then
                If frmForm.Controls(i).TabIndex > intTab Then
                    If objNext Is Nothing Then
                        Set objNext = frmForm.Controls(i)
                    ElseIf frmForm.Controls(i).TabIndex < objNext.TabIndex Then
                        Set objNext = frmForm.Controls(i)
                    End If
                End If
            End If
        End If
    Next
    If objNext Is Nothing Then
        '没有则找比当前控件TabIndex小的
        For i = 0 To frmForm.Controls.Count - 1
            If InStr("TextBox,ComboBox,VSFlexGrid", TypeName(frmForm.Controls(i))) > 0 Then
                If frmForm.Controls(i).Enabled And frmForm.Controls(i).Visible And frmForm.Controls(i).TabStop Then
                    If frmForm.Controls(i).TabIndex < intTab Then
                        If objNext Is Nothing Then
                            Set objNext = frmForm.Controls(i)
                        ElseIf frmForm.Controls(i).TabIndex < objNext.TabIndex Then
                            Set objNext = frmForm.Controls(i)
                        End If
                    End If
                End If
            End If
        Next
    End If
    Set GetNextControl = objNext
End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function Custom_WndMessage(ByVal Hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.x = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.x = 1600
        MinMax.ptMaxTrackSize.y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, Hwnd, Msg, wp, lp)
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckAdviceWindow(ByVal strTitle As String) As Boolean
'功能：检查医嘱编辑窗口是否已经打开
    Dim lngHwnd As Long
    
    '其它窗口打开了
    lngHwnd = FindWindow("ThunderFormDC", strTitle)
    If lngHwnd <> 0 Then
        MsgBox "医嘱编辑窗口已经打开，请先完成当前操作后再执行。", vbInformation, gstrSysName
        Call ShowWindow(lngHwnd, SW_RESTORE)
        Call BringWindowToTop(lngHwnd)
        Exit Function
    End If
    CheckAdviceWindow = True
End Function

Public Function GetWeekBase(ByVal datDate As Date) As Date
'功能：获取指定时间所在星期的星期一的时间
    GetWeekBase = Format(datDate - (Weekday(datDate, vbMonday) - 1), "yyyy-MM-dd 00:00:00")
End Function
Public Function Wndproc(ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'接收windows消息
    Dim hw As Long             '窗口"ZLPACS Viewer"句柄
    Dim lngX As Long           '主窗体的宽
    Dim LngY As Long           '主窗体的高
    Dim objPacsCore As Object
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim GetCurr As POINTAPI
    Dim SetCurr As POINTAPI
    Dim FrmReprot As Form
    On Error Resume Next
    If Msg = WM_HOTKEY Then
        If wParam = idHotKey Then
            Dim lp As taLong, i2 As t2Int
            lp.ll = lParam
            LSet i2 = lp
            If (i2.lWord = Modifiers) And i2.hWord = uVirtKey Then
                hw = FindWindow(vbNullString, "ZLPACS Viewer")
                If hw <> 0 Then
                    lngX = Screen.Width / Screen.TwipsPerPixelX
                    LngY = Screen.Height / Screen.TwipsPerPixelY
                    GetCursorPos GetCurr
                    If GetCurr.x < 0 Or GetCurr.x > lngX Then
                        SetCurr.x = (Screen.Width / Screen.TwipsPerPixelX) / 2
                        SetCurr.y = (Screen.Height / Screen.TwipsPerPixelY) / 2
                        SetCursorPos SetCurr.x, SetCurr.y
                        Set FrmReprot = frmPACStation.GetReprotFrm()
                        If Not FrmReprot Is Nothing Then
                            FrmReprot.SetFocus
                        Else
                            frmPACStation.SetFocus
                        End If
                    Else
                        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
                        lngLeft = objPacsCore.GetLeft / Screen.TwipsPerPixelX
                        lngTop = objPacsCore.GetTop / Screen.TwipsPerPixelY
                        lngWidth = objPacsCore.Getwidth / Screen.TwipsPerPixelX
                        lngHeight = objPacsCore.GetHeight / Screen.TwipsPerPixelY
                        Debug.Print lngLeft & " " & lngWidth
                        If lngLeft < 0 Or lngLeft + lngWidth > lngX Then
                            SetCurr.x = (lngLeft + lngWidth) - (lngWidth / 2)
                            SetCurr.y = (lngHeight + lngTop) - (lngHeight / 2)
                            SetCursorPos SetCurr.x, SetCurr.y
                            objPacsCore.SetViewerFocus
                        End If
                    End If
                End If
            End If
        End If
    End If
    '如果不是热键信息则调用原来的程序
    Wndproc = CallWindowProc(preWinProc, Hwnd, Msg, wParam, lParam)
End Function

