Attribute VB_Name = "mdlPublic"
Option Explicit



Public lngTXTProc As Long '保存默认的消息函数的地址
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long型最大值


Public Modifiers As Long, uVirtKey As Long, idHotKey As Long

Public Enum Enum_Inside_Program
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p辅诊记录管理 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
End Enum
Public gcolPrivs As Collection              '记录内部模块的权限
'DICOM图象参数
Public Const ATTR_检查日期 As String = "8:20"
Public Const ATTR_检查时间 As String = "8:30"
Public Const ATTR_影像类别 As String = "8:60"
Public Const ATTR_检查设备 As String = "8:1090"

'报告内容分隔符
Public Const SPLITER_REPORT = "[[@]]"
Public Const SPLITER_ELEMENT = "[[;]]"
'报告窗体
Public Const Report_Form_frmReportES  As String = "内镜基本信息"
Public Const Report_Form_frmReportPathology As String = "病理妇科液基薄层信息"
Public Const Report_Form_frmReportUS As String = "B超心脏测量信息"
Public Const Report_Form_frmReportCustom As String = "自定义专科报告"

Public gobjLogFile As clsLogFile
Public gblnUseDebugLog As Boolean


'获取虚拟键对应的名称
Public Function GetKeyAliasEx(ByVal lngVirtualKey As Long) As String
    GetKeyAliasEx = ""
    
    If lngVirtualKey >= 59 And lngVirtualKey <= 68 Then
        GetKeyAliasEx = "F" & (lngVirtualKey - 58)
    End If
    
    If lngVirtualKey >= 87 And lngVirtualKey <= 88 Then
        GetKeyAliasEx = "F" & (lngVirtualKey - 76)
    End If
End Function

'取得组合键别名
Public Function GetKeyAlias(ByVal KeyCode As Integer, ByVal Shift As Integer) As String

    Dim strShift As String
    Dim strTemp As String
    
    
    strShift = IIf((Shift And vbCtrlMask) <> 0, "CTRL", "")
    
    strTemp = IIf((Shift And vbShiftMask) <> 0, "SHIFT", "")
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
    strTemp = IIf((Shift And vbAltMask) <> 0, "ALT", "")
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
     
    
             
    strTemp = ""
    If KeyCode >= 48 And KeyCode <= 90 Then
        strTemp = Chr(KeyCode)
        
        If strShift = "" Then strShift = "MENU"
    End If
    
    If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
        strTemp = "F" & (KeyCode - 111)
    End If
    
    Select Case KeyCode
        Case vbKeySpace
            strTemp = "SPACE"
    End Select
    
    
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
    GetKeyAlias = strShift
                
End Function

Public Sub WriteLog(ByVal strLog As String)
'写入日志
    '如果未启用调试日志，则直接退出
    If Not gblnUseDebugLog Then Exit Sub
    
    '初始化日志对象
    If gobjLogFile Is Nothing Then
        Set gobjLogFile = New clsLogFile
    End If
    
    If gobjLogFile.OpenLog() Then
        Call gobjLogFile.WriteLog(strLog)
        Call gobjLogFile.CloseLog
    End If
End Sub



Private Sub DoEventEx(ByVal lngTime As Long)
    Dim curMsg As MSG
    Dim curTime As Double

    curTime = GetTickCount
    While True
        If GetMessage(curMsg, 0, 0, 0) Then
            TranslateMessage curMsg
            DispatchMessage curMsg
        End If
        
        If GetTickCount - curTime > lngTime Then Exit Sub
    Wend
End Sub

Public Sub OutputDebug(ByVal strMethob As String, objErr As ErrObject)
    OutputDebugString "[" & App.ProductName & "]" & strMethob & "：" & objErr.Description
End Sub


Public Sub RaiseErr(objErr As ErrObject)
    Call err.Raise(objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext)
End Sub



Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
    Dim vRect As RECT, vPos As PointAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
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
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
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
    Dim vRect As RECT, vDot1 As PointAPI, vDot2 As PointAPI
    
    Call GetWindowRect(ToolBar.hWnd, vRect)
    vDot1.X = vRect.Left: vDot1.Y = vRect.Top
    vDot2.X = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hWnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hWnd, vDot2)
    
    vDot1.X = vDot1.X * 15: vDot1.Y = vDot1.Y * 15
    vDot2.X = vDot2.X * 15: vDot2.Y = vDot2.Y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.X + Button.Left, vDot2.Y
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
Public Function GetColNum(listTemp As Object, strHead As String) As Integer
    Dim i As Integer
    Select Case UCase(TypeName(listTemp))
        Case UCase("ReportControl")
            For i = 0 To listTemp.Columns.Count - 1
                If listTemp.Columns.Column(i).Caption = strHead Then GetColNum = listTemp.Columns.Column(i).ItemIndex: Exit Function
            Next
        Case UCase("ListView")
            For i = 1 To listTemp.ColumnHeaders.Count
                If listTemp.ColumnHeaders(i).Text = strHead Then GetColNum = i: Exit Function
            Next
        Case UCase("MSHFlexGrid") '以下类型待增，尚未用到
        Case UCase("BillEdit")
        Case UCase("VSFlexGrid")
            For i = 0 To listTemp.Cols - 1
                If listTemp.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
            Next
        Case UCase("BillEdit")
        Case UCase("DataGrid")
    End Select
End Function
Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal LngY As Long) As PointAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As PointAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = LngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function
'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal MSG As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If MSG <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hWnd, MSG, wp, lp)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
'功能：将文本按Varchar2的长度计算方法进行截断
    Dim strText As String
    
    strText = IIf(IsNull(varText), "", varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    '去掉可能出现的半个字符
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function
Public Function To_Date(ByVal dat日期 As Date) As String
'功能:将入参中的日期传换成ORACLE需要的日期格式串
    To_Date = "To_Date('" & Format(dat日期, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function
Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
'参数：blnForceNum=当为Null时，是否强制表示为数字型
    ZVal = IIf(Val(varValue) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varValue))
End Function

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
Public Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean, Optional blnPreserve As Boolean = False, Optional blnIsSearchNo As Boolean = False)
'功能：在ComboBox中查找并定位
'参数：blnEvent=定位时是否触发Click事件
      'blnPreserve--如果找不到匹配项目，则保持原有项目
      'blnIsSearchNo --是否是通过编码定位
'说明：未能定位时,设置ListIndex=-1
    Dim i As Long

    For i = 0 To objCbo.ListCount - 1
        If IIf(blnIsSearchNo, NeedNo(objCbo.list(i)), NeedName(objCbo.list(i))) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlControl.CboSetIndex(objCbo.hWnd, i)
            End If
            Exit Sub
        End If
    Next
    
    If blnPreserve = True Then
        If blnEvent = False Then
            Call zlControl.CboSetIndex(objCbo.hWnd, objCbo.ListIndex)
        End If
    Else
        If blnEvent Then
            objCbo.ListIndex = -1
        Else
            Call zlControl.CboSetIndex(objCbo.hWnd, -1)
        End If
    End If
    
End Sub
Public Sub SeekIndexWithNo(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean)
'功能：在ComboBox中查找并定位
'参数：blnEvent=定位时是否触发Click事件
'说明：未能定位时,设置ListIndex=-1
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If NeedNo(objCbo.list(i)) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlControl.CboSetIndex(objCbo.hWnd, i)
            End If
            Exit Sub
        End If
    Next
    If blnEvent Then
        objCbo.ListIndex = -1
    Else
        Call zlControl.CboSetIndex(objCbo.hWnd, -1)
    End If
End Sub
Public Function NeedNo(strList As String) As String
    If InStr(strList, "[") > 0 And InStr(strList, "-") = 0 Then
        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "[") - 1))
    ElseIf InStr(strList, "(") > 0 And InStr(strList, "-") = 0 Then
        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "(") - 1))
    ElseIf InStr(strList, "-") > 0 Then
        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "-") - 1))
    Else
        NeedNo = LTrim(strList)
    End If
End Function
Public Function Get年龄(str出生日期 As String) As Integer
'功能:根据出生日期取得年龄
    If IsDate(str出生日期) Then
        Get年龄 = DateDiff("yyyy", CDate(str出生日期), Format(zlDatabase.Currentdate, "YYYY-MM-DD"))
    End If
End Function
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

Public Function Between(X, a, b) As Boolean
'功能：判断x是否在a和b之间
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function
Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '主要有空格引起的
        strTmp = Substr(strCode, 1, lngLen)
    End If
    '取掉最后半个字符
    Rpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取指定字串的值,字串中可以包含汉字
    '--入参数:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '--出参数:
    '--返  回:子串
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    err = 0
    On Error GoTo errHand:

    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
    Exit Function
errHand:
    Substr = ""
End Function

Public Sub GetCboIndex(objCbo As Object, strFind As String, Optional Keep As Boolean)
'功能：由字符串在ComboBox中查找索引
'参数：Keep=如果未匹配，是否保持原索引
    Dim i As Integer
    
    '先精确查找
    For i = 0 To objCbo.ListCount - 1
        If objCbo.list(i) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        ElseIf NeedName(objCbo.list(i)) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        End If
    Next
    
    '最后模糊查找
    If strFind <> "" Then
        For i = 0 To objCbo.ListCount - 1
            If InStr(objCbo.list(i), strFind) > 0 Then
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
    If err.Number <> 0 Then err.Clear: InDesign = True
End Function
Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
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

Public Function HIWORD(LongIn As Long) As Integer
    ' 取出32位值的高16位
    HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function LOWORD(LongIn As Long) As Integer
    ' 取出32位值的低16位
    If (LongIn And &HFFFF&) > &H7FFF Then
        LOWORD = (LongIn And &HFFFF&) - &H10000
    Else
        LOWORD = LongIn And &HFFFF&
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
''为了处理双屏时对话框的正确显示位置，用API函数改写了一下MsgBox函数
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgBoxD(ByRef frmParent As Form, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = MB_OK, Optional Title As String = "") As Long

    MsgBoxD = MessageBox(frmParent.hWnd, Prompt, Title, Buttons)

End Function
