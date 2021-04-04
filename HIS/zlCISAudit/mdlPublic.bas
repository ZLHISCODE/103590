Attribute VB_Name = "mdlPublic"
Option Explicit

'######################################################################################################################
'常量定义



'----------------------------------------------------------------------------------------------------------------------
'枚举
Public Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Public Enum COLOR_NativeXpPlain
    BackgroundDark = 14054755
    BackgroundLight = 15180411
    HighlightBorderBottomRight = 8388608
    HighlightBorderTopLeft = 8388608
    HighlightHot = 12775167
    HighlightPressed = 4096254
    HighlightSelected = 7323903
    NormalGroupCaptionDark = 14215660
    NormalGroupCaptionLight = 14215660
    NormalGroupCaptionTextHot = 0
    NormalGroupCaptionTextNormal = 0
    NormalGroupClient = 16244694
    NormalGroupClientBorder = 16777215
    NormalGroupClientLink = 12999969
    NormalGroupClientLinkHot = 16748098
    NormalGroupClientText = 0
    SpecialGroupCaptionDark = 14215660
    SpecialGroupCaptionLight = 14215660
    SpecialGroupCaptionTextHot = 0
    SpecialGroupCaptionTextSpecial = 0
    SpecialGroupClient = 16244694
    SpecialGroupClientBorder = 16777215
    SpecialGroupClientLink = 12999969
    SpecialGroupClientLinkHot = 16748098
    SpecialGroupClientText = 0
End Enum

Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21

Public Type PointAPI
     x As Long
     Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const ETO_CLIPPED = 4
Public Const ETO_GRAYED = 1
Public Const ETO_OPAQUE = 2
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

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Sub InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long)
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

'######################################################################################################################
'过程清单

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


Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    '******************************************************************************************************************
    '功能：去掉TextBox的默认右键菜单
    '参数：
    '返回：
    '说明：如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    '******************************************************************************************************************
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal x As Single, ByVal Y As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngX As Long
    Dim lngY As Long
    Dim lngLoop As Long
    Dim lngXY As Long
            
    lngX = x / 15
    lngY = Y / 15
        
    lngXY = 2
    For lngLoop = 1 To 15
        lngXY = lngXY * 2
    Next
    
    lngXY = lngXY * lngY + lngX
    
    SendMessage lngHwnd, WM_LBUTTONDOWN, 0, ByVal lngXY
    SendMessage lngHwnd, WM_LBUTTONUP, 0, ByVal lngXY

End Sub


Public Function AppendCode(ByVal strName As String, ByVal strCode As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If strName <> "" And strCode <> "" Then
        AppendCode = "【" & strCode & "】" & strName
    Else
        AppendCode = strName
    End If
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

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function

Public Function GetVBKey() As Long
    
    Dim strTmp As String
    
    strTmp = Timer
    strTmp = Replace(strTmp, ".", "")
    
    GetVBKey = Format(Now, "1dd") & strTmp
    
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1表示开始;2表示结束
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = ParamInfo.系统名称
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
        
            strSQL = CStr(rs("SQL").Value)
            Call zlDatabase.ExecuteProcedure(strSQL, strTitle)
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
errHand:
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
    
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim strChar As String
    
    strChar = "ZXCVBNMASDFGHJKLQWERTYUIOPzxcvbnmasdfghjklqwertyuiop"
    
    Select Case bytMode
    Case 1          '全数字
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 2          '全字母
    
        For lngLoop = 1 To Len(Text)
            If InStr(strChar, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
        
    Case 99
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
    End Select
End Function

Public Function IsExitsField(ByVal rsData As ADODB.Recordset, ByVal strFieldName As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    
    On Error Resume Next
    
    strTmp = ""
    strTmp = rsData.Fields(strFieldName).Name
    IsExitsField = (strTmp = strFieldName)
    
End Function

Public Function CopyRecordStruct(ByVal rsFrom As ADODB.Recordset, Optional ByVal blnRowID As Boolean = False, Optional ByVal blnNotOpen As Boolean = False) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim lngLoop As Long
    Dim rs As ADODB.Recordset
    
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic

    For lngLoop = 0 To rsFrom.Fields.count - 1
        
'        rs.Fields.Append rsFrom.Fields(lngLoop).Name, rsFrom.Fields(lngLoop).Type, rsFrom.Fields(lngLoop).DefinedSize + 10, rsFrom.Fields(lngLoop).Attributes
        
        Select Case rsFrom.Fields(lngLoop).Type
        Case 135            'Oracle的Date型
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, adVarChar, 30, rsFrom.Fields(lngLoop).Attributes
        Case 5
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, rsFrom.Fields(lngLoop).Type, 30, rsFrom.Fields(lngLoop).Attributes
        Case Else
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, adVarChar, rsFrom.Fields(lngLoop).DefinedSize + 12
        End Select

    Next
    If blnRowID Then
        rs.Fields.Append "行号", adVarChar, 30
    End If
    
    If blnNotOpen = False Then rs.Open
    
    Set CopyRecordStruct = rs
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function CopyRecordData(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset, Optional blnAll As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    If blnAll Then
        If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    End If
    
    Do While Not rsFrom.EOF
        rsTo.AddNew
        For lngLoop = 0 To rsFrom.Fields.count - 1
            
            On Error Resume Next
            strTmp = ""
            strTmp = rsTo.Fields(rsFrom.Fields(lngLoop).Name).Name
            On Error GoTo errHand
            
            If UCase(strTmp) = UCase(rsFrom.Fields(lngLoop).Name) Then
                rsTo.Fields(strTmp).Value = Trim(zlCommFun.NVL(rsFrom.Fields(lngLoop).Value))
            End If

        Next
        If blnAll = False Then Exit Do
        rsFrom.MoveNext
        rsTo.Update
    Loop
    
    CopyRecordData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function AppendRecord(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer
    
    Do While Not rsFrom.EOF
        rsTo.AddNew
        For intCol = 0 To rsFrom.Fields.count - 1
            rsTo(rsFrom.Fields(intCol).Name).Value = zlCommFun.NVL(rsFrom.Fields(intCol).Value)
        Next
        
        rsFrom.MoveNext
    Loop
    
    AppendRecord = True
    
End Function

Public Function DeleteRecordData(rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:删除记录集
    '参数:rs        要删除的记录集
    '返回:无
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If Not (rs Is Nothing) Then
        If rs.RecordCount > 0 Then rs.MoveFirst
        While Not rs.EOF
            rs.Delete
            rs.MoveNext
        Wend
    End If
    
    DeleteRecordData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
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
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    Rpad = zl9Comlib.zlStr.Rpad(strCode, lngLen, strChar, True)
End Function
Public Sub SetDockRight(cbsMain As Object, BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsMain.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsMain.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function



Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        objControl.ID = lngID
    End With
    
    Set NewCommandBar = objControl
    
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.STYLE = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Sub LocationObj(ByRef objTxt As Object, Optional ByVal blnDoevents As Boolean = False)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If blnDoevents Then DoEvents
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
    
errHand:
    
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '******************************************************************************************************************
    '功能:检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    '参数:
    '返回:
    '******************************************************************************************************************
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, ParamInfo.系统名称
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字符！", vbExclamation, ParamInfo.系统名称
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    MsgBox strInfo, vbInformation, ParamInfo.系统名称
    
End Sub

Public Function LoadTree(ByRef objTvw As Object, ByVal rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objNode As Node
    Dim strTmp As String
    Dim strIcon As String
    Dim strIconSel As String
    Dim blnIcon As Boolean
    Dim blnIconSel As Boolean
    
    On Error GoTo errHand

    On Error Resume Next
    
    blnIcon = (rs("图标").Name = "图标")
    blnIconSel = (rs("选中图标").Name = "选中图标")
    
    On Error GoTo errHand
    
    LockWindowUpdate objTvw.hWnd

    Do While Not rs.EOF
        strTmp = AppendCode(zlCommFun.NVL(rs("名称").Value), zlCommFun.NVL(rs("编码").Value))
        
        If blnIcon Then strIcon = zlCommFun.NVL(rs("图标").Value)
        If blnIconSel Then strIconSel = zlCommFun.NVL(rs("选中图标").Value)
        
        If IsNull(rs("上级id").Value) Then
            Set objNode = objTvw.Nodes.Add(, , "K" & zlCommFun.NVL(rs("ID").Value, 0), strTmp, strIcon, strIconSel)
        Else
            Set objNode = objTvw.Nodes.Add("K" & rs("上级id").Value, tvwChild, "K" & zlCommFun.NVL(rs("ID").Value, 0), strTmp, strIcon, strIconSel)
        End If

        rs.MoveNext
    Loop

    LockWindowUpdate 0

    LoadTree = True

    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AddComboData(objSource As Object, ByVal rsData As ADODB.Recordset, ByVal strValueFields As String, Optional ByVal strKeyField As String = "", Optional ByVal strDefaultField As String = "", Optional ByVal blnClear As Boolean = True, Optional ByVal blnSelect As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：装载数据入指定的组合下拉框或网格中的下拉框中
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim varValueField As Variant
    Dim strValue As String
    Dim intCol As Integer
    
    If blnClear = True Then objSource.Clear
    
    On Error Resume Next
    If IsExitsField(rsData, strKeyField) = False Then strKeyField = ""
    If IsExitsField(rsData, strDefaultField) = False Then strDefaultField = ""
    On Error GoTo 0
    
    varValueField = Split(strValueFields, ",")
    
    If rsData.BOF = False Then
    
        rsData.MoveFirst
        
        While Not rsData.EOF
            strValue = ""
            For intCol = 0 To UBound(varValueField)
                strValue = strValue & "-" & CStr(rsData.Fields(varValueField(intCol)).Value)
            Next
            
            If strValue <> "" Then strValue = Mid(strValue, 2)
            If strValue <> "" Then
            
                objSource.AddItem strValue
                
                If strKeyField <> "" Then objSource.ItemData(objSource.NewIndex) = Val(rsData.Fields(strKeyField).Value)
                
                If strDefaultField <> "" Then
                
                    If rsData.Fields.count > 2 Then
                        If Val(rsData.Fields(2).Value) = 1 Then objSource.ListIndex = objSource.NewIndex
                    End If
                    
                End If
            End If
            
            rsData.MoveNext
        Wend
        rsData.MoveFirst
    End If
    
    If blnSelect Then
        If objSource.ListCount > 0 And objSource.ListIndex = -1 Then objSource.ListIndex = 0
    End If
    
    AddComboData = True
End Function

Public Function AnalyseIDCard(ByVal strIDCard As String, ByRef strBirdthDay As String, ByRef strSex As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If Len(strIDCard) = 0 Then
        AnalyseIDCard = True
        Exit Function
    End If
    
    strIDCard = UCase(strIDCard)
    
    If Len(strIDCard) <> 15 And Len(strIDCard) <> 18 Then Exit Function
    If Len(strIDCard) = 18 And InStr(strIDCard, "X") <> 18 And InStr(strIDCard, "X") > 0 Then Exit Function
    
    Select Case Len(strIDCard)
    Case 15
    
        strBirdthDay = "19" & Mid(strIDCard, 7, 2) & "-" & Mid(strIDCard, 9, 2) & "-" & Mid(strIDCard, 11, 2)
        
    Case 18         '510221197309262119
    
        strBirdthDay = Mid(strIDCard, 7, 4) & "-" & Mid(strIDCard, 11, 2) & "-" & Mid(strIDCard, 13, 2)
                
    End Select
    
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '功能:获取特殊时间
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim intDay As Integer
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(zlDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(zlDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365 * 2, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function

Public Function DockPannelCreate(ByRef dkpMain As DockingPane, ByVal intIndex As Integer, _
                                    ByVal lngCX As Long, ByVal lngCY As Long, _
                                    ByVal bytDirection As DockingDirection, _
                                    Optional ByVal objNeighbour As Pane = Nothing, _
                                    Optional ByVal strTitle As String, _
                                    Optional ByVal bytOptions As PaneOptions) As Pane
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Set DockPannelCreate = dkpMain.CreatePane(intIndex, lngCX, lngCY, bytDirection, objNeighbour)
    DockPannelCreate.Title = strTitle
    DockPannelCreate.Options = PaneNoCaption
    
End Function

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Function TabControlInit(ByRef tbc As TabControl, _
                                Optional ByVal bytAppearance As XTPTabAppearanceStyle = xtpTabAppearancePropertyPage2003) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With tbc
        
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
'            .Position = bytPosition
        End With
        
        Set .Icons = frmPubResource.imgPublic.Icons
        

        
    End With

    TabControlInit = True
    
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto 'xtpSystemThemeBlue
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    CommandBarInit = True
    
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '弹出菜单处理
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Set cbrPopupItem2 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
                cbrPopupItem2.Parameter = cbrControl2.Parameter
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function FillTreeData(ByRef objTvw As Object, ByVal rs As ADODB.Recordset) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '--------------------------------------------------------------------------------------------------------
    Dim objNode As Node
    Dim strTmp As String
    Dim strIcon As String
    Dim strIconSel As String
    Dim blnIcon As Boolean
    Dim blnIconSel As Boolean
    
    On Error GoTo errHand
    
    LockWindowUpdate objTvw.hWnd

    Do While Not rs.EOF
        strTmp = AppendCode(zlCommFun.NVL(rs("名称").Value), zlCommFun.NVL(rs("编码").Value))
        
        strIcon = zlCommFun.NVL(rs("图标").Value)
        strIconSel = zlCommFun.NVL(rs("选中图标").Value)
        
        If IsNull(rs("上级id").Value) Then
            Set objNode = objTvw.Nodes.Add(, , "K" & zlCommFun.NVL(rs("ID").Value, 0), strTmp, zlCommFun.NVL(rs("图标").Value), strIconSel)
        Else
            Set objNode = objTvw.Nodes.Add("K" & rs("上级id").Value, tvwChild, "K" & zlCommFun.NVL(rs("ID").Value, 0), strTmp, zlCommFun.NVL(rs("图标").Value), strIconSel)
        End If

        rs.MoveNext
    Loop

    LockWindowUpdate 0

    FillTreeData = True

    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function IncStr(ByVal strVal As String) As String
    '******************************************************************************************************************
    '功能：对一个字符串自动加1。
    '说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
    '******************************************************************************************************************
    Dim i As Long, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function

Public Function RestoreTaskPanelPaterrn(ByVal objTpl As Object)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With objTpl
        
        .ColorSet.BackgroundDark = COLOR_NativeXpPlain.BackgroundDark
        .ColorSet.BackgroundLight = COLOR_NativeXpPlain.BackgroundLight
        .ColorSet.HighlightBorderBottomRight = COLOR_NativeXpPlain.HighlightBorderBottomRight
        .ColorSet.HighlightBorderTopLeft = COLOR_NativeXpPlain.HighlightBorderTopLeft
        .ColorSet.HighlightHot = COLOR_NativeXpPlain.HighlightHot
        .ColorSet.HighlightPressed = COLOR_NativeXpPlain.HighlightPressed
        .ColorSet.HighlightSelected = COLOR_NativeXpPlain.HighlightSelected
        
        .ColorSet.NormalGroupCaptionDark = COLOR_NativeXpPlain.NormalGroupCaptionDark
        .ColorSet.NormalGroupCaptionLight = COLOR_NativeXpPlain.NormalGroupCaptionLight
        .ColorSet.NormalGroupCaptionTextHot = COLOR_NativeXpPlain.NormalGroupCaptionTextHot
        .ColorSet.NormalGroupCaptionTextNormal = COLOR_NativeXpPlain.NormalGroupCaptionTextNormal
        .ColorSet.NormalGroupClient = COLOR_NativeXpPlain.NormalGroupClient
        .ColorSet.NormalGroupClientBorder = COLOR_NativeXpPlain.NormalGroupClientBorder
        .ColorSet.NormalGroupClientLink = COLOR_NativeXpPlain.NormalGroupClientLink
        
        .ColorSet.NormalGroupClientLinkHot = COLOR_NativeXpPlain.NormalGroupClientLinkHot
        .ColorSet.NormalGroupClientText = COLOR_NativeXpPlain.NormalGroupClientText
        .ColorSet.SpecialGroupCaptionDark = COLOR_NativeXpPlain.SpecialGroupCaptionDark
        .ColorSet.SpecialGroupCaptionLight = COLOR_NativeXpPlain.SpecialGroupCaptionLight
        .ColorSet.SpecialGroupCaptionTextHot = COLOR_NativeXpPlain.SpecialGroupCaptionTextHot
        .ColorSet.SpecialGroupCaptionTextSpecial = COLOR_NativeXpPlain.SpecialGroupCaptionTextSpecial
        .ColorSet.SpecialGroupClient = COLOR_NativeXpPlain.SpecialGroupClient
        .ColorSet.SpecialGroupClientBorder = COLOR_NativeXpPlain.SpecialGroupClientBorder
        .ColorSet.SpecialGroupClientLink = COLOR_NativeXpPlain.SpecialGroupClientLink
        .ColorSet.SpecialGroupClientLinkHot = COLOR_NativeXpPlain.SpecialGroupClientLinkHot
        .ColorSet.SpecialGroupClientText = COLOR_NativeXpPlain.SpecialGroupClientText
    End With
End Function

Public Function RestoreDockPanelPaterrn(ByVal objDkp As Object)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With objDkp
        
        .ColorSet.BackgroundDark = COLOR_NativeXpPlain.BackgroundDark
        .ColorSet.BackgroundLight = COLOR_NativeXpPlain.BackgroundLight
        .ColorSet.HighlightBorderBottomRight = COLOR_NativeXpPlain.HighlightBorderBottomRight
        .ColorSet.HighlightBorderTopLeft = COLOR_NativeXpPlain.HighlightBorderTopLeft
        .ColorSet.HighlightHot = COLOR_NativeXpPlain.HighlightHot
        .ColorSet.HighlightPressed = COLOR_NativeXpPlain.HighlightPressed
        .ColorSet.HighlightSelected = COLOR_NativeXpPlain.HighlightSelected
        
        .ColorSet.NormalGroupCaptionDark = COLOR_NativeXpPlain.NormalGroupCaptionDark
        .ColorSet.NormalGroupCaptionLight = COLOR_NativeXpPlain.NormalGroupCaptionLight
        .ColorSet.NormalGroupCaptionTextHot = COLOR_NativeXpPlain.NormalGroupCaptionTextHot
        .ColorSet.NormalGroupCaptionTextNormal = COLOR_NativeXpPlain.NormalGroupCaptionTextNormal
        .ColorSet.NormalGroupClient = COLOR_NativeXpPlain.NormalGroupClient
        .ColorSet.NormalGroupClientBorder = COLOR_NativeXpPlain.NormalGroupClientBorder
        .ColorSet.NormalGroupClientLink = COLOR_NativeXpPlain.NormalGroupClientLink
        
        .ColorSet.NormalGroupClientLinkHot = COLOR_NativeXpPlain.NormalGroupClientLinkHot
        .ColorSet.NormalGroupClientText = COLOR_NativeXpPlain.NormalGroupClientText
        .ColorSet.SpecialGroupCaptionDark = COLOR_NativeXpPlain.SpecialGroupCaptionDark
        .ColorSet.SpecialGroupCaptionLight = COLOR_NativeXpPlain.SpecialGroupCaptionLight
        .ColorSet.SpecialGroupCaptionTextHot = COLOR_NativeXpPlain.SpecialGroupCaptionTextHot
        .ColorSet.SpecialGroupCaptionTextSpecial = COLOR_NativeXpPlain.SpecialGroupCaptionTextSpecial
        .ColorSet.SpecialGroupClient = COLOR_NativeXpPlain.SpecialGroupClient
        .ColorSet.SpecialGroupClientBorder = COLOR_NativeXpPlain.SpecialGroupClientBorder
        .ColorSet.SpecialGroupClientLink = COLOR_NativeXpPlain.SpecialGroupClientLink
        .ColorSet.SpecialGroupClientLinkHot = COLOR_NativeXpPlain.SpecialGroupClientLinkHot
        .ColorSet.SpecialGroupClientText = COLOR_NativeXpPlain.SpecialGroupClientText
    End With
End Function

 Public Function IsCompiled() As Boolean
'得到出当前系统是否是编译后在运行
    
    On Error Resume Next
    Debug.Print 1 / 0
    If Err <> 0 Then
        '用源程序在运行
        IsCompiled = False
        Err.Clear
    Else
        IsCompiled = True
    End If
End Function

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error Resume Next
    
    strSQL = "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1"
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic")
    GetMaxLength = rs.Fields(0).DefinedSize

End Function

Public Function ParamCreate(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "参数名", adVarChar, 50
        .Fields.Append "参数值", adVarChar, 50
        
        .Open
    End With
    
    ParamCreate = True
    
    Exit Function
    
errHand:
    
End Function

Public Function ParamAdd(ByRef rs As ADODB.Recordset, ByVal strParamName As String, Optional ByVal strParamValue As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    
    rs("参数名").Value = strParamName
    rs("参数值").Value = strParamValue
    
    ParamAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function ParamRead(ByRef rs As ADODB.Recordset, ByVal strParamName As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.Filter = ""
    rs.Filter = "参数名='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        ParamRead = rs("参数值").Value
    End If
    
    Exit Function
    
errHand:
End Function

Public Function ParamWrite(ByRef rs As ADODB.Recordset, ByVal strParamName As String, ByVal strParamValue As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.Filter = ""
    rs.Filter = "参数名='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        rs("参数值").Value = strParamValue
    End If
    
    Exit Function
    
errHand:
End Function

Public Function MakeNO(ByVal intBillID As Integer, ByRef strNo As String, Optional ByVal lng科室ID As Long) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    '根据编号规则自动产生号码
    Dim intYear As Integer, strYear As String
    Dim intMonth As Integer, strMonth As String
    Dim str编号 As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    strNo = UCase(LTrim(strNo))
    intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
    strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
    intMonth = Month(zlDatabase.Currentdate())
    strMonth = intMonth
    strMonth = String(2 - Len(strMonth), "0") & strMonth
    
    gstrSQL = "Select 编号规则 From 号码控制表 Where 项目序号=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取部门性质", intBillID)
    
    Dim bln年度 As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    If zlCommFun.NVL(rsTemp!编号规则, 0) = 2 And lng科室ID <> 0 Then

    Else
        bln年度 = False
    End If
        
    If zlCommFun.NVL(rsTemp!编号规则, 0) = 0 Or bln年度 Then
        If Len(strNo) < 8 Then strNo = strYear & String(7 - Len(strNo), "0") & strNo
    ElseIf rsTemp!编号规则 = 2 Then
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取部门性质", intBillID, lng科室ID)
        If rsTemp.RecordCount = 0 Then
            ShowSimpleMsg "还未设置科室编号，无法产生号码！"
            Exit Function
        End If
        If zlCommFun.NVL(rsTemp!编号) = "" Then
            ShowSimpleMsg "还未设置科室编号，无法产生号码！"
            Exit Function
        End If
        str编号 = zlCommFun.NVL(rsTemp!编号)
        
        '小于四位，按本月产生号码
        '五位或六位，则认为是指定月份的号码
        '七位，则认为是产生本年指定科室、月份的号码
        '大于等于八位，不处理
        If Len(strNo) <= 4 Then
            strNo = strYear & str编号 & strMonth & String(4 - Len(strNo), "0") & strNo
        ElseIf Len(strNo) <= 6 Then
            strNo = String(6 - Len(strNo), "0") & strNo
            strNo = strYear & str编号 & strNo
        ElseIf Len(strNo) = 7 Then
            strNo = strYear & strNo
        End If
    Else
        ShowSimpleMsg "不支持这种编号规则！"
        Exit Function
    End If
    
    MakeNO = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
    
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function LongIDsTable(ByVal strIDs As String, ByRef idPar() As String, Optional ByVal idParStart As Long = 1, Optional ByVal Alias As String = "B", Optional ByVal ColumnAlias As String = "") As String
Dim strSQL As String, lngs As String, N As Integer, strReturn As String, strThis As String
    
    ReDim idPar(10) As String
    strSQL = "Select Column_Value " & IIf(ColumnAlias = "", "", " As ID ") & "From Table(Cast(f_Str2List([1]) As Zltools.t_StrList))"
    N = 0
    Do While True
        If Len(strIDs) <= 4000 Then
            strThis = strIDs
            strIDs = ""
        Else
            strThis = Mid(strIDs, 1, InStrRev(Mid(strIDs, 1, 4000), ",") - 1)
            strIDs = Mid(strIDs, InStrRev(Mid(strIDs, 1, 4000), ",") + 1)
        End If
        
        If N > 9 Then
            strReturn = strReturn & vbNewLine & " Union " & Replace(strSQL, "[1]", "'" & strThis & "'")
        Else
            idPar(N) = strThis
            strReturn = IIf(strReturn = "", "", strReturn & vbNewLine & " Union ") & Replace(strSQL, "[1]", "[" & (N + idParStart) & "]")
        End If
        
        N = N + 1
        If strIDs = "" Then Exit Do
    Loop
    
    LongIDsTable = " (" & strReturn & ") " & Alias & " "
End Function
