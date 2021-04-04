VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcnOracle As New Connection
Private mlngTXTProc As Long
Private mstrStOwners As String

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Or InStr(strInput, """") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
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
        objControl.Id = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
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
    
ErrHand:
    
End Function

Public Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal X As Single, ByVal Y As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngX As Long
    Dim lngY As Long
    Dim lngLoop As Long
    Dim lngXY As Long
            
    lngX = X / 15
    lngY = Y / 15
        
    lngXY = 2
    For lngLoop = 1 To 15
        lngXY = lngXY * 2
    Next
    
    lngXY = lngXY * lngY + lngX
    
    SendMessage lngHwnd, WM_LBUTTONDOWN, 0, ByVal lngXY
    SendMessage lngHwnd, WM_LBUTTONUP, 0, ByVal lngXY

End Sub

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
    
    On Error GoTo ErrHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.Id, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        cbrPopupItem.Parameter = cbrControl.Parameter
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Set cbrPopupItem2 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.Id, cbrControl2.Caption)
                cbrPopupItem2.Parameter = cbrControl2.Parameter
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
ErrHand:
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Function


Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
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

Public Function CallNewWindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call CallWindowProc(mlngTXTProc, hwnd, msg, wParam, lParam)
    
    CallNewWindowProc = True
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

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
    '******************************************************************************************************************
    '功能：将VB的系统颜色转换为RGB色
    '参数：
    '返回：
    '******************************************************************************************************************

    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    Select Case bytMode
    Case 1
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
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

Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function ZVal(ByVal varValue As Variant) As String
    '******************************************************************************************************************
    '功能：将0零转换为"NULL"串,在生成SQL语句时用
    '******************************************************************************************************************
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function CreateTmpFile(Optional ByVal strFileType As String = "tmp") As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim strFileTemp As String
       
    
    strFileTemp = GetTmpPath
    
    strFileTemp = strFileTemp & "zlNewPic" & Format(Now, "yyyymmdd") & Format(Timer, "0") & "." & strFileType
    
    CreateTmpFile = strFileTemp
End Function

Public Function GetTmpPath() As String
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    GetTmpPath = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
End Function

Public Function ParamCreate(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "参数名", adVarChar, 50
        .Fields.Append "参数值", adVarChar, 50
        
        .Open
    End With
    
    ParamCreate = True
    
    Exit Function
    
ErrHand:
    
End Function

Public Function ParamAdd(ByRef rs As ADODB.Recordset, ByVal strParamName As String, Optional ByVal strParamValue As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    rs.AddNew
    
    rs("参数名").value = strParamName
    rs("参数值").value = strParamValue
    
    ParamAdd = True
    
    Exit Function
    
ErrHand:
End Function

Public Function ParamRead(ByRef rs As ADODB.Recordset, ByVal strParamName As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    rs.Filter = ""
    rs.Filter = "参数名='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        ParamRead = rs("参数值").value
    End If
    
    Exit Function
    
ErrHand:
End Function

Public Function ParamWrite(ByRef rs As ADODB.Recordset, ByVal strParamName As String, ByVal strParamValue As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    rs.Filter = ""
    rs.Filter = "参数名='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        rs("参数值").value = strParamValue
    End If
    
    Exit Function
    
ErrHand:
End Function

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    '******************************************************************************************************************
    '
    '
    '
    '******************************************************************************************************************
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
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
    
    On Error GoTo ErrHand

    On Error Resume Next
    
    blnIcon = (rs("图标").name = "图标")
    blnIconSel = (rs("选中图标").name = "选中图标")
    
    On Error GoTo ErrHand
    
    LockWindowUpdate objTvw.hwnd

    Do While Not rs.EOF
        strTmp = AppendCode(rs("名称").value, rs("编码").value)
        
        If blnIcon Then strIcon = Nvl(rs("图标").value)
        If blnIconSel Then strIconSel = Nvl(rs("选中图标").value)
        If strIconSel = "" Then strIconSel = strIcon
        
        If IsNull(rs("上级id").value) Then
            Set objNode = objTvw.Nodes.Add(, , "K" & Nvl(rs("ID").value, 0), strTmp, strIcon, strIconSel)
        Else
            Set objNode = objTvw.Nodes.Add("K" & rs("上级id").value, tvwChild, "K" & Nvl(rs("ID").value, 0), strTmp, strIcon, strIconSel)
        End If

        rs.MoveNext
    Loop

    LockWindowUpdate 0

    LoadTree = True

    Exit Function
ErrHand:
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Function

Public Function CopyRecordStruct(ByVal rsFrom As ADODB.Recordset, Optional ByVal blnRowID As Boolean = False, Optional ByVal blnNotOpen As Boolean = False) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim lngLoop As Long
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic

    For lngLoop = 0 To rsFrom.Fields.Count - 1

        Select Case rsFrom.Fields(lngLoop).Type
        Case 135            'Oracle的Date型
            rs.Fields.Append rsFrom.Fields(lngLoop).name, adVarChar, 30, rsFrom.Fields(lngLoop).Attributes
        Case 5
            rs.Fields.Append rsFrom.Fields(lngLoop).name, rsFrom.Fields(lngLoop).Type, 30, rsFrom.Fields(lngLoop).Attributes
        Case Else
            rs.Fields.Append rsFrom.Fields(lngLoop).name, adVarChar, rsFrom.Fields(lngLoop).DefinedSize + 12
        End Select

    Next
    If blnRowID Then
        rs.Fields.Append "行号", adVarChar, 30
    End If
    
    If blnNotOpen = False Then rs.Open
    
    Set CopyRecordStruct = rs
    
    Exit Function
ErrHand:
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Function

Public Function CopyRecordData(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset, Optional blnAll As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    Dim lngLoop As Long
    
    On Error GoTo ErrHand
    
    If blnAll Then
        If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    End If
    
    Do While Not rsFrom.EOF
        rsTo.AddNew
        For lngLoop = 0 To rsFrom.Fields.Count - 1
            
            On Error Resume Next
            strTmp = ""
            strTmp = rsTo.Fields(rsFrom.Fields(lngLoop).name).name
            On Error GoTo ErrHand
            
            If UCase(strTmp) = UCase(rsFrom.Fields(lngLoop).name) Then
                rsTo.Fields(strTmp).value = Trim(Nvl(rsFrom.Fields(lngLoop).value))
            End If

        Next
        If blnAll = False Then Exit Do
        rsFrom.MoveNext
        rsTo.Update
    Loop
    
    CopyRecordData = True
    
    Exit Function
    
ErrHand:
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Function

Public Function DeleteRecordData(rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:删除记录集
    '参数:rs        要删除的记录集
    '返回:无
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    If Not (rs Is Nothing) Then
        If rs.RecordCount > 0 Then rs.MoveFirst
        While Not rs.EOF
            rs.Delete
            rs.MoveNext
        Wend
    End If
    
    DeleteRecordData = True
    
    Exit Function
    
ErrHand:
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Function

Public Sub CopyRecord(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset)
    '******************************************************************************************************************
    '功能:删除记录集
    '参数:rs        要删除的记录集
    '返回:无
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    On Error GoTo ErrHand
    
    Set rsTo = New ADODB.Recordset
    For lngLoop = 0 To rsFrom.Fields.Count - 1
        rsTo.Fields.Append rsFrom.Fields(lngLoop).name, rsFrom.Fields(lngLoop).Type, rsFrom.Fields(lngLoop).DefinedSize
    Next
    rsTo.Open
    
    If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    While Not rsFrom.EOF
        rsTo.AddNew
        For lngLoop = 0 To rsFrom.Fields.Count - 1
            rsTo.Fields(lngLoop).value = rsFrom.Fields(lngLoop).value
        Next
        rsFrom.MoveNext
    Wend
    
ErrHand:
    
End Sub

Public Sub DeleteRecord(rs As ADODB.Recordset)
    '******************************************************************************************************************
    '功能:删除记录集
    '参数:rs        要删除的记录集
    '返回:无
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    If rs.RecordCount > 0 Then rs.MoveFirst
    While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Wend
    
ErrHand:
End Sub

Public Function AppendRecord(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer
    
    Do While Not rsFrom.EOF
        rsTo.AddNew
        For intCol = 0 To rsFrom.Fields.Count - 1
            rsTo(rsFrom.Fields(intCol).name).value = Nvl(rsFrom.Fields(intCol).value)
        Next
        
        rsFrom.MoveNext
    Loop
    
    AppendRecord = True
    
End Function

Public Function AddPeriodToComboBox(ByRef cbo As Object, Optional ByVal blnParameter As Boolean)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    With cbo
        .Clear
        .addItem "今  天"
        .addItem "昨  天"
        .addItem "本  周"
        .addItem "本  月"
        .addItem "本  季"
        .addItem "本半年"
        .addItem "本  年"
        .addItem "前三天"
        .addItem "前一周"
        .addItem "前半月"
        .addItem "前一月"
        .addItem "前二月"
        .addItem "前三月"
        .addItem "前半年"
        .addItem "前一年"
        .addItem "前二年"
        If blnParameter = False Then .addItem "自定义"
    End With
    
    AddPeriodToComboBox = True
    
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

Public Function IntEx(vNumber As Variant) As Variant
    '******************************************************************************************************************
    '功能：取大于指定数值的最小整数
    '******************************************************************************************************************
    IntEx = -1 * Int(-1 * vNumber)
End Function

Public Function InDesign() As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error Resume Next
    Debug.Print 1 / 0
    If err.Number <> 0 Then err.Clear: InDesign = True
End Function

Public Function Lpad(ByVal strText As String, ByVal lngLen As Long, ByVal strReplace As String) As String
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim lngL As Long
    
    lngL = Len(strText)
    If lngL > lngLen Then
        Lpad = Left(strText, lngLen)
    ElseIf lngL < lngLen Then
        Lpad = String(lngLen - lngL, strReplace) & strText
    Else
        Lpad = strText
    End If
End Function

Public Sub DrawPicture(pic As Object, objPic As StdPicture, ByVal W As Long, ByVal H As Long)
    '******************************************************************************************************************
    '功能：在PictureBox中央按适当比例画一幅图
    '参数：W,H=要作图的尺寸
    '******************************************************************************************************************
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    
    If W <= pic.ScaleWidth And H <= pic.ScaleHeight Then
        lngW = W: lngH = H
    Else
        sngW = W / pic.ScaleWidth
        sngH = H / pic.ScaleHeight
        If sngW > sngH Then
            lngW = W / sngW: lngH = H / sngW
        Else
            lngW = W / sngH: lngH = H / sngH
        End If
    End If
    
    pic.Cls
    On Error Resume Next
    pic.PaintPicture objPic, (pic.ScaleWidth - lngW) / 2, (pic.ScaleHeight - lngH) / 2, lngW, lngH
    
End Sub

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    rs.AddNew
    rs("SQL").value = strSQL
    rs("Trans").value = intTrans
    rs("Custom").value = intCustom
    rs("Parameter").value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
ErrHand:
End Function

Public Function GetProgram(ByVal strName As String, Optional ByRef strOwner As String, Optional ByVal blnAdjustName As Boolean) As String
'功能：获取对象源码，并获取对象所有者
'参数：strName=对象名称
'      blnAdjustName=是否去掉过程的名称所带的双引号
'返回：对象源码
'          strOwner=对象所有者
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strProcText As String
    
    strOwner = ""
    strSQL = "Select Name,Text,Owner,Line From All_Source Where Owner In (Select Distinct 所有者 From zlSystems) And Name = [1] Order By Owner, Line "
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取过程源码", UCase(strName))
    If Not rsTmp.EOF Then
        strOwner = rsTmp!Owner & ""
        Do While Not rsTmp.EOF
            If rsTmp!Owner & "" <> strOwner Then Exit Do '防止多个用户存在该过程
            If rsTmp!Line = 1 And blnAdjustName Then
                '名称带双引号，则去掉
                If UCase(rsTmp!Text) Like "*" & """" & UCase(strName) & """" & "*" Then
                    strProcText = strProcText & Replace(UCase(rsTmp!Text), """" & UCase(strName) & """", strName)
                Else
                    strProcText = strProcText & rsTmp!Text
                End If
            Else
                strProcText = strProcText & rsTmp!Text
            End If
            rsTmp.MoveNext
        Loop
    End If
    GetProgram = "CREATE OR REPLACE " & strProcText
End Function

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

Public Sub LocationObj(ByRef objTxt As Object, Optional ByVal blnDoevents As Boolean = False)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    If blnDoevents Then DoEvents
    TxtSelAll objTxt
    objTxt.SetFocus
ErrHand:
    
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

Public Function GetNextId(strTable As String) As Long
    '------------------------------------------------------------------------------------
    '功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
    '参数：
    '   strTable：表名称
    '返回：
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTab As String
    
    '不能用错误错处理,原因是序列失效和没有序列时,应该返回错误,不然返回零,就有问题!
    '31730
    'On Error GoTo errH
    strTab = Trim(strTable)
    If strTab = "门诊费用记录" Or strTab = "住院费用记录" Then strTab = "病人费用记录"
    
    strSQL = "Select " & strTab & "_ID.Nextval From Dual"
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    GetNextId = rsTmp.Fields(0).value
'    Exit Function
'errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function OpenSQLRecord(ByVal cnOracle As ADODB.Connection, ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant, i As Long
    arrPars = arrInput
    Set OpenSQLRecord = OpenSQLRecordByArray(cnOracle, strSQL, strTitle, arrPars)
End Function

Public Function OpenSQLRecordByArray(ByVal cnOracle As ADODB.Connection, ByVal strSQL As String, ByVal strTitle As String, arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        If lngRight = 0 Then Exit Do
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        err.Raise 9527, strTitle, "SQL语句绑定变量不全，调用来源：" & strTitle
    End If

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
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
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next
    
    '执行返回记录集
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = cnOracle '这句比较慢(这句执行1000次约0.5x秒)
    'End If

    cmdData.CommandText = strSQL
    
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing

End Function

Public Sub ExecuteCmdText(ByVal strSQL As String, ByVal strFormCaption As String, Optional cnOracle As ADODB.Connection, Optional ByVal blnErrResume As Boolean)
'功能：执行无返回值语句
    If blnErrResume Then
        On Error Resume Next
    Else
        On Error GoTo errh
    End If
    If cnOracle Is Nothing Then
        gcnOracle.Execute strSQL, , adCmdText
    Else
        cnOracle.Execute strSQL, , adCmdText
    End If
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Public Function ShowSQLSelectEx(ByVal cnOracle As ADODB.Connection, frmParent As Object, ByVal objControl As Object, ByVal strSQL As String, bytStyle As Byte, _
                                                ByVal strTitle As String, ByVal bln末级 As Boolean, ByVal strSeek As String, ByVal strNote As String, _
                                                ByVal blnShowSub As Boolean, ByVal blnShowRoot As Boolean, ByVal blnNoneWin As Boolean, ByRef Cancel As Boolean, _
                                                ByVal blnMultiOne As Boolean, ByVal blnSearch As Boolean, ByVal blnMulti As Boolean, _
                                                ByVal strOtherInfo As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：多功能选择器，使用ADO.Command打开,允许使用[x]参数
'           (替代以前的ShowSelect、ShowSQLMultiSelect、ShowSQLSelect）
'           传入输入框对象，就可以自动定位
'参数：
'     frmParent=显示的父窗体
'     objControl=调用界面输入框
'     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
'     bytStyle=选择器风格
'       为0时:列表风格:ID,…
'       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
'       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
'             双表风格如果列名存在含Check结尾的字段，则该字段作为是否勾选的值存储字段。=1为勾选，0=不勾选。
'             双表风格如果列名存在*名称，*简码，*编码的，则显示右上角的查询功能，以供查询项目，
'                    编码列必须整个匹配，匹配成功后定位到该分类的该项目上，按F3支持查找下一个。
'     strTitle=选择器功能命名,也用于个性化区分
'     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
'     strSeek=当bytStyle<>2时有效,缺省定位的项目。
'             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
'             bytStyle=1时,可以是编码或名称
'     strNote=选择器的说明文字
'     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
'     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
'     blnNoneWin=处理成非窗体风格
'     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
'     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
'     blnSearch=是否显示行号,并可以输入行号定位
'     blnMulti=是否允许多选
'     strOtherInfo=格式为：项目名称1=内容1#项目2=内容2#......
'                           当前项目有：bytSize=0,1;字体大小(0-小字体,1-大字体;小字体为9号字,大字体为12号字),默认小字体
'                                              ColSet=列宽设置|列名1,宽度1;列名2,宽度2.....|悬浮提示|列名。
'                                              NotShowNon=0,1;0-默认处理，显示没有子项的分类，1-不显示没有子项的分类;bytStyle=2有作用
'     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型
'返回：取消=Nothing,选择=SQL源的单行记录集
'说明：
'     1.ID和上级ID可以为字符型数据
'     2.末级等字段不要带空值
'应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Dim frmNew As New frmPubSel
    Dim arrPar() As Variant
    arrPar = arrInput
    Set ShowSQLSelectEx = frmNew.ShowSelectV2(cnOracle, frmParent, objControl, strSQL, bytStyle, strTitle, bln末级, strSeek, strNote, _
                                                 blnShowSub, blnShowRoot, blnNoneWin, Cancel, blnMultiOne, blnSearch, blnMulti, _
                                                 strOtherInfo, arrPar)
End Function

'-------------------------------------------------------------------------------------------

Public Function GetOraConn(ByVal paraName As String) As String
    On Error GoTo ErrHand
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select 参数号,参数名,参数值 from zlOptions Where 参数名=[1]"
    Set rsData = OpenSQLRecord(gcnOracle, strSQL, "", paraName)
    If rsData.BOF = False Then
        GetOraConn = Nvl(rsData("参数值").value)
    End If
    Exit Function
ErrHand:
    MsgBox err.Description, vbCritical, gstrSysName
End Function

Public Function GetProcByState(ByVal intState As Integer, Optional intState2 As Integer) As ADODB.Recordset
    On Error GoTo ErrHand
    
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "Select 1 From zlProcedure Where 状态=[1]"
    If intState2 <> 0 Then
        strSQL = strSQL & " Or 状态=[2]"
    End If
    Set rsData = OpenSQLRecord(gcnOracle, strSQL, "", intState, intState2)
    Set GetProcByState = rsData
    
    Exit Function
ErrHand:
    MsgBox err.Description, vbCritical, gstrSysName
End Function

Public Function GetProcByOwner(ByVal strOwner As String) As ADODB.Recordset
    On Error GoTo ErrHand
    
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "Select A.ID,A.名称,B.内容 From zlProcedure A,zlProcedureText B Where A.ID = B.过程ID And B.性质=3 And A.所有者=[1] Order by A.Id,B.序号"
    Set rsData = OpenSQLRecord(gcnOracle, strSQL, "", strOwner)
    Set GetProcByOwner = rsData
    
    Exit Function
ErrHand:
    MsgBox err.Description, vbCritical, gstrSysName
End Function

Public Sub GetProcSQL(ByVal lngID As Long, ByVal pttType As ProcTextType, ByVal strSQL As String, ByRef arrSQL As Variant, Optional ByVal blnReadFile As Boolean)
'功能：blnReadFile 是否存文件读取SQL, True-strSQL为文件名，False-strSQL为要保存的SQL
    Dim lngCount As Long, i As Long, lng序号 As Long
    Dim strTmp As String, strStep As String
    Dim lngMaxLength As Long
    Dim objText  As TextStream
    
    On Error GoTo errh
    lngMaxLength = 3500 '最长4000，由于要经过拼接处理，因此取3500
    lng序号 = 1
    If blnReadFile Then
        Set objText = gobjFile.OpenTextFile(strSQL)
        Do While Not objText.AtEndOfStream
            strTmp = strTmp & vbCrLf & objText.ReadLine
        Loop
        strSQL = Mid(strTmp, Len(vbCrLf) + 1)
        Call GetProcSQL(lngID, pttType, strTmp, arrSQL)
    End If
    If ActualLen(SQLAdjust(strSQL)) < lngMaxLength Then
        Call addItem(arrSQL, "Zl_Zlproceduretext_Update(" & lngID & "," & pttType & "," & lng序号 & "," & SQLAdjust(strSQL) & ")")
    Else
        lngCount = Len(strSQL): strTmp = ""
        For i = 1 To lngCount Step 100
            strStep = Mid(strSQL, i, 100)
            If ActualLen(strTmp & strStep) >= lngMaxLength Then
                Call addItem(arrSQL, "Zl_Zlproceduretext_Update(" & lngID & "," & pttType & "," & lng序号 & "," & SQLAdjust(strTmp) & ")")
                lng序号 = lng序号 + 1
                If ActualLen(strStep) >= lngMaxLength Then
'                        strTmp=MidB(strStep,1,
                    Call addItem(arrSQL, "Zl_Zlproceduretext_Update(" & lngID & "," & pttType & "," & lng序号 & "," & SQLAdjust(strTmp) & ")")
                    lng序号 = lng序号 + 1
                End If
                strTmp = strStep
            Else
                strTmp = strTmp & strStep
            End If
        Next
        If strTmp <> "" Then
            Call addItem(arrSQL, "Zl_Zlproceduretext_Update(" & lngID & "," & pttType & "," & lng序号 & "," & SQLAdjust(strTmp) & ")")
        End If
    End If
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Function ExecuteProcedureBeach(ByVal cnOracle As ADODB.Connection, ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnTrans As Boolean = True, Optional blnCommit As Boolean = True) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集，可以为数组，也可以为集合，不能为其他类型
    '     strCaption -执行过程的父窗口标题
    '     blnTrans-是否存在事务
    '     blnCommit-执行完过程后,提交数据(前题:blnTrans=true)
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    On Error GoTo errh
    If cnOracle Is Nothing Then Set cnOracle = gcnOracle
    If blnTrans Then cnOracle.BeginTrans
    If TypeName(cllProcs) = "Collection" Then '集合形式
        For i = 1 To cllProcs.Count
            strSQL = cllProcs(i)
            Call ExecuteProcedure(strSQL, strCaption, cnOracle)
        Next
    ElseIf Not IsObject(cllProcs) Then
        If varType(cllProcs) = vbArray + vbVariant Or varType(cllProcs) = vbArray + vbString Then  '数组形式
            For i = LBound(cllProcs) To UBound(cllProcs)
                strSQL = cllProcs(i)
                Call ExecuteProcedure(strSQL, strCaption, cnOracle)
            Next
        End If
    End If
    If blnCommit And blnTrans Then
        cnOracle.CommitTrans
    End If
    ExecuteProcedureBeach = True
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    If blnCommit And blnTrans Then
        cnOracle.RollbackTrans
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'功能:读取指定字串的值,字串中可以包含汉字
 '入参:strInfor-原串
 '         lngStart-直始位置
'         lngLen-长度
'返回:子串
    Dim strTmp As String, i As Long
    err = 0: On Error GoTo errh:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
errh:
    err.Clear
    SubB = ""
End Function

Public Sub addItem(ByRef varItems As Variant, ByVal varItem As Variant, Optional ByVal blnClear As Boolean)
'功能：向数组中或集合中增加一个元素
'参数：varItems=集合或数组
'         varItem=要添加的元素
'         blnClear=是否清空数组中元素
'出参：varItems=增加元素后的集合或数组
    Dim i As Long
    
     If TypeName(varItems) = "Collection" Then
        If blnClear Then Set varItems = New Collection
         i = varItems.Count + 1
        varItems.Add varItem, "K" & i
     ElseIf Not IsObject(varItems) Then
         '数组形式，数组形式的varType(varItems) =vbArray+数组元素类型，如string()=vbArray+vbString
         If varType(varItems) >= vbArray Then
            On Error Resume Next
            i = UBound(varItems)
            If err.Number <> 0 Or blnClear Then
                varItems = Array()
                 If err.Number <> 0 Then err.Clear
            End If
            On Error GoTo 0
            ReDim Preserve varItems(UBound(varItems) + 1)
            varItems(UBound(varItems)) = varItem
        End If
     End If
End Sub

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    Set rs = OpenSQLRecord(gcnOracle, "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1", "mdlMedical")
    GetMaxLength = rs.Fields(0).DefinedSize
End Function

Public Function GetSystems(Optional varInput As Variant) As ADODB.Recordset
'功能：获取系统所有者
'varInput:=字符串类型：为系统所有者,若为数字，则自动转换为数值类型处理
'                数字类型为：系统序号相关。varInput>=100:系统编号.        0< varInput <100:系统简码即(系统号\100)
'返回：系统信息记录集
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngSys As Long, lngSysCode As Long, strUser As String
    
    strSQL = "Select 编号, 名称, 共享号, Upper(所有者) 所有者, 安装日期, 正常安装, 版本号 From Zlsystems Order By 共享号 Desc, 编号"
    If TypeName(varInput) = "Error" Then
    ElseIf IsNumeric(varInput) Then
        lngSys = Val(varInput)
        If lngSys < 100 Then
            strSQL = "Select 编号, 名称, 共享号, Upper(所有者) 所有者, 安装日期, 正常安装, 版本号" & vbNewLine & _
                            "From Zlsystems" & vbNewLine & _
                            "Where Floor(编号 / 100) = [3]" & vbNewLine & _
                            "Order By 共享号 Desc, 编号"
            lngSysCode = lngSys
        Else
            strSQL = "Select 编号, 名称, 共享号, Upper(所有者) 所有者, 安装日期, 正常安装, 版本号 From Zlsystems Where 编号 = [2]"
        End If
    ElseIf TypeName(varInput) = "String" Then
        If varInput <> "" Then
            strSQL = "Select 编号, 名称, 共享号, Upper(所有者) 所有者, 安装日期, 正常安装, 版本号" & vbNewLine & _
                            "From Zlsystems" & vbNewLine & _
                            "Where Upper(所有者) = [1]" & vbNewLine & _
                            "Order By 共享号 Desc, 编号"
            strUser = varInput
        End If
    End If
    On Error GoTo errh
    Set rsTmp = OpenSQLRecord(gcnOracle, strSQL, "GetSystems", UCase(strUser), lngSys, lngSysCode)
    Set GetSystems = rsTmp
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbCritical, gstrSysName
End Function

Public Function GetMenSystems(Optional ByVal blnAllSys As Boolean, Optional ByVal blnOnlyShowMainSystems As Boolean = False) As ADODB.Recordset
'功能：获取具有人员管理的所有系统
'参数：
'      blnAllSys:展示所有系统
'      blnOnlyShowMainSystems:仅展示主系统，不展示共享系统
'返回：系统信息记录集
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngSys As Long, lngSysCode As Long, strUser As String
    
    strSQL = "Select a.编号, a.名称, a.共享号, Upper(a.所有者) 所有者, a.安装日期, a.正常安装, a.版本号,Decode(b.Owner,Null,-1,1) 人员管理" & vbNewLine & _
                    "From Zlsystems a," & vbNewLine & _
                    "     (Select Owner" & vbNewLine & _
                    "       From All_Tables" & vbNewLine & _
                    "       Where Table_Name In ('部门表', '人员表', '部门人员', '上机人员表')" & vbNewLine & _
                    "       Group By Owner" & vbNewLine & _
                    "       Having Count(Owner) = 4) b" & vbNewLine & _
                    "Where a.所有者 = b.Owner" & IIf(blnAllSys, "(+)", "") & IIf(blnOnlyShowMainSystems, " And a.共享号 Is Null", "") & vbNewLine & _
                    "Order By a.共享号 Desc, a.编号"
    On Error GoTo errh
    Set rsTmp = OpenSQLRecord(gcnOracle, strSQL, "GetMenSystems")
    Set GetMenSystems = rsTmp
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbCritical, gstrSysName
End Function

Public Function IsStSystemUser(ByVal strUser As String) As Boolean
'功能：判断当前用户是否是标准版的所有者
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select 1 From Zlsystems Where Trunc(To_Number(编号) / 100) = 1 And Upper(所有者) ='" & UCase(strUser) & "'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "IsStSystemUser")
    IsStSystemUser = rsTmp.RecordCount > 0
    
End Function

Public Function ReadLob(ByVal cnOracle As ADODB.Connection, ByVal Action As Long, ByVal KeyWord As String, _
                        Optional ByVal strFile As String) As String
'功能：将指定的LOB字段复制为临时文件
'调用：ZLTOOLS.b_Public.Zllobread
'参数：
'Action:操作类型（用以区别是操作哪个表）
'0-zlUnitInfoImage
'KeyWord:确定数据记录的关键字
'strFile:用户指定存放的文件名；不指定时，自动取临时文件名
'返回：存放内容的文件名，失败则返回零长度""
    Const conChunkSize As Integer = 10240
    Dim rsLob As ADODB.Recordset
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim strSQL As String
    Dim objFile As New FileSystemObject
    
    err = 0: On Error GoTo ErrHand
    strSQL = "Select ZLTOOLS.b_Public.Zllobread([1],[2],[3]) as 片段 From Dual"
    If strSQL = "" Then strFile = "": Exit Function
    If strFile = "" Then
        strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & objFile.GetTempName
    End If
    lngFileNum = FreeFile
    Open strFile For Binary As lngFileNum
    lngCount = 0
    Do
        Set rsLob = OpenSQLRecord(cnOracle, strSQL, "zllobRead", Action, KeyWord, lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).value) Then Exit Do
        strText = rsLob.Fields(0).value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    ReadLob = strFile
    Exit Function
ErrHand:
    Close lngFileNum
    Kill strFile: ReadLob = ""
    err.Clear
End Function

Public Function SaveLob(ByVal cnOracle As ADODB.Connection, ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String) As Boolean
'功能：保存指定的文件到指定表记录BLOB字段
'调用：ZLTOOLS.b_Public.Zllobappend
'参数：
'cnOracle:保存用的连接
'Action:操作类型（用以区别是操作哪个表）
'       0-zlUnitInfoImage
'KeyWord:确定数据记录的关键字
'strFile:用户要上传的图片路径
'返回：成功返回True，失败返回False
    Dim arrSQL() As Variant

    If GetLobSql(Action, KeyWord, strFile, arrSQL) Then
        Call ExecuteProcedureBeach(cnOracle, arrSQL, "zllobSave", False, False)
    Else
        SaveLob = False
    End If
    SaveLob = True
    Exit Function
ErrHand:
    err.Clear
    SaveLob = False
End Function

Public Function GetLobSql(ByVal Action As Long, ByVal KeyWord As String, _
                ByVal strFile As String, ByRef arySql() As Variant) As Boolean
'功能：产生保存指定的文件到指定表记录BLOB的SQL语句
'调用：ZLTOOLS.b_Public.Zllobappend
'参数：
'Action:操作类型（用以区别是操作哪个表）
'       0-zlUnitInfoImage
'KeyWord:确定数据记录的关键字
'strFile:用户要上传的图片路径，空路径，则表示删除数据
'arySql():在该数据的基础上扩展增加保存的SQL语句
'返回：成功返回True，失败返回False
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    Dim strSQLRoot As String
    Dim lngLBound As Long, lngUbound As Long    '传入数组的最小最大下标
    Dim strSubTxt As String
    
    err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUbound = UBound(arySql)
    If err <> 0 Then lngLBound = 0: lngUbound = -1
    err = 0: On Error GoTo 0
    strSQLRoot = "ZLTOOLS.b_Public.Zllobappend(" & Action & ",'" & KeyWord & "'"
    '空文件名，则删除数据
    If strFile = "" Then
        ReDim Preserve arySql(lngUbound + 1)
        arySql(lngUbound + lngCount + 1) = strSQLRoot & ",Null,1)"
        GetLobSql = True:
        Exit Function
    End If
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    err = 0: On Error GoTo ErrHand
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    ReDim Preserve arySql(lngLBound To lngUbound + lngBlocks + 1)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        arySql(lngUbound + lngCount + 1) = strSQLRoot & ",'" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
    Next
    Close lngFileNum
    GetLobSql = True
    Exit Function
ErrHand:
    err.Clear
    Close lngFileNum
    GetLobSql = False
End Function

Public Sub GetRegInfoSQL(ByVal str项目 As String, ByVal strInfo As String, ByVal strStation As String, ByRef arrSQL As Variant)
'功能：blnReadFile 是否存文件读取SQL, True-strSQL为文件名，False-strSQL为要保存的SQL
    Dim lngCount As Long, i As Long, lng序号 As Long
    Dim strTmp As String, strStep As String
    Dim lngMaxLength As Long
    
    On Error GoTo errh
    lngMaxLength = 3000 '最长4000，由于要经过拼接处理，因此取3500
    If strInfo <> "" Then
        If ActualLen(SQLAdjust(strInfo)) < lngMaxLength Then
            Call addItem(arrSQL, "Zltools.b_Public.Zlreginfoupdate('" & str项目 & "',Null," & SQLAdjust(strInfo) & ",'" & strStation & "')")
        Else
            lngCount = Len(strInfo): strTmp = "": lng序号 = 0
            For i = 1 To lngCount Step 100
                strStep = Mid(strInfo, i, 100)
                If ActualLen(strTmp & strStep) >= lngMaxLength Then
                    lng序号 = lng序号 + 1
                    Call addItem(arrSQL, "Zltools.b_Public.Zlreginfoupdate('" & str项目 & "'," & lng序号 & "," & SQLAdjust(strTmp) & ",'" & strStation & "')")
                    strTmp = strStep
                Else
                    strTmp = strTmp & strStep
                End If
            Next
            If strTmp <> "" Then
                lng序号 = lng序号 + 1
                Call addItem(arrSQL, "Zltools.b_Public.Zlreginfoupdate('" & str项目 & "'," & lng序号 & "," & SQLAdjust(strTmp) & ",'" & strStation & "')")
            End If
        End If
    Else
        Call addItem(arrSQL, "Zltools.b_Public.Zlreginfoupdate('" & str项目 & "',Null,Null,'" & strStation & "')")
    End If
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Public Function UpdateZLReginfo(ByVal strInfoName As String, Optional ByVal strValue As String, Optional ByVal intType As Integer) As Boolean
'功能：进行ZLReginfo数据修正
'参数：strInfoName=项目名称
'      strValue=项目内容
'intType=0-更新,1-新增,2-自动识别新增或更新，-1删除
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errh
    If intType = 2 Then
        strSQL = "Select 1 From Zlreginfo Where 项目 =[1]"
        Set rsTmp = OpenSQLRecord(gcnOracle, strSQL, "UpdateZLReginfo", strInfoName)
        intType = IIf(rsTmp.EOF, 1, 0)
    End If
    
    If intType = 0 Then
        strSQL = "Update Zlreginfo Set 内容 = '" & strValue & "' Where 项目 = '" & strInfoName & "'"
    ElseIf intType = 1 Then
        strSQL = "Insert Into Zlreginfo (项目, 内容) Values ('" & strInfoName & "', '" & strValue & "')"
    ElseIf intType = -1 Then
        strSQL = "Delete Zlreginfo Where 项目 = '" & strInfoName & "'"
    End If
    gcnOracle.Execute strSQL, , adCmdText
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, _
                                            Optional strPrivate As String, Optional strLocal As String, Optional ByVal strDefault As String) As String
'功能：读取指定的参数值
'参数：varPara=参数号或参数名，以数字或字符类型传入区分
'      lngSys=使用该参数的系统编号，如100
'      lngModual=使用该参数的模块号，如1230
'      strPrivate-私有参数  strLocal-本机参数
'      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
'返回：参数值，字符串形式
    Dim strSQL As String
    Dim strSqlFilter As String, strSqlUser As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errh
    '过滤条件
    If TypeName(varPara) = "String" Then
        strSqlFilter = "Where 参数名=[3] And Nvl(模块,0)=[2] And Nvl(系统,0)= [1] "
    Else
        strSqlFilter = "Where 参数号=[4] And Nvl(模块,0)=[2] And Nvl(系统,0)= [1] "
    End If

    '参数表，获取参数特征
    If strPrivate + strLocal <> "" Then '本机/私有参数
        strSQL = "Select Nvl(B.参数值,0) 参数值" & vbNewLine & _
                        "From zlParameters A , zlUserParas B " & strSqlFilter & vbNewLine & _
                        " And B.参数ID = A.ID  And Nvl(机器名,0) =[5] And Nvl(用户名,0) = [6] "
    Else
        strSQL = "Select Nvl(A.参数值,缺省值) as 参数值" & vbNewLine & _
                "From zlParameters A " & strSqlFilter
    End If
    
    Set rsTmp = OpenSQLRecord(gcnOracle, strSQL, "GetPara", Val(lngSys), Val(lngModual), Nvl(CStr(varPara), 0), Val(varPara), IIf(strLocal = "", 0, strLocal), IIf(strPrivate = "", 0, strPrivate))
    If Not rsTmp.EOF Then
        GetPara = rsTmp!参数值 & ""
    Else
        GetPara = strDefault
    End If
    Exit Function
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

