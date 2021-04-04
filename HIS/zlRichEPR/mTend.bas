Attribute VB_Name = "mTend"

Option Explicit

Public glngTXTProc As Long


'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    
    Dim RS As New ADODB.Recordset
    
    On Error Resume Next
    
    Set RS = zlDatabase.OpenSQLRecord("SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1", "mdlPublic")
    
    GetMaxLength = RS.Fields(0).DefinedSize
    
End Function

Public Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '功能：重新定义数组
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function Zero(ByVal varValue As Variant, Optional ByVal varNewValue As Variant = "") As Variant
    Zero = IIf(Val(varValue) = 0, varNewValue, varValue)
End Function

Public Function SetColorIcon(frmMain As Form, ByVal Key As String, ByVal COLOR As OLE_COLOR, ByRef ils As ImageList) As Boolean
    
    Dim ctlPictureBox As VB.PictureBox
    
    On Error GoTo ErrHand
    
    Set ctlPictureBox = frmMain.Controls.Add("VB.PictureBox", "ctlPictureBox1")
    
    Dim ListImage As ListImage
    Set ListImage = ils.ListImages("User")
    
    ctlPictureBox.AutoRedraw = True
    ctlPictureBox.AutoSize = True
    ctlPictureBox.BackColor = ils.MaskColor
    
    ctlPictureBox.Picture = ListImage.ExtractIcon
    
    If COLOR = vbWhite Then COLOR = RGB(254, 254, 254)
    
    ctlPictureBox.Cls
    ctlPictureBox.Line (30, 30)-(ctlPictureBox.Width - 105, ctlPictureBox.Height - 105), COLOR, BF
    ctlPictureBox.Refresh

    'Replace icon
    On Error Resume Next
    ils.ListImages.Remove ils.ListImages(Key).Index
    On Error GoTo ErrHand
    
    ils.ListImages.Add , Key, ctlPictureBox.Image
    ils.ListImages(Key).Tag = CStr(COLOR)

    frmMain.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
    
    SetColorIcon = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub LocationObj(ByRef objTXT As Object)
    On Error Resume Next
    
    zlControl.TxtSelAll objTXT
    objTXT.SetFocus
End Sub

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
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
    Case 3      '非数字
        If InStr("0123456789", Chr(KeyAscii)) > 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
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
    Case 2
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) > 0 Then
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

Public Function ShowTxtFilterDialog(ByVal frmParent As Object, _
                                    ByVal objTXT As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 7500, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True, _
                                    Optional ByVal blnExecuted As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能;显示文本输入选择列表(只用于文本框控件)
    '------------------------------------------------------------------------------------------------------------------
    
    Dim RS As New ADODB.Recordset
    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo ErrHand
    
    If blnExecuted = False Then
        If InStr(objTXT.Text, "'") > 0 Then Exit Function
    Else
        Set RS = rsResult
    End If
    
    '参数初始化
    Call ClientToScreen(objTXT.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTXT.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
    
    If blnExecuted = False Then
        '执行查询
        Set RS = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption)
        If RS.BOF Then
            If blnPrompt Then MsgBox "没有找到相匹配的结果！", , gstrSysName
            Exit Function                            '没有结果，直接返回
        End If
    End If
    
    If RS.RecordCount = 1 And blnFilter Then GoTo Over                    '因为是输入查找，如果只有一条，则直接返回
    If frmSelectDialog.ShowSelect(frmParent, 2, RS, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTXT.Height, , strSavePath, , False) Then GoTo Over
    Exit Function
    
Over:
    
    Set rsResult = RS
    
    ShowTxtFilterDialog = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowTxtSelectDialog(ByVal frmParent As Object, _
                                    ByVal objTXT As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False, _
                                    Optional strInitKey As String = "", _
                                    Optional ByVal WinStyle As Byte = 3) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开树型+列表结构
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim RS As New ADODB.Recordset
    Dim objPoint As POINTAPI
        
    If Trim(strSQL) = "" Then Exit Function
    
    On Error GoTo ErrHand
    
    Set RS = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption)
    If RS.BOF Then
        MsgBox "没有可选择的数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objTXT.hWnd, objPoint)
                
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTXT.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
    
    If frmSelectDialog.ShowSelect(frmParent, WinStyle, RS, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTXT.Height, strInitKey, strSavePath, , False, blnMuliSel) Then
                            
        Set rsResult = RS
        ShowTxtSelectDialog = True
        
    End If
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
    
End Function

Public Function ShowGrdSelectDialog(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional ByVal blnMuliSel As Boolean = False, _
                                    Optional ByVal bytStyle As Byte = 3) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开树型+列表结构,应用于表格控件
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim RS As New ADODB.Recordset
    Dim objPoint As POINTAPI
        
    If Trim(strSQL) = "" Then Exit Function
    
    On Error GoTo ErrHand
    
    Set RS = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption)
    If RS.BOF Then
        MsgBox "没有可选择的数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objVsf.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight
    
    If frmSelectDialog.ShowSelect(frmParent, bytStyle, RS, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objVsf.CellHeight, , strSavePath, , False, blnMuliSel) Then
                            
        Set rsResult = RS
        ShowGrdSelectDialog = True
        
    End If
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
    
End Function

Public Function ShowGrdSelDialog(ByVal frmParent As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal objVsf As Object, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal strSQL As String, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 6000, _
                                Optional ByVal lngCY As Long = 3000, _
                                Optional ByVal blnFilter As Boolean = True, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal blnPrompt As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能;显示文本输入选择列表(只用于表格控件)
    '------------------------------------------------------------------------------------------------------------------
    
    Dim RS As New ADODB.Recordset
    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo ErrHand

    If InStr(objVsf.EditText, "'") > 0 Then Exit Function
        
    Call ClientToScreen(objVsf.hWnd, objPoint)
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight
            
    If rsResult.BOF Then
        If blnPrompt Then MsgBox "没有找到相匹配的结果！", , gstrSysName
        Exit Function                            '没有结果，直接返回
    End If
    
    If rsResult.RecordCount = 1 And blnFilter Then GoTo Over                    '因为是输入查找，如果只有一条，则直接返回
    If frmSelectDialog.ShowSelect(frmParent, bytStyle, rsResult, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objVsf.CellHeight, , strSavePath, , False, blnMuliSel) Then GoTo Over
    Exit Function
    
Over:
    
    Set rsResult = rsResult
    
    ShowGrdSelDialog = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowGrdFilterDialog(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能;显示文本输入选择列表(只用于表格控件)
    '------------------------------------------------------------------------------------------------------------------
    
    Dim RS As New ADODB.Recordset
    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo ErrHand

    If InStr(objVsf.EditText, "'") > 0 Then Exit Function
        
    Call ClientToScreen(objVsf.hWnd, objPoint)
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight
            
    '执行查询
    Set RS = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption)
    If RS.BOF Then
        If blnPrompt Then MsgBox "没有找到相匹配的结果！", , gstrSysName
        Exit Function                            '没有结果，直接返回
    End If
            
    If RS.RecordCount = 1 And blnFilter Then GoTo Over                    '因为是输入查找，如果只有一条，则直接返回
    If frmSelectDialog.ShowSelect(frmParent, 2, RS, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objVsf.CellHeight, , strSavePath, , False) Then GoTo Over
    Exit Function
    
Over:
    
    Set rsResult = RS
    
    ShowGrdFilterDialog = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    'Or InStr(strInput, ";") > 0 Or InStr(strInput, ",") > 0 Or InStr(strInput, "`") > 0 Or InStr(strInput, """") > 0
    If InStr(strInput, "'") > 0 Then
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

Public Function CreateParam(ByRef RS As ADODB.Recordset, ByVal strParamName As String, ByVal dteType As DataTypeEnum, Optional ByVal lngSize As Long) As Boolean
    
    If RS Is Nothing Then Set RS = New ADODB.Recordset
    
    If lngSize > 0 Then
        RS.Fields.Append strParamName, dteType, lngSize
    Else
        RS.Fields.Append strParamName, dteType
    End If
    
End Function

Public Function GetCombList(ByVal strSQL As String) As String
    
    Dim RS As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    Set RS = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical")
    If RS.BOF = False Then
        Do While Not RS.EOF
            GetCombList = GetCombList & "|" & zlCommFun.NVL(RS.Fields(0).Value)
            RS.MoveNext
        Loop
    End If
    If GetCombList = "" Then
        GetCombList = " |"
    Else
        GetCombList = Mid(GetCombList, 2)
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CreateVsf(ByRef objVsf As Object, ByVal strVsf As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim varArray As Variant
    Dim varItem As Variant
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    objVsf.Cols = 0
    
    varArray = Split(strVsf, ";")
    For lngLoop = 0 To UBound(varArray)
        varItem = Split(varArray(lngLoop), ",")
                
        objVsf.Cols = objVsf.Cols + 1
        i = objVsf.Cols - 1
    
        objVsf.TextMatrix(0, i) = varItem(0)
        objVsf.ColWidth(i) = Val(varItem(1))
        objVsf.ColAlignment(i) = Val(varItem(2))
        objVsf.ColHidden(i) = (Val(varItem(4)) = 0)
        objVsf.Cell(flexcpData, 0, i) = IIf(varItem(5) = "", varItem(0), varItem(5))
        
    Next
    
    CreateVsf = True
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
'功能：判断当前屏幕鼠标是否在指定窗口的显示区域内
    Dim vRect As RECT, vPos As POINTAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
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

Public Function CheckNumber(ByVal strText As String, ByVal intLen As Integer, Optional ByVal intDec As Integer = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intSplit As Integer
    
    If Trim(strText) = "" Or Val(strText) = 0 Then
        CheckNumber = True
        Exit Function
    End If
    
    intSplit = InStr(strText, ".")
    
    If intSplit = 0 Then
        '整数
        
        If Len(strText) > (intLen - intDec) Then
            Exit Function
        End If
    Else
        If (intSplit - 1) > (intLen - intDec) Then Exit Function
        If (Len(strText) - intSplit) > intDec Then Exit Function
        
    End If
    
    CheckNumber = True
    
End Function

Public Function LocationGrid(vsf As Object, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error Resume Next
    
    vsf.Row = lngRow
    vsf.Col = lngCol
    vsf.ShowCell vsf.Row, vsf.Col
    vsf.SetFocus
    
    LocationGrid = True
End Function

