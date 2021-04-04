Attribute VB_Name = "mdlLisFace"
Option Explicit

Public glngTXTProc As Long
Public gstrSysName As String
Public gstrDBUser As String
Public gstrSQL As String
Public gstrPrive As String

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, Optional ByVal strFormCaption As String)
'功能：打开记录。同时保存SQL语句
    If rsTemp.State = adStateOpen Then rsTemp.Close

    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly

End Sub

Public Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    
    Dim blnForeColor As Boolean
    Dim blnBkColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("前景色").Name = "前景色")
    blnBkColor = (rsData("背景色").Name = "背景色")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
        lngRow = 0
    Else
        
        If Val(objMsf.RowData(objMsf.Rows - 1)) <= 0 Then
            lngRow = objMsf.Rows - 2
        Else
            lngRow = objMsf.Rows - 1
        End If
                
    End If
    
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            If Trim(objMsf.TextMatrix(0, lngLoop)) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If strMask <> "" Then
                    objMsf.TextMatrix(lngRow, lngLoop) = Format(NVL(rsData(objMsf.TextMatrix(0, lngLoop))), strMask)
                Else
                    objMsf.TextMatrix(lngRow, lngLoop) = NVL(rsData(objMsf.TextMatrix(0, lngLoop)))
                End If
            End If
            
        Next
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("前景色").Value)
        If blnBkColor Then objMsf.Cell(flexcpBackColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("背景色").Value)
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then Resume
End Function

Public Sub LocationObj(ByRef objTxt As Object)
    On Error Resume Next
    
    TxtSelAll objTxt
    objTxt.SetFocus
End Sub



Public Sub LocationGrid(ByRef vsf As Object, Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1)
    
    On Error Resume Next
    
    If lngRow <> -1 Then vsf.Row = lngRow
    If lngCol <> -1 Then vsf.Col = lngCol
    
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Public Function GetCol(ByVal objVsf As Object, ByVal strData As String) As Long
    
    Dim lngLoop As Long
    
    GetCol = -1
    For lngLoop = 0 To objVsf.Cols - 1
        If objVsf.Cell(flexcpText, 0, lngLoop, 0, lngLoop) = strData Then
            GetCol = lngLoop
            Exit Function
        End If
    Next
    
End Function


Public Function CreateVsf(ByRef objVsf As Object, ByVal strVsf As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim varArray As Variant
    Dim varItem As Variant
    Dim i As Integer
    
    On Error GoTo errHand
    
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
    
errHand:
'    If ErrCenter = 1 Then Resume
End Function

Public Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    On Error GoTo errHand
    
'    Exit Function
    
    If objVsf.Rows = 0 Then Exit Function
    lngTop = objVsf.Cell(flexcpTop, objVsf.Rows - 1, 0) + objVsf.RowHeight(objVsf.Rows - 1)
    
    '1.隐藏所有的线
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeightMin) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeightMin + IIf(lngIndex = 1, 30, 0)
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
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

Public Function NextNo(intBillID As Integer, Optional ByVal intStep As Integer = 1) As Variant
'功能：根据特定规则产生新的号码,规则如下：
'   一、项目序号：
'   1   病人ID         数字
'   2   住院号         数字
'   3   门诊号         数字
'   10  医嘱发送号     数字,顺序递增编号
'   x   其它单据号     字符,根据编号规则顺序递增编号,不自动补缺
'   二、年度位确定原则:
'       以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码

    Dim rsCtrl As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim vntNo As Variant, strSQL As String
    Dim intYear, strYear As String
    Dim curDate As Date, blnByDate As Boolean
ReStart:
    Err = 0
    On Error GoTo errHand

    If intBillID = 1 Then '病人ID
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            strSQL = "Select Nvl(Max(病人ID),0)+1 as 病人ID From 病人信息 Where 病人ID>=" & vntNo
            
            With rsTmp
                If .State = adStateOpen Then .Close
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 2 Then '住院号
        '顺序编号还是日期编号
        strSQL = "Select A.*,Sysdate as 日期 From 系统参数表 A Where A.参数号=27"
        With rsTmp
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!参数值), 1, !参数值) = 2)
                curDate = !日期
            End If
        End With
        
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(住院号),0)+1 as 住院号 From 病人信息 Where 住院号>=" & vntNo
            Else
                strSQL = "Select Nvl(Max(住院号),To_Number(To_Char(Sysdate,'YYMM')||'0000'))+1 as 住院号" & _
                    " From 病人信息 Where 住院号 Like To_Number(To_Char(Sysdate,'YYMM'))||'%' And 住院号>=" & vntNo
            End If
            
            With rsTmp
                If .State = adStateOpen Then .Close
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "最大号码", IIf(vntNo - 10 > Val(Format(curDate, "YYMM0000")), vntNo - 10, Val(Format(curDate, "YYMM0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 3 Then '门诊号
        '顺序编号还是日期编号
        strSQL = "Select A.*,Sysdate as 日期 From 系统参数表 A Where A.参数号=46"
        With rsTmp
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!参数值), 1, !参数值) = 2)
                curDate = !日期
            End If
        End With
    
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(门诊号),0)+1 as 门诊号 From 病人信息 Where 门诊号>=" & vntNo
            Else
                strSQL = "Select Nvl(Max(门诊号),To_Number(To_Char(Sysdate,'YYMMDD')||'0000'))+1 as 门诊号" & _
                    " From 病人信息 Where 门诊号 Like To_Number(To_Char(Sysdate,'YYMMDD'))||'%' And 门诊号>=" & vntNo
            End If
            
            With rsTmp
                If .State = adStateOpen Then .Close
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "最大号码", IIf(vntNo - 10 > Val(Format(curDate, "YYMMDD0000")), vntNo - 10, Val(Format(curDate, "YYMMDD0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 10 Then '医嘱发送号
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillID
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            vntNo = Val(IIf(IsNull(!最大号码), 0, !最大号码)) + 1
            
            On Error Resume Next
            .Update "最大号码", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
        
    ElseIf intBillID = 81 Then      '体检编号
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillID
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            vntNo = Val(IIf(IsNull(!最大号码), 0, !最大号码)) + 1
            
            On Error Resume Next
            .Update "最大号码", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    Else
        
        Dim vtnEndNo As Variant
        
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillID
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            intYear = Format(!Today, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            vntNo = IIf(IsNull(!最大号码), "", !最大号码)
            
            If IIf(IsNull(!编号规则), 0, !编号规则) = 1 Then
                '按日顺序编号
                If vntNo < strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000" Then
                    vntNo = strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000"
                End If
                vtnEndNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + intStep), 4)
                vntNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + 1), 4)
            Else
                '按年顺序编号
                If Left(vntNo, 1) < strYear Then
                    vntNo = strYear & "0000000"
                End If
                vtnEndNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + intStep), 7)
                vntNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + 1), 7)
            End If
            
            If Not (UCase(strYear) >= "A" And UCase(strYear) <= "Z") Or ActualLen(vntNo) > 8 Then GoTo ReStart
            
            On Error Resume Next
            .Update "最大号码", vtnEndNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    End If
    Exit Function
errHand:
    'If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
    NextNo = Null
End Function

Public Function GetNextPatientID() As Long
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select max(病人id) as 病人id From 病人信息 "
    rs.Open strSQL, gcnOracle
    If rs.BOF = False Then
        GetNextPatientID = NVL(rs("病人id"), 0)
    End If
    GetNextPatientID = GetNextPatientID + 1
    
End Function

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'功能: 装载数据入指定的组合下拉框或网格中的下拉框中
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            
            If rsTemp1.Fields.Count > 2 Then
                If Val(rsTemp1.Fields(2).Value) = 1 Then
                    objSource.ListIndex = objSource.NewIndex
                End If
            End If
            
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Public Function ShowTxtSelectDialog(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开树型+列表结构
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
        
    If Trim(strSQL) = "" Then Exit Function
    
    On Error GoTo errHand
    
    Call OpenRecord(rs, strSQL, frmParent.Caption, adOpenStatic, adLockBatchOptimistic)
    If rs.BOF Then
        MsgBox "没有可选择的数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objTxt.hWnd, objPoint)
                
    lngX = objPoint.x * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
    
    If frmSelectDialog.ShowSelect(frmParent, 3, rs, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False, blnMuliSel) Then
                            
        Set rsResult = rs
        ShowTxtSelectDialog = True
        
    End If
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then Resume
    
End Function

Public Function ShowTxtFilterDialog(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
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
    '功能;显示文本输入选择列表(只用于文本框控件)
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    Dim strInput As String
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo errHand

    If InStr(objTxt.Text, "'") > 0 Then Exit Function
    
    '参数初始化
    strInput = "'%" & UCase(objTxt.Text) & "%'"
    Call ClientToScreen(objTxt.hWnd, objPoint)
    
    lngX = objPoint.x * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        
    '执行查询
    Call OpenRecord(rs, strSQL, frmParent.Caption)
    If rs.BOF Then
        If blnPrompt Then MsgBox "没有找到相匹配的结果！", , gstrSysName
        Exit Function                            '没有结果，直接返回
    End If
            
    If rs.RecordCount = 1 And blnFilter Then GoTo over                    '因为是输入查找，如果只有一条，则直接返回
    'If frmSelectList.ShowSelect(frmParent, rs, strLvw, lngX, lngY, lngCX, lngCY, strSavePath, strDescrible, , , objTxt.Height) Then GoTo Over
    
    If frmSelectDialog.ShowSelect(frmParent, 2, rs, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False, False) Then GoTo over
    
    Exit Function
    
over:
    
    Set rsResult = rs
    
    ShowTxtFilterDialog = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then Resume
End Function

Public Function OpenRecord(rsTmp As ADODB.Recordset, strSQL As String, ByVal strTitle As String, _
    Optional CursorType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    
    rsTmp.Open strSQL, gcnOracle, CursorType, LockType
        
    Set OpenRecord = rsTmp
End Function

Public Sub ExecuteProc(ByVal strSQL As String, ByVal strCaption As String)
'功能：执行SQL语句
    
    If UCase(Left(strSQL, 3)) = "ZL_" Then
        gcnOracle.Execute strSQL, , adCmdStoredProc
    Else
        gcnOracle.Execute strSQL
    End If
    
End Sub

Public Function CloseChildWindows(ByVal frmMain As Object, ByVal FrmSon As Object) As Boolean
    '功能:关闭所有子窗口
    
    Dim frmThis As Form
    
    On Error Resume Next
    
    CloseChildWindows = True
    
    For Each frmThis In Forms
        If frmThis.Caption <> frmMain.Caption And frmThis.Caption <> FrmSon.Caption Then Unload frmThis
    Next
    
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '-----------------------------------------------------------------------------------------
    '功能:获取特殊时间
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    End Select
    
End Function

Public Function LoadGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal objIls As Object, Optional ByVal blnCharge As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    Dim strField As String
    Dim strIcon As String
    Dim blnField As Boolean
    Dim blnForeColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("前景色").Name = "前景色")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            strField = objMsf.Cell(flexcpData, 0, lngLoop)
            
            If Trim(strField) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If Left(strField, 1) = "[" Then
                
                    strField = Mid(strField, 2, Len(strField) - 2)
                    strIcon = ""
                    
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                    If Not (objIls Is Nothing) Then
                        strIcon = NVL(rsData(strField))
                        If strIcon <> "" Then
                            Set objMsf.Cell(flexcpPicture, lngRow, lngLoop) = objIls.ListImages(strIcon).Picture
                        End If
                    End If
                    
                    objMsf.Cell(flexcpData, lngRow, lngLoop) = strIcon
                    objMsf.TextMatrix(lngRow, lngLoop) = strIcon
                Else
                
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                     If strMask <> "" Then
                        objMsf.TextMatrix(lngRow, lngLoop) = Format(NVL(rsData(strField)), strMask)
                    Else
                        objMsf.TextMatrix(lngRow, lngLoop) = NVL(rsData(strField))
                    End If
                
                    objMsf.Cell(flexcpData, lngRow, lngLoop, lngRow, lngLoop) = objMsf.TextMatrix(lngRow, lngLoop)
                End If
                
            End If
NextCol:
            '下一列
        Next
        
pointNext:
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("前景色").Value)
        
        rsData.MoveNext
    Loop
    
    LoadGrid = True
    Exit Function
    
errHand:
'
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Public Function SQLInit(ByRef rs As ADODB.Recordset) As Boolean
    
    Set rs = New ADODB.Recordset
    
    With rs
        .Fields.Append "SQL", adVarChar, 4000
        .Open
    End With
    
    SQLInit = True
    
End Function

Public Function SQLAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String) As Boolean
    
    With rs
        .AddNew
        .Fields("SQL").Value = strSQL
    End With
    
    SQLAdd = True
    
End Function

Public Sub ResetVsf(objVsf As Object)
    '
    objVsf.Rows = 2
    objVsf.RowData(1) = ""
    objVsf.Cell(flexcpText, 1, 0, 1, objVsf.Cols - 1) = ""
    
    On Error Resume Next
    
    Set objVsf.Cell(flexcpPicture, 1, 0, 1, objVsf.Cols - 1) = Nothing
End Sub

